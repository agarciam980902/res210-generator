from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import fitz
import shutil
import os
import zipfile
import pandas as pd
import io
import tempfile
 
app = Flask(__name__)
CORS(app)  # Allow Lovable frontend to call this API
 
 
def fill_pdf(template_path, output_path, data):
    shutil.copy(template_path, output_path)
    doc = fitz.open(output_path)
    field_map = {
        "Fp":                            str(data.get("Fp", "1")),
        "DCAL":                          str(round(float(data["DCAL"]), 1)),
        "S":                             str(round(float(data["S"]), 2)),
        "DACs":                          str(data.get("DACs", "")),
        "FRi":                           str(round(float(data["FRi"]), 2)),
        "FR":                            str(round(float(data["FRj"]), 2)),
        "AEtotal":                       str(round(float(data["AEtotal"]))),
        "Di":                            str(data.get("Di", "1")),
        "Fecha inicio actuación":        data.get("fecha_inicio", ""),
        "Fecha fin actuación":           data.get("fecha_fin", ""),
        "Representante del solicitante": data.get("representante", ""),
        "NIFNIE":                        data.get("nif", ""),
    }
    for page in doc:
        for widget in page.widgets():
            if widget.field_name in field_map:
                widget.field_value = field_map[widget.field_name]
                widget.update()
    doc.save(output_path, incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP)
    doc.close()
 
 
def safe_name(s):
    return s.replace(" ", "_").replace(",", "").replace("/", "-")
 
 
def detect_header_row(df_raw):
    """Find the row index that contains 'Direccion' as a column header."""
    for i, row in df_raw.iterrows():
        if any(str(v).strip().lower() == "direccion" for v in row):
            return i
    raise ValueError("Could not find 'Direccion' header row in Excel file.")
 
 
def read_buildings(excel_file, sheet_name):
    """
    Reads any Excel with this column structure:
    Direccion | Demanda en calefaccion | Superficie construida | Superficie habitable | % conectadas | Fri | Frj | AETotal
    Auto-detects which row the headers are on.
    """
    df_raw = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
    header_row = detect_header_row(df_raw)
 
    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=header_row)
    df = df.dropna(subset=["Direccion"])
    # Drop totals/summary rows (no real address)
    df = df[df["Direccion"].astype(str).str.strip().str.len() > 2]
 
    buildings = []
    for _, row in df.iterrows():
        try:
            buildings.append({
                "direccion": str(row["Direccion"]).strip(),
                "DCAL":      float(row["Demanda en calefaccion"]),
                "S":         float(row["Superficie habitable (m2)"]),
                "FRi":       float(row["Fri"]),
                "FRj":       float(row["Frj"]),
                "AEtotal":   float(row["AETotal"]),
            })
        except (ValueError, KeyError):
            continue  # skip malformed rows silently
 
    return buildings
 
 
@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})
 
 
@app.route("/sheets", methods=["POST"])
def get_sheets():
    """Return list of sheet names in the uploaded Excel."""
    excel = request.files.get("excel")
    if not excel:
        return jsonify({"error": "No Excel file uploaded"}), 400
    xl = pd.ExcelFile(excel)
    return jsonify({"sheets": xl.sheet_names})
 
 
@app.route("/preview", methods=["POST"])
def preview():
    """
    Upload an Excel file, get back a list of buildings found.
    Form fields:
      - excel: the .xlsx file
      - sheet: sheet name (optional, auto-detected if omitted)
    """
    excel = request.files.get("excel")
    if not excel:
        return jsonify({"error": "No Excel file uploaded"}), 400
 
    sheet = request.form.get("sheet")
    if not sheet:
        xl = pd.ExcelFile(excel)
        sheet = xl.sheet_names[0]  # default to first sheet
        excel.seek(0)
 
    try:
        buildings = read_buildings(excel, sheet)
    except Exception as e:
        return jsonify({"error": str(e)}), 400
 
    return jsonify({
        "sheet": sheet,
        "count": len(buildings),
        "buildings": buildings
    })
 
 
@app.route("/generate", methods=["POST"])
def generate():
    """
    Upload Excel + PDF template + metadata, get back a ZIP of filled PDFs.
    Form fields:
      - excel:        the .xlsx file
      - template:     the blank RES210 PDF template
      - sheet:        sheet name (optional)
      - representante, nif, fecha_inicio, fecha_fin, di (optional overrides)
    """
    excel    = request.files.get("excel")
    template = request.files.get("template")
 
    if not excel or not template:
        return jsonify({"error": "Both 'excel' and 'template' files are required"}), 400
 
    sheet = request.form.get("sheet")
    if not sheet:
        xl = pd.ExcelFile(excel)
        sheet = xl.sheet_names[0]
        excel.seek(0)
 
    # Metadata — passed from the UI form, falls back to empty if not provided
    meta = {
        "representante": request.form.get("representante", ""),
        "nif":           request.form.get("nif", ""),
        "fecha_inicio":  request.form.get("fecha_inicio", ""),
        "fecha_fin":     request.form.get("fecha_fin", ""),
        "Di":            request.form.get("di", "1"),
        "Fp":            request.form.get("fp", "1"),
    }
 
    try:
        buildings = read_buildings(excel, sheet)
    except Exception as e:
        return jsonify({"error": str(e)}), 400
 
    with tempfile.TemporaryDirectory() as tmp:
        template_path = os.path.join(tmp, "template.pdf")
        template.save(template_path)
 
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for b in buildings:
                data = {**b, **meta}  # merge building data with metadata
                folder   = safe_name(b["direccion"])
                out_path = os.path.join(tmp, f"{folder}.pdf")
                fill_pdf(template_path, out_path, data)
                zf.write(out_path, arcname=f"{folder}/E1-3-1_Ficha_RES210_{folder}.pdf")
 
        zip_buffer.seek(0)
        return send_file(
            zip_buffer,
            mimetype="application/zip",
            as_attachment=True,
            download_name="dossiers_RES210.zip"
        )
 
 
if __name__ == "__main__":
    app.run(debug=True, port=5000)
