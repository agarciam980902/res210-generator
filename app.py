from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import fitz
import shutil
import os
import zipfile
import pandas as pd
import io
import tempfile
import unicodedata
import re
 
app = Flask(__name__)
CORS(app)
 
 
# ── HELPERS ──────────────────────────────────────────────────────────────────
 
def safe_name(s):
    return s.replace(" ", "_").replace(",", "").replace("/", "-")
 
 
def normalize(s):
    """Lowercase, remove accents and punctuation — used for fuzzy filename matching."""
    s = s.lower().strip()
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")  # strip accents
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s
 
 
def match_cee_files(direccion, cee_files):
    """
    Given a building address and a dict of {filename: FileStorage},
    return the CEE files whose filename contains the address keywords.
    Returns (cee_antes, cee_despues) — either can be None if not found.
    
    Matching logic:
    - Normalize both the address and filenames
    - A file matches if it contains the main address tokens (street + number)
    - 'antes'/'anterior'/'prior' → CEE before
    - 'posterior'/'despues'/'after'/'post' → CEE after
    - If only one match found for address, assign based on keywords
    """
    norm_addr = normalize(direccion)
    addr_tokens = norm_addr.split()
 
    matched = []
    for fname, fobj in cee_files.items():
        norm_fname = normalize(fname)
        # Match if all main address tokens appear in filename
        if all(token in norm_fname for token in addr_tokens):
            matched.append((fname, fobj, norm_fname))
 
    cee_antes   = None
    cee_despues = None
 
    for fname, fobj, norm_fname in matched:
        if any(k in norm_fname for k in ["anterior", "antes", "prior", "previo"]):
            cee_antes = (fname, fobj)
        elif any(k in norm_fname for k in ["posterior", "despues", "after", "post"]):
            cee_despues = (fname, fobj)
 
    # Fallback: if only 2 matched and neither had keywords, assign by order
    if not cee_antes and not cee_despues and len(matched) == 2:
        cee_antes   = (matched[0][0], matched[0][1])
        cee_despues = (matched[1][0], matched[1][1])
 
    return cee_antes, cee_despues
 
 
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
 
 
def detect_header_row(df_raw):
    for i, row in df_raw.iterrows():
        if any(str(v).strip().lower() == "direccion" for v in row):
            return i
    raise ValueError("Could not find 'Direccion' header row in Excel file.")
 
 
def read_buildings(excel_file, sheet_name):
    df_raw = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
    header_row = detect_header_row(df_raw)
    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=header_row)
    df = df.dropna(subset=["Direccion"])
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
            continue
    return buildings
 
 
# ── ROUTES ───────────────────────────────────────────────────────────────────
 
@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})
 
 
@app.route("/sheets", methods=["POST"])
def get_sheets():
    excel = request.files.get("excel")
    if not excel:
        return jsonify({"error": "No Excel file uploaded"}), 400
    xl = pd.ExcelFile(excel)
    return jsonify({"sheets": xl.sheet_names})
 
 
@app.route("/preview", methods=["POST"])
def preview():
    excel = request.files.get("excel")
    if not excel:
        return jsonify({"error": "No Excel file uploaded"}), 400
    sheet = request.form.get("sheet")
    if not sheet:
        xl = pd.ExcelFile(excel)
        sheet = xl.sheet_names[0]
        excel.seek(0)
    try:
        buildings = read_buildings(excel, sheet)
    except Exception as e:
        return jsonify({"error": str(e)}), 400
    return jsonify({"sheet": sheet, "count": len(buildings), "buildings": buildings})
 
 
@app.route("/generate", methods=["POST"])
def generate():
    """
    Form fields:
      - excel           → .xlsx with building data
      - template        → blank RES210 PDF
      - static_docs     → multiple files, repeated field (10+ files, same in all dossiers)
      - cee_files       → multiple files, repeated field (named by address, auto-matched)
      - sheet, representante, nif, fecha_inicio, fecha_fin, di, fp
    """
    excel    = request.files.get("excel")
    template = request.files.get("template")
 
    if not excel or not template:
        return jsonify({"error": "Both 'excel' and 'template' are required"}), 400
 
    # Collect static docs — same in every dossier
    static_docs = request.files.getlist("static_docs")
 
    # Collect CEE files — auto-matched per building by filename
    cee_file_list = request.files.getlist("cee_files")
    cee_files = {f.filename: f for f in cee_file_list}  # {filename: FileStorage}
 
    sheet = request.form.get("sheet")
    if not sheet:
        xl = pd.ExcelFile(excel)
        sheet = xl.sheet_names[0]
        excel.seek(0)
 
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
 
        # Save static docs to tmp once — reused across all buildings
        static_paths = []
        for f in static_docs:
            p = os.path.join(tmp, "static_" + f.filename)
            f.save(p)
            static_paths.append((f.filename, p))
 
        # Save CEE files to tmp
        cee_paths = {}
        for fname, fobj in cee_files.items():
            p = os.path.join(tmp, "cee_" + fname)
            fobj.seek(0)
            fobj.save(p)
            cee_paths[fname] = p
 
        zip_buffer = io.BytesIO()
        manifest = []  # track what was matched for each building
 
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for b in buildings:
                data   = {**b, **meta}
                folder = safe_name(b["direccion"])
 
                # 1. Generated E1-3-1 PDF
                e131_path = os.path.join(tmp, f"{folder}.pdf")
                fill_pdf(template_path, e131_path, data)
                zf.write(e131_path, arcname=f"{folder}/E1-3-1_Ficha_RES210_{folder}.pdf")
 
                # 2. Static docs — same for every building
                for fname, fpath in static_paths:
                    zf.write(fpath, arcname=f"{folder}/{fname}")
 
                # 3. CEE files — auto-matched by address
                cee_antes, cee_despues = match_cee_files(b["direccion"], 
                    {k: type("F", (), {"filename": k})() for k in cee_paths})
 
                building_manifest = {
                    "direccion": b["direccion"],
                    "cee_antes": None,
                    "cee_despues": None
                }
 
                if cee_antes:
                    fname = cee_antes[0]
                    zf.write(cee_paths[fname], arcname=f"{folder}/E1-3-5_CEE_Anterior_{folder}.pdf")
                    building_manifest["cee_antes"] = fname
 
                if cee_despues:
                    fname = cee_despues[0]
                    zf.write(cee_paths[fname], arcname=f"{folder}/E1-3-5_CEE_Posterior_{folder}.pdf")
                    building_manifest["cee_despues"] = fname
 
                manifest.append(building_manifest)
 
        # Add a manifest.json to the ZIP so user can verify matching
        manifest_json = io.BytesIO()
        import json
        manifest_json.write(json.dumps(manifest, indent=2, ensure_ascii=False).encode())
        manifest_json.seek(0)
 
        with zipfile.ZipFile(zip_buffer, "a") as zf:
            zf.writestr("_matching_report.json", manifest_json.read())
 
        zip_buffer.seek(0)
        return send_file(
            zip_buffer,
            mimetype="application/zip",
            as_attachment=True,
            download_name="dossiers_RES210.zip"
        )
 
 
if __name__ == "__main__":
    app.run(debug=True, port=5000)
