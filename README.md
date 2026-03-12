RES210 Dossier Generator
API that takes an Excel file with building data and a blank RES210 PDF template,
and returns a ZIP of filled PDFs — one per building.
Endpoints
GET /health
Check the server is running.
POST /sheets
Upload an Excel file, get back the list of sheet names.

excel — the .xlsx file

POST /preview
Upload an Excel file, get back the list of buildings found.

excel — the .xlsx file
sheet — sheet name (optional, defaults to first sheet)

POST /generate
Upload Excel + PDF template + metadata, get back a ZIP of filled PDFs.

excel — the .xlsx file
template — the blank RES210 PDF template
sheet — sheet name (optional)
representante — legal representative name
nif — NIF/NIE number
fecha_inicio — start date (dd/mm/yyyy)
fecha_fin — end date (dd/mm/yyyy)
di — indicative duration in years (default: 1)
fp — weighting factor (default: 1)

Excel column structure required
| Direccion | Demanda en calefaccion | Superficie construida (m2) | Superficie habitable (m2) | % de viviendas conectadas | Fri | Frj | AETotal |
Local development
pip install -r requirements.txt
python app.py
Server runs at http://localhost:5000
Deploy to Railway

Push this repo to GitHub
Go to railway.app → New Project → Deploy from GitHub
Railway auto-detects Python and deploys using the Procfile
Copy the public URL and use it in your Lovable frontend
