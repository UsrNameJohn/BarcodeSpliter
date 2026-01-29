from fastapi import FastAPI, Request, Form
from fastapi.responses import StreamingResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from openpyxl import Workbook
from io import BytesIO
from datetime import date

# 🔹 Zorg dat je parser.py in dezelfde map app/ staat
from app.parser import parse_barcode

# =========================================
# 1️⃣ FastAPI app aanmaken
# =========================================
app = FastAPI()

# =========================================
# 2️⃣ Static files (xlsx.full.min.js)
# =========================================
app.mount("/static", StaticFiles(directory="static"), name="static")

# =========================================
# 3️⃣ Templates
# =========================================
templates = Jinja2Templates(directory="templates")

# =========================================
# 4️⃣ Routes
# =========================================

@app.get("/", response_class=HTMLResponse)
def index(request: Request):
    """
    Homepagina
    """
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/export")
def export_to_excel(barcodes: str = Form(...)):
    lines = [l.strip() for l in barcodes.splitlines() if l.strip()]
    records = []

    for line in lines:
        parsed = parse_barcode(line)
        if parsed and isinstance(parsed, dict):
            records.append(parsed)

    if not records:
        return {"error": "Geen geldige barcodes gevonden. Controleer het formaat."}

    wb = Workbook()
    ws = wb.active
    ws.title = "Export"

    headers = [
        "Internal reference",
        "Fixed identifier",
        "Variable identifier data",
        "Amount",
        "Amount data",
        "Readable number"
    ]
    ws.append(headers)

    for r in records:
        ws.append([
            r.get("internal_reference", ""),
            r.get("fixed_identifier", ""),
            r.get("variable_identifier_data", ""),
            r.get("amount", ""),
            r.get("amount_data", ""),
            r.get("readable_number", "")
        ])

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    # ✅ Unieke bestandsnaam
    from datetime import datetime
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"Export_{timestamp}_Records_{len(records)}.xlsx"

    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )

