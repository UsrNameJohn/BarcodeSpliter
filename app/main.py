from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.templating import Jinja2Templates
from openpyxl import Workbook
from io import BytesIO
from datetime import date

from parser import parse_barcode

app = FastAPI()

# Templates (HTML)
templates = Jinja2Templates(directory="templates")


# ==============================
# STARTPAGINA
# ==============================
@app.get("/", response_class=HTMLResponse)
def index(request: Request):
    return templates.TemplateResponse(
        "index.html",
        {"request": request}
    )


# ==============================
# EXPORT NAAR EXCEL
# ==============================
@app.post("/export")
def export_to_excel(barcodes: str = Form(...)):

    lines = barcodes.splitlines()
    records = []

    # Barcodes parsen
    for line in lines:
        parsed = parse_barcode(line)
        if parsed:
            records.append(parsed)

    # Excel workbook aanmaken
    wb = Workbook()
    ws = wb.active
    ws.title = "Export"

    # Headers (exact zoals in Excel)
    headers = [
        "Internal reference",
        "Fixed identifier",
        "Variable identifier data",
        "Amount",
        "Amount data",
        "Readable number"
    ]
    ws.append(headers)

    # Data toevoegen
    for r in records:
        ws.append([
            r["internal_reference"],
            r["fixed_identifier"],
            r["variable_identifier_data"],
            r["amount"],
            r["amount_data"],
            r["readable_number"]
        ])

    # Workbook naar geheugen schrijven
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    # Bestandsnaam (zonder waarde, alleen datum + records)
    filename = f"Export_{date.today()}_Records_{len(records)}.xlsx"

    # Bestand aanbieden als download
    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f"attachment; filename={filename}"
        }
    )

