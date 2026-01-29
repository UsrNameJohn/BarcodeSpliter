from fastapi import FastAPI, Request, Form
from fastapi.responses import StreamingResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from openpyxl import Workbook
from io import BytesIO
from datetime import datetime
from app.parser import parse_barcode  # jouw bestaande logica

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

@app.get("/", response_class=JSONResponse)
def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/preview")
def preview_barcodes(barcodes: str = Form(...)):
    """
    Voor realtime preview: split barcodes en markeer ongeldig
    """
    lines = [l.strip() for l in barcodes.splitlines() if l.strip()]
    preview = []

    for line in lines:
        parsed = parse_barcode(line)
        if parsed and isinstance(parsed, dict):
            parsed["valid"] = True
            preview.append(parsed)
        else:
            preview.append({
                "internal_reference": "",
                "fixed_identifier": "",
                "variable_identifier_data": "",
                "amount": "",
                "amount_data": "",
                "readable_number": line,
                "valid": False
            })

    return JSONResponse(preview)

@app.post("/export")
def export_to_excel(barcodes: str = Form(...)):
    lines = [l.strip() for l in barcodes.splitlines() if l.strip()]
    records = []

    for line in lines:
        parsed = parse_barcode(line)
        if parsed and isinstance(parsed, dict):
            records.append(parsed)

    if not records:
        return JSONResponse({"error": "Geen geldige barcodes gevonden. Controleer het formaat."}, status_code=400)

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

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"Export_{timestamp}_Records_{len(records)}.xlsx"

    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )
