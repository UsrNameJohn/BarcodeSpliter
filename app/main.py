from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles
from openpyxl import Workbook
from io import BytesIO
from datetime import date

from app.parser import parse_barcode

app = FastAPI()

# Static files (voor xlsx.full.min.js)
app.mount("/static", StaticFiles(directory="static"), name="static")

templates = Jinja2Templates(directory="templates")


@app.get("/", response_class=HTMLResponse)
def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/export")
def export_to_excel(barcodes: str = Form(...)):
    lines = barcodes.splitlines()
    records = []

    for line in lines:
        line = line.strip()
        if not line:
            continue
        parsed = parse_barcode(line)
        if parsed:
            records.append(parsed)

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
            r["internal_reference"],
            r["fixed_identifier"],
            r["variable_identifier_data"],
            r["amount"],
            r["amount_data"],
            r["readable_number"]
        ])

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    filename = f"Export_{date.today()}_Records_{len(records)}.xlsx"

    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"'
        }
    )
