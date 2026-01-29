@app.post("/export")
def export_to_excel(barcodes: str = Form(...)):
    lines = [l.strip() for l in barcodes.splitlines() if l.strip()]
    records = []

    for line in lines:
        parsed = parse_barcode(line)

        if not parsed or not isinstance(parsed, dict):
            continue  # ongeldige barcode → overslaan

        records.append(parsed)

    # 🔒 BELANGRIJK: stop als niets geldig is
    if not records:
        return {
            "error": "Geen geldige barcodes gevonden. Controleer het formaat."
        }

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

    filename = f"Export_{date.today()}_{len(records)}.xlsx"

    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"'
        }
    )
