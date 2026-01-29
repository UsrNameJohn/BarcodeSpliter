def parse_barcode(s: str):
    s = s.strip()
    if not s:
        return None

    pos1 = s.find(")")
    pos2 = s.find("(", pos1 + 14)
    pos3 = s.find(")", pos2)

    if pos1 == -1 or pos2 == -1 or pos3 == -1:
        return {
            "internal_reference": "",
            "fixed_identifier": "",
            "variable_identifier_data": "",
            "amount": "",
            "amount_data": "",
            "readable_number": "Ongeldig formaat"
        }

    return {
        "internal_reference": s[1:pos1],
        "fixed_identifier": s[pos1 + 1 : pos1 + 14],
        "variable_identifier_data": s[pos1 + 14 : pos2],
        "amount": s[pos2 + 1 : pos3],
        "amount_data": s[pos3 + 1 :],
        "readable_number": s
    }

