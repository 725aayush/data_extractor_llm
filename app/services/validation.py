import re


def _number(value: object) -> float | None:
    if value is None:
        return None
    text = re.sub(r"[₹$€£,%\s]", "", str(value).strip())
    if text.startswith("(") and text.endswith(")"):
        text = f"-{text[1:-1]}"
    try:
        return float(text)
    except ValueError:
        return None


def validate_invoice(data: dict) -> dict:
    issues: list[object] = []
    for name in ("document", "metadata", "seller", "buyer", "ship_to", "shipping", "tables", "taxes", "totals", "pages"):
        if name not in data:
            issues.append(f"Missing section: {name}")
    for name in ("tables", "taxes", "pages"):
        if not isinstance(data.get(name, []), list):
            issues.append(f"'{name}' must be a list.")

    math_issues = []
    for table in data.get("tables", []):
        if not isinstance(table, dict):
            issues.append("A table is not an object.")
            continue
        if not isinstance(table.get("columns", []), list) or not isinstance(table.get("rows", []), list):
            issues.append(f"Table {table.get('table_id', 'unknown')} has invalid columns or rows.")
            continue
        table_issues = []
        for index, row in enumerate(table["rows"], start=1):
            cells = row.get("cells", {}) if isinstance(row, dict) else {}
            if not isinstance(cells, dict):
                continue
            values = {str(key).lower(): _number(value) for key, value in cells.items()}
            quantity = next((v for k, v in values.items() if k == "qty" or "quantity" in k), None)
            rate = next((v for k, v in values.items() if k == "rate" or "unit price" in k or k == "price"), None)
            amount = next((v for k, v in values.items() if k in {"amount", "total", "line total", "line amount"}), None)
            if None not in (quantity, rate, amount) and abs(quantity * rate - amount) > max(1.0, abs(quantity * rate) * .02):
                table_issues.append({"row": index, "expected_amount": round(quantity * rate, 2), "extracted_amount": amount})
        if table_issues:
            math_issues.append({"table": table.get("table_id", "unknown"), "issues": table_issues})
    if math_issues:
        issues.append({"type": "mathematical_validation", "details": math_issues})
    data["validation"] = {"needs_review": bool(issues), "issues": issues}
    return data
