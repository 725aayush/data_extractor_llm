"""Build the distinct metadata and tabular representations of an extraction."""

from copy import deepcopy
import re


def _field_name(value: object) -> str:
    """Make document column labels safe, stable field names for Excel and JSON."""
    normalized = re.sub(r"[^a-z0-9]+", "_", str(value).strip().lower())
    return normalized.strip("_") or "value"


def metadata_data(data: dict) -> dict:
    """Keep document context while removing extracted table cell values."""
    result = deepcopy(data)
    result["tables"] = [
        {key: value for key, value in table.items() if key != "rows"} | {"row_count": len(table.get("rows", []))}
        for table in data.get("tables", [])
        if isinstance(table, dict)
    ]
    return result


def tabular_data(data: dict) -> dict:
    """Create one flat record set: invoice, parties, totals, and each line item.

    Each product is a row. Invoice, seller, buyer, and total values are repeated
    on that row so the result can be consumed as one Excel sheet or SAP payload.
    """
    base: dict[str, object] = {}
    for section in ("document", "metadata", "seller", "buyer", "totals"):
        values = data.get(section, {})
        if not isinstance(values, dict):
            continue
        for key, value in values.items():
            if isinstance(value, (str, int, float, bool)) or value is None:
                base[f"{section}_{_field_name(key)}"] = value

    records: list[dict] = []
    for index, table in enumerate(data.get("tables", []), start=1):
        if not isinstance(table, dict):
            continue
        for line_number, row in enumerate(table.get("rows", []), start=1):
            if not isinstance(row, dict) or not isinstance(row.get("cells"), dict):
                continue
            record = {
                **base,
                "source_table_id": table.get("table_id") or f"table_{index}",
                "source_table_title": table.get("title") or f"Table {index}",
                "line_number": line_number,
            }
            record.update({f"line_{_field_name(key)}": value for key, value in row["cells"].items()})
            records.append(record)

    if not records:
        records.append(base)
    columns = list(dict.fromkeys(key for record in records for key in record))
    return {"columns": columns, "records": records}
