import re

from openpyxl import load_workbook

# accepted header names (compared case-insensitively, spaces/dots/underscores ignored)
PO_HEADERS = {"PO", "PONUMBER", "PONR", "PURCHASEORDER", "ORDER", "ORDERNO", "UZSAKYMAS"}
LINE_HEADERS = {"LINE", "LN", "LINENO", "LINENUMBER", "LINENR", "EILUTE"}
MAX_HEADER_ROWS = 20  # look for the header row only near the top of a sheet


def _header_key(value):
    text = str(value or "").upper()
    text = text.translate(str.maketrans("ŲŽĖĮŠĄČĘŪ", "UZEISACEU"))
    return re.sub(r"[\s._#-]+", "", text)


def _find_columns(row):
    keys = [_header_key(cell) for cell in row]
    po_column = next((i for i, k in enumerate(keys) if k in PO_HEADERS), None)
    line_column = next((i for i, k in enumerate(keys) if k in LINE_HEADERS), None)
    if po_column is None or line_column is None:
        return None
    return po_column, line_column


def _clean_po(value):
    if value is None:
        return None
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    po = re.sub(r"\s+", "", str(value)).upper()
    if not po:
        return None
    return po if po.startswith("PO") else "PO" + po  # 260531 -> PO260531


def _clean_line(value):
    if value is None or str(value).strip() == "":
        return None
    try:
        number = float(value)
        return int(number) if number.is_integer() else number
    except (TypeError, ValueError):
        return str(value).strip()


def parse_xlsx(path):
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        for sheet in workbook.worksheets:  # header may be on any sheet, not only the first
            rows = sheet.iter_rows(values_only=True)
            columns = None
            for _, row in zip(range(MAX_HEADER_ROWS), rows):
                columns = _find_columns(row)
                if columns:
                    break
            if not columns:
                continue

            po_column, line_column = columns
            orders, seen = [], set()
            for row in rows:
                po = _clean_po(row[po_column] if po_column < len(row) else None)
                line = _clean_line(row[line_column] if line_column < len(row) else None)
                if po is None:
                    continue  # skip empty rows instead of stopping at the first one
                key = (po, line)
                if key in seen:
                    continue  # same PO + LINE twice in the export -> show once
                seen.add(key)
                orders.append({"PO": po, "LINE": line})
            return orders

        raise ValueError("PO and LINE headers not found.")
    finally:
        workbook.close()
