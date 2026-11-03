from openpyxl import load_workbook


def parse_xlsx(path):
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        rows = workbook.worksheets[0].iter_rows(values_only=True)
        for row in rows:
            if "PO" in row and "LINE" in row:
                po_column = row.index("PO")
                line_column = row.index("LINE")
                break
        else:
            raise ValueError("PO and LINE headers not found.")

        orders = []
        for row in rows:
            po = row[po_column]
            line = row[line_column]
            if po is None and line is None:
                break
            orders.append({"PO": po, "LINE": line})
        return orders
    finally:
        workbook.close()
