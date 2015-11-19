from openpyxl import load_workbook


def xmsx2json(path):
    wb = load_workbook(path)
    ws = wb.worksheets[0]
    rows = ws.rows
    header_row = rows[0]
    content_rows = rows[1:]
    data_dict = []

    for row in content_rows:
        pos = 0
        row_dict = {}
        for header in header_row:
            row_dict[header.value] = row[pos].value
            pos += 1
        data_dict.append(row_dict)
    return data_dict
