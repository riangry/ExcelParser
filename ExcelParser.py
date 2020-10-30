from openpyxl import load_workbook


def estimate_parser(file):
    wb = load_workbook(file)
    sheet = wb.active
    data = []
    results = ['Материалы', 'Машины и механизмы', 'ФОТ']
    row_count = sheet.max_row
    for row_number in range(1, row_count):
        cell_a = str(sheet[f'A{row_number}'].value)
        cell_b = str(sheet[f'B{row_number}'].value)
        cell_c = str(sheet[f'C{row_number}'].value)
        cell_d = str(sheet[f'D{row_number}'].value)
        cell_f = str(sheet[f'F{row_number}'].value)
        cell_g = str(sheet[f'G{row_number}'].value)
        if "Раздел" in cell_a:
            data.append(cell_a.strip())
        elif sheet[f'B{row_number}'].value is None:
            if sheet[f'A{row_number}'].value is not None and cell_a.strip() not in results and \
                    "итог" not in cell_a.lower() and "ВСЕГО" not in cell_a:
                data.append(cell_a.strip())
        if 'СРК' in cell_b:
            data.append([cell_b, cell_c])
        if "Цена по прайсу" in cell_b:
            data.append([cell_b, cell_c, cell_d, cell_f, cell_g])
        if "Механизмы" in cell_b:
            data.append([cell_b, cell_c, cell_d,cell_f, cell_g])
    return data


section_ids = []
stage_ids = []
srk_ids = []
material_ids = []
mechanic_ids = []
data = estimate_parser('2.xlsx')
for item in data:
    print(item)
    if 'Раздел' in item:
        section_ids.append(data[data.index(item)])
    if isinstance(item, str) and 'Раздел' not in item:
        stage_ids.append(data[data.index(item)])
    if len(item) == 2:
        srk_ids.append(data[data.index(item)])
    if len(item) == 5 and 'Цена по прайсу' in item:
        material_ids.append(data[data.index(item)])
    if len(item) == 5 and 'Механизмы' in item:
        mechanic_ids.append(data[data.index(item)])
n = 0
for item in data:

