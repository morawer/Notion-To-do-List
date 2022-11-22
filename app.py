from openpyxl import load_workbook

class Panel:
    def __init__(self, item, name, qty):
        self.item = item
        self.name = name
        self.qty = qty


wb = load_workbook("Listado paneles.xlsx")

# grab the active worksheet
ws = wb.active

co = 0
line = 0
qty = 0
row_counter = 0

panels_array = []

for row in ws.iter_rows(min_row=2, max_row=ws.max_row):

    if (row[0].value != co or row[1].value != line):

        print(f"{row[0].value}-{row[1].value}")

        co = row[0].value
        line = row[1].value
        qty = qty + 1

        wb_panels = load_workbook("Plantilla Paneles.xlsx")

        ws_panels = wb_panels.active
        ws_panels["B2"].value = f"{co}-{line}"
        row_counter = 0


    ws_panels.cell(row=11 + row_counter, column=2, value=row[2].value)
    ws_panels.cell(row=11 + row_counter, column=4, value=row[3].value)
    ws_panels.cell(row=11 + row_counter, column=10, value=row[4].value)
    
    row_counter = row_counter + 1

    wb_panels.save(f"{row[0].value}-{row[1].value}.xlsx")
    


print(f"TOTAL = {qty}")
