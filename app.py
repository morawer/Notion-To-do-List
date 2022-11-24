from openpyxl import load_workbook
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import Alignment

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
        
    ws_panels.merge_cells(start_row=11 + row_counter,
                          start_column=2, end_row=11 + row_counter, end_column=3)
    
    ws_panels.merge_cells(start_row=11 + row_counter,
                          start_column=4, end_row=11 + row_counter, end_column=9)
    
    ws_panels.merge_cells(start_row=11 + row_counter,
                          start_column=10, end_row=11 + row_counter, end_column=11)

    thin_border = Border(left=Side(style='thin'),
                         right=Side(style='thin'),
                         top=Side(style='thin'),
                         bottom=Side(style='thin'))
    
    cell_alignment = Alignment(horizontal="center", vertical="center")
    
    
    
    ws_panels.cell(row=11 + row_counter, column=2, 
                               value=row[2].value).alignment=cell_alignment
    ws_panels.cell(row=11 + row_counter, column=2).border = thin_border
    ws_panels.cell(row=11 + row_counter, column=3).border = thin_border

    
    ws_panels.cell(row=11 + row_counter, column=4,
                   value=row[3].value).alignment = cell_alignment
    ws_panels.cell(row=11 + row_counter, column=4).border=thin_border
    ws_panels.cell(row=11 + row_counter, column=5).border = thin_border
    ws_panels.cell(row=11 + row_counter, column=6).border = thin_border
    ws_panels.cell(row=11 + row_counter, column=7).border = thin_border
    ws_panels.cell(row=11 + row_counter, column=8).border = thin_border
    ws_panels.cell(row=11 + row_counter, column=9).border = thin_border

    ws_panels.cell(row=11 + row_counter, column=10,
                   value=row[4].value).alignment = cell_alignment
    ws_panels.cell(row=11 + row_counter, column=10).border = thin_border
    ws_panels.cell(row=11 + row_counter, column=11).border = thin_border

    
    ws_panels.cell(row=11 + row_counter, column=12,
                   value=" ").border=thin_border
        
    row_counter = row_counter + 1

    wb_panels.save(f"{row[0].value}-{row[1].value}-PANELES.xlsx")
    
print(f"TOTAL = {qty}")