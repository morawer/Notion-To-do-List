from openpyxl import load_workbook
wb = load_workbook("Listado paneles.xlsx")

# grab the active worksheet
ws = wb.active

co = ws["A2"].value
line = ws["B3"].value
qty = 0

print(f"{co}-{line}")

for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
    
  if row[0].value != co or row[1].value != line: 
    
    print(f"{row[0].value}-{row[1].value}")
    
    co = row[0].value
    line = row[1].value
    qty = qty + 1
    
    wb_panels = load_workbook("Plantilla Paneles.xlsx")
    
    ws_panels = wb_panels.active
    
    ws_panels["B2"].value = f"{row[0].value}-{row[1].value}"
    
    wb_panels.save(f"{row[0].value}-{row[1].value}.xlsx")


  name = row[3].value
  cant = row[4].value
  print(f"  {name} >>> {cant}")
  
      
    

    
    
    

    
print(f"TOTAL = {qty}")
    






    
 
