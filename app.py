from xlrd import open_workbook
import pandas as pd

wb = open_workbook('Libro1.xlsx')
sheet = wb.sheet_by_index(0)

df = pd.read_excel('Libro1.xlsx', usecols=[2,5,6,7,8])

co = int(sheet.cell_value(1, 0))
line = int(sheet.cell_value(1, 1))

order = f"{co}-{line}"

print(order)
print()


for index, row in df.iterrows():
    
    mo = row[0]
    itemNo = row[1]
    itemName = row[2]
    itemDesc = row[3]
    qty = row[4]
    
    
    print(f"MO: {mo}    No: {itemNo}    Name: {itemName}    Qty: {qty}")
    print("POST")
 
