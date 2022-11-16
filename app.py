from xlrd import open_workbook
import pandas as pd
from order_page import create_page
from order_page import create_item


wb = open_workbook('Libro1.xlsx')
sheet = wb.sheet_by_index(0)

df = pd.read_excel('Libro1.xlsx', usecols=[2,5,6,7,8])

co = int(sheet.cell_value(1, 0))
line = int(sheet.cell_value(1, 1))

order = f"{co}-{line}"

id_database = create_page(order)

print(">>>" + str(id_database))


for index, row in df.iterrows():
    
    mo = row[0]
    itemNo = str(row[1])
    itemName = str(row[2])
    itemDesc = row[3]
    qty = row[4]
    
    create_item(database=id_database, name=itemName, item=itemNo, desc=itemDesc, mo=mo)

    
 
