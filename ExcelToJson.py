import xlrd 
from collections import OrderedDict
import  json
# Open the workbook and select the first worksheet
wb = xlrd.open_workbook("D:ExcelPoc.xlsx")
#wb = xlrd.open_workbook('ExcelPoc.xls')
sh = wb.sheet_by_index(0)
# List to hold dictionaries
cars_list = []
# Iterate through each row in worksheet and fetch values into dict
for rownum in range(1, sh.nrows):
    cars = OrderedDict()
    row_values = sh.row_values(rownum)
    cars[row_values[0]] = row_values[1]
   
   
    cars_list.append(cars)
# Serialize the list of dicts to JSON
j = json.dumps(cars_list)
# Write to file
with open('D:\\ExcelToJson.txt', 'w') as f:
    f.write(j)
