import xlrd 
from collections import OrderedDict
import  json
# Open the workbook and select the first worksheet
wb = xlrd.open_workbook("D:ExcelPoc.xlsx")
#wb = xlrd.open_workbook('ExcelPoc.xls')
sh = wb.sheet_by_index(0)
# List to hold dictionaries

emyList={}
def lower_dict(d):
   new_dict = dict((k.lower(), v) for k, v in d.items())
   return new_dict
# Iterate through each row in worksheet and fetch values into dict
for rownum in range(1, sh.nrows):
   
    row_values = sh.row_values(rownum)

    emyList.setdefault(row_values[0], []).append(row_values[1])
    #cars[row_values[0]] = row_values[1]
   
   
    
# Serialize the list of dicts to JSON





s=json.dumps(lower_dict({x.translate({32: None}): y for x, y in emyList.items()}))
# Write to file
with open('D:\\ExcelToJson.txt', 'w') as f:
    f.write(s)
