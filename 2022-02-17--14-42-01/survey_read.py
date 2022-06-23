
import xlrd


class Param:
    def __init__(self,name,value):
        self.value = value
        self.name = name
    def __str__(self):
        return f"{self.name} {self.value}"

# output file name
constant_file_name = 'BigData_tools.xlsx'

workbook = xlrd.open_workbook(constant_file_name)
worksheet = workbook.sheet_by_name("Sheet1")

constants = []
constants_d = {}
for row in range(1,worksheet.nrows):
    name = int(worksheet.cell_value(row,0))
    value = int(worksheet.cell_value(row,1)) 
    constants.append(Param(name,value))
    constants_d[name] = value

sort = sorted(constants_d.items(),key=lambda x: x[1],reverse=True)

for e in sort[:20]:
    print(e[0])