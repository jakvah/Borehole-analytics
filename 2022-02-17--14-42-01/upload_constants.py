
import xlrd
import xlsxwriter
import requests
from bs4 import BeautifulSoup as bs
import json

toolSN = '3331'

class Param:
    def __init__(self,name,value):
        self.value = value
        self.name = name
    def __str__(self):
        return f"{self.name} {self.value}"

# output file name
constant_file_name = toolSN + ' calibration constants MAB edit.xlsx'

workbook = xlrd.open_workbook(constant_file_name)
worksheet = workbook.sheet_by_name("Ark1")

constants = []
constants_d = {}
for row in range(0,worksheet.nrows):
    name = worksheet.cell_value(row,0)
    value = worksheet.cell_value(row,1) 
    constants.append(Param(name,value))
    constants_d[name] = value


# Request
TOKEN = "eyJhbGciOiJodHRwOi8vd3d3LnczLm9yZy8yMDAxLzA0L3htbGRzaWctbW9yZSNobWFjLXNoYTI1NiIsInR5cCI6IkpXVCJ9.eyJVc2VyIjoiYXV0aDB8NjIxZjYwMDIyOGQ4NmQwMDZhNjY4ZTc2IiwiS2V5IjoiMjU5YzYxNGYtOThlMS1lY2QxLTYzZjEtNGE5NmJjNWY1OWYxIn0.doupf7OS5vEEQ_VuDT4sQqZvbxDDRetDjZ9lB18s4Jc"
HEADERS = {"Authorization": f"Bearer {TOKEN}"}

#url = f"https://ptest-api.devi.cloud/api/V2/Search/Rows/survey?database=bigdata&lengthUnit=m&tempUnit=C&filter=tool.config.serialnumber%20in%20({str(toolSN)}%2C{str(toolSN)})&limit=1&sort=created%20desc&attributes=tool.config.Constants"
url = "https://ptest-api.devi.cloud/api/V2/Search/Rows/survey?database=bigdata&lengthUnit=m&tempUnit=C&filter=tool.config.serialnumber%20in%20(6217%2C2510)&limit=1&sort=created%20desc&attributes=tool.config.instanceid%2Ctool.config.Constants"
r = requests.get(url, headers=HEADERS)
j = r.json()

html = j[0]["tool.config.constants"]["value"]
instance_id = j[0]["instanceId"]["value"]
soup = bs(html,"html.parser",from_encoding="utf-8")


# Prepare new string
new_string = "<CodedConstants>" + str(soup.findAll("itemcount")[0])
num_values = int(soup.findAll("itemcount")[0].text.strip())

for i in range(num_values):
    i += 1
    name = str(soup.findAll(f"item{i}")[0].text.strip().split(":")[0]) + ":"
    if name in constants_d:
        value = str(soup.findAll(f"item{i}")[0].text.strip().split(":")[1])
        sub_string = f"<item{i}>{name}{value}</item{i}>"
        new_string += sub_string
    else:
        item = str(soup.findAll(f"item{i}")[0])
        new_string += item

new_string += "</CodedConstants>"

# Upload
update_object = {
    "id": instance_id,
    "attributes": {
        "constants": new_string
    }
}

json_object = json.dumps([update_object])
print(json_object)

post_url = "https://ptest-api.devi.cloud/api/Instance?database=bigdata"
#r = requests.put(post_url,headers={"Authorization": f"Bearer {TOKEN}", 'Content-Type': 'application/json'},data = json_object)
#print(r.status_code)