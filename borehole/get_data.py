import requests
import json

BASIS_ENDPOINT = "https://ptest-api.devi.cloud/api/V2/"
TOKEN = "eyJhbGciOiJodHRwOi8vd3d3LnczLm9yZy8yMDAxLzA0L3htbGRzaWctbW9yZSNobWFjLXNoYTI1NiIsInR5cCI6IkpXVCJ9.eyJVc2VyIjoiYXV0aDB8NjIxZjYwMDIyOGQ4NmQwMDZhNjY4ZTc2IiwiS2V5IjoiMjU5YzYxNGYtOThlMS1lY2QxLTYzZjEtNGE5NmJjNWY1OWYxIn0.doupf7OS5vEEQ_VuDT4sQqZvbxDDRetDjZ9lB18s4Jc"
headers = {"Authorization": f"Bearer {TOKEN}"}

endpoint = BASIS_ENDPOINT + "/"


# 
# hole.group.project.instanceId == "Project_deT6X"
# hole.group.project.site.name,hole.group.project.name,hole.group.name,hole.name,starttime,session.legacy_name,tool.config.type.name,tool.config.serialnumber,calculation.eohresult.*
# 
url ="https://ptest-api.devi.cloud/api/V2/Search/Rows/survey?database=Bigdata&lengthUnit=m&tempUnit=C&filter=hole.group.project.instanceId%20%3D%3D%20%22Project_deT6X%22&sort=created%20desc&attributes=calculation.eohresult.*"

r = requests.get(url, headers=headers)
j = r.json()
print(len(j))

for i in range(len(j)):
    print(j[i]["calculation.eohresult.StddevPosPer1000m"]["value"])
