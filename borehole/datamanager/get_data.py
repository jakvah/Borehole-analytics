import requests
import json
import matplotlib.pyplot as plt
import numpy as np
BASIS_ENDPOINT = "https://ptest-api.devi.cloud/api/V2/"
TOKEN = "eyJhbGciOiJodHRwOi8vd3d3LnczLm9yZy8yMDAxLzA0L3htbGRzaWctbW9yZSNobWFjLXNoYTI1NiIsInR5cCI6IkpXVCJ9.eyJVc2VyIjoiYXV0aDB8NjIxZjYwMDIyOGQ4NmQwMDZhNjY4ZTc2IiwiS2V5IjoiMjU5YzYxNGYtOThlMS1lY2QxLTYzZjEtNGE5NmJjNWY1OWYxIn0.doupf7OS5vEEQ_VuDT4sQqZvbxDDRetDjZ9lB18s4Jc"
HEADERS = {"Authorization": f"Bearer {TOKEN}"}

"""
url ="https://ptest-api.devi.cloud/api/V2/Search/Rows/survey?database=Bigdata&lengthUnit=m&tempUnit=C&sort=created%20desc&attributes=calculation.eohresult.*"

r = requests.get(url, headers=HEADERS)
j = r.json()


x = []
y = []
posstd = []
poststd1000 = []
for i in range(len(j)):
    yt = j[i]["calculation.eohresult.StddevElevation"]["value"]
    xt = j[i]["calculation.eohresult.AvgElevation"]["value"]

    post = j[i]["calculation.eohresult.StddevPos"]["value"]
    post100t = j[i]["calculation.eohresult.StddevPosPer1000m"]["value"]

    if yt != None and xt != None:
        y.append(yt)
        x.append(abs(xt))
    if post != None and xt != None:
        posstd.append(post)
    if post100t != None and xt != None:
        poststd1000.append(post100t)
    
print(len(x))

fig,ax = plt.subplots(2,1,figsize=(10,10))
ax[0].scatter(x,y,s=1)
ax[0].set_xlabel("Average elevation [m]")
ax[0].set_ylabel("Standard deviation of elevation [sqrt(m)]")
ax[0].set_title("Standard deviation of elevation vs. average elevation")
ax[1].scatter(x,posstd,s=1)
ax[1].set_xlabel("Average elevation [m]")
ax[1].set_ylabel("Standard deviation of position [m]")
ax[1].set_title("Standard deviation of position vs. average elevation")
plt.show()
"""

def get_data(attribute):
    url =f"https://ptest-api.devi.cloud/api/V2/Search/Rows/survey?database=Bigdata&lengthUnit=m&tempUnit=C&sort=created%20desc&attributes={attribute}"
    r = requests.get(url, headers=HEADERS)
    j = r.json()

    values = []
    for i in range(len(j)):
        values.append(j[i][attribute]["value"])
    
    return values

def pair_datasets(series1,series2):
    assert len(series1) == len(series2)

    final_data = []
    for i in range (len(series1)):
        if series1[i] == None or series2[i] == None:
            continue
        else:
            d = {}
            d["x"] = series1[i]
            d["y"] = series2[i]
            final_data.append(d)
    return final_data


if __name__ == "__main__":
    print(1)
    avge = get_data("calculation.eohresult.avgelevation")
    print(2)
    stddev = get_data("calculation.eohresult.stddevelevation")
    print(3)
    dur = get_data("calculation.resultoverview.surveyduration")
    assert len(avge) == len(stddev)
    assert len(avge) == len(dur)



