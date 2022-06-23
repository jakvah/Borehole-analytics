import requests
import json
import matplotlib.pyplot as plt
import numpy as np
BASIS_ENDPOINT = "https://ptest-api.devi.cloud/api/V2/"
TOKEN = "eyJhbGciOiJodHRwOi8vd3d3LnczLm9yZy8yMDAxLzA0L3htbGRzaWctbW9yZSNobWFjLXNoYTI1NiIsInR5cCI6IkpXVCJ9.eyJVc2VyIjoiYXV0aDB8NjIxZjYwMDIyOGQ4NmQwMDZhNjY4ZTc2IiwiS2V5IjoiMjU5YzYxNGYtOThlMS1lY2QxLTYzZjEtNGE5NmJjNWY1OWYxIn0.doupf7OS5vEEQ_VuDT4sQqZvbxDDRetDjZ9lB18s4Jc"
HEADERS = {"Authorization": f"Bearer {TOKEN}"}


def get_data(attribute):
    url =f"https://ptest-api.devi.cloud/api/V2/Search/Rows/survey?database=Bigdata&lengthUnit=m&tempUnit=C&sort=created%20desc&attributes={attribute}"
    r = requests.get(url, headers=HEADERS)
    j = r.json()

    values = []
    for i in range(len(j)):
        values.append(j[i][attribute]["value"])
    
    return values

  

def get_unique_legacy_names(toolid=6158,maxlength=250):
    #url = f"https://ptest-api.devi.cloud/api/V2/Search/Rows/survey?database=Bigdata&lengthUnit=m&tempUnit=C&filter=tool.config.serialnumber%20in%20%5B{toolid}%5D&sort=created%20desc&attributes=session.legacy_name"
    url = f"https://ptest-api.devi.cloud/api/V2/Search/Rows/survey?database=Bigdata&lengthUnit=m&tempUnit=C&filter=tool.config.serialnumber%20in%20%5B{toolid}%5D&sort=created%20desc&attributes=session.legacy_name%2Ccalculation.resultoverview.enddepth"
    r = requests.get(url, headers=HEADERS)
    j = r.json()
    names = []


    num_non = 0
    num_big = 0
    for i in range(len(j)):
        d = j[i]["calculation.resultoverview.enddepth"]["value"]
        if d is None:
            num_non +=1
            continue
        if float(d) < maxlength:
            c = j[i]["session.legacy_name"]["value"]
            if c not in names:
                names.append(c)
        else:
            num_big += 1

    print(f"Num none: {num_non}")
    print(f"Num above 250: {num_big}")
    return names


def get_survey_means_pr_meter(tool_id=6158):
    legacy_names = get_unique_legacy_names(tool_id,maxlength=250)
    print(f"Number of surveys under 250m: {len(legacy_names)}")
    azi_diff_matrix = []
    incl_diff_matrix = []

    mean_azis = []
    mean_incls = []

    longest_depth = [0]

    for name in legacy_names: # <- Cut here
        azi_diffs, incl_diffs, depths = get_survey_data(tool_id,name)
        try:
            if depths[-1] > longest_depth[-1]:
                longest_depth = depths
        except IndexError:
            continue
        azi_diff_matrix.append(azi_diffs)
        incl_diff_matrix.append(incl_diffs)

    assert len(azi_diff_matrix) == len(incl_diff_matrix)

    num_rows = len(azi_diff_matrix)
    print(f"Num rows: {num_rows}")

    
    for j in range(len(longest_depth)):
        azi_col = []
        incl_col = []
        for i in range(num_rows):
            try:
                azi_col.append(azi_diff_matrix[i][j])
                incl_col.append(incl_diff_matrix[i][j])
            except IndexError:
                continue
        mean_azis.append(np.mean(azi_col))
        mean_incls.append(np.mean(incl_col))

    return mean_azis, mean_incls, longest_depth

def plot_tool(tool_id=6158):
    azi, incl, depth = get_survey_means_pr_meter()
    fig,ax = plt.subplots(2,1)
    ax[0].scatter(depth,azi)
    ax[0].set_title("Mean difference between azimuth angles for in-run and out-run")
    ax[0].set_xlabel("Meters")
    ax[0].set_ylabel("Degrees")
    ax[1].scatter(depth,incl)
    ax[1].set_title("Mean difference between inclination angles for in-run and out-run")
    ax[1].set_xlabel("Meters")
    ax[1].set_ylabel("Degrees")
    plt.show()



def get_survey_data(tool_id,legacy_name):
    url = f"https://ptest-api.devi.cloud/api/V2/Search/Rows/survey?database=Bigdata&lengthUnit=m&tempUnit=C&filter=tool.config.serialnumber%20in%20%5B{tool_id}%5D%20AND%20session.legacy_name%20%3D%3D%20%22{legacy_name}%22&sort=created%20desc&attributes=session.legacy_name%2CisDirectionUp%2Ccalculation.result.Depth%2Ccalculation.result.azimuth%2Ccalculation.result.inclination"
    r = requests.get(url, headers=HEADERS)
    j = r.json()

    azi_up = []
    azi_down = []
    incl_up = []
    incl_down = []

    depths = []
    for i in range(len(j)):
        up = j[i]["isdirectionup"]["value"]
        depth = float(j[i]["calculation.result.depth"]["value"])
        azi = j[i]["calculation.result.azimuth"]["value"]
        incl = j[i]["calculation.result.inclination"]["value"]
        if up:
            depths.append(depth)
            azi_up.append(azi)
            incl_up.append(incl)
        else:
            azi_down.append(azi)
            incl_down.append(incl)

    azi_diffs = []
    incl_diffs = []
    for i in range(len(azi_up)):
        azi_diff = abs(azi_up[i] - azi_down[i])
        incl_diff = abs(incl_up[i] - incl_down[i])
        azi_diffs.append(azi_diff)
        incl_diffs.append(incl_diff)
    assert len(azi_diffs) == len(incl_diffs)
    assert len(azi_diffs) == len(depths)
    return azi_diffs, incl_diffs, depths
      
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
    plot_tool()

    """
    print(1)
    avge = get_data("calculation.eohresult.avgelevation")
    print(2)
    stddev = get_data("calculation.eohresult.stddevelevation")
    print(3)
    dur = get_data("calculation.resultoverview.surveyduration")
    assert len(avge) == len(stddev)
    assert len(avge) == len(dur)
    """



