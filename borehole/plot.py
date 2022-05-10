import matplotlib.pyplot as plt
import numpy as np
from borehole.datamanager import get_data as gd


def plot_param(param_1,param_2,remove_outliers = True,log = False):    
    param_1_data = gd.get_data(param_1)
    param_2_data = gd.get_data(param_2)

    param_1_data, param_2_data = remove_none(param_1_data,param_2_data)

    if remove_outliers:
        p2_mean = np.mean(param_2_data)
        p2_std = np.std(param_2_data)

        new_p1 = []
        new_p2 = []
        for i,p in enumerate(param_2_data):
            if p > (p2_mean + 3 * p2_std) or p < (p2_mean - 3 * p2_std):
                continue
            else:
                new_p1.append(param_1_data[i])
                new_p2.append(p)
    else:
        new_p1 = param_1_data
        new_p2 = param_2_data

    plt.scatter(new_p1,new_p2)
    plt.title(f'{param_1.split(".")[-1]} vs {param_2.split(".")[-1]}')
    plt.xlabel(param_1)
    plt.ylabel(param_2)
    if log:
        plt.yscale("log")
    plt.show()

def remove_none(l1,l2):
    new_l1 = []
    new_l2 = []
    assert len(l1) == len(l2), "Lengths not the same"
    for i in range(len(l1)):
        if l1[i] == None or l2[i] == None:
            continue
        new_l1.append(l1[i])
        new_l2.append(l2[i])
    return new_l1, new_l2


if __name__ == "__main__":
    PARAM1 = "calculation.resultoverview.surveydistance"
    PARAM2 = "calculation.eohresult.stddevpos"

    plot_param(PARAM1,PARAM2)
