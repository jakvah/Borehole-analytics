# -*- coding: utf-8 -*- 
from turtle import st
from flask import Flask, request, render_template, Markup
import numpy as np

from borehole.datamanager import get_data as gd

app = Flask(__name__)
NUM_TABS = 2
@app.route("/")
def index():
    navbar_status = [""]*NUM_TABS
    navbar_status[0] = "active"
    return render_template("index.html",
                    navbar_status=navbar_status,
    )   

@app.route("/multiview")
def multiview():
    # For the navbar
    navbar_status = [""]*NUM_TABS
    navbar_status[1] = "active"

    return render_template("multiview.html",navbar_status=navbar_status,
    )                 

@app.route("/multi_load_data", methods=["POST"])
def multi_load_data():

    # For the navbar
    navbar_status = [""]*NUM_TABS
    navbar_status[1] = "active"

    x_axis = request.form["xaxis"]
    y_axis = request.form.getlist("yaxis")
    num_y_params = len(y_axis)
    y_tag_string = ""
    for i in range(num_y_params):
        y_tag_string += y_axis[i]
        if i != num_y_params-1:
            y_tag_string += ","

    url_string = f"multiview_show/{x_axis}/{y_tag_string}/{num_y_params}"


    return render_template("loading_data.html",
                            navbar_status=navbar_status, 
                            url_string = url_string
                            )

@app.route("/multiview_show/<x_axis_tag>/<y_tag_string>/<num_y_params>")
def multiview_show(x_axis_tag,y_tag_string,num_y_params):
    # For the navbar
    navbar_status = [""]*NUM_TABS
    navbar_status[1] = "active"
    

    y_axis_tags = y_tag_string.split(",")
    num_y_params = int(num_y_params)

    row_data = []
    x_axis_data = gd.get_data(x_axis_tag)
    for y_param in y_axis_tags:
        y_axis_data = gd.get_data(y_param)
        row_data.append(gd.pair_datasets(x_axis_data,y_axis_data))

    return render_template("multiview_show.html",
                            navbar_status=navbar_status,
                            row_data = row_data,
                            x_axis_tag = x_axis_tag,
                            y_axis_tags = y_axis_tags,
                            num_y_params = num_y_params
                            )
if __name__ == '__main__':
    # Set debug false if it is ever to be deployed
    app.run(debug=True)