# -*- coding: utf-8 -*-
import os
import yaml
import pyutilib.subprocess.GlobalData

import dash
import dash_auth
import dash_core_components as dcc
import dash_html_components as html

from tabs import *


# set environment
os.chdir(os.path.dirname(os.path.realpath(__file__)))
with open("config.yaml") as f:
    configinfo = yaml.load(f, Loader=yaml.Loader)
pyutilib.subprocess.GlobalData.DEFINE_SIGNAL_HANDLERS_DEFAULT = False


# start dash
app = dash.Dash()
app.title = 'Supply Optimization'
users = configinfo['user']
valid_users = {i: j['password'] for i, j in users.items()}
auth = dash_auth.BasicAuth(app, valid_users)

# dash layout
app.layout = html.Div([
    dcc.Tabs(id="tabs", children=[
        dcc.Tab(label='Optimization', children=tab_opt.tab),
        dcc.Tab(label='Visualization', children=tab_viz.tab),
        dcc.Tab(label='Manual', children=tab_manual.tab)
    ])
])

tab.set_callbacks(app)
tab_opt.set_callbacks(app)
tab_viz.set_callbacks(app)


if __name__ == '__main__':
    app.run_server(debug=configinfo['app']['debug'], host='0.0.0.0', port=8080)
