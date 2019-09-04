import os

from dash.dependencies import Input, Output
import dash_core_components as dcc
import dash_html_components as html
import plotly.graph_objects as go
from flask import request

from models import visualize
import mod

# get config data
p = mod.PathFile()
configinfo = p.configinfo()

# select supply
refresh_btn = html.Div(children=[
    html.Button(id='refresh', children='Refresh')
],)

# header
header = html.Div(children=[
    html.Div([
        html.H4("Visualize from file"),
        html.P(id='plot-user'),
        html.P(id='plot-filename'),
        html.P(id='plot-datetime'),
        html.P("Download: ", style={'display': 'inline-block',
                                    'padding-right': '5px', 'margin-top': '0px'}),
        html.A("input", href="/download/input",
               style={'display': 'inline-block', 'padding-right': '5px'}),
        html.A("output", href="/download/output"),
    ], style={'width': '20%', 'display': 'inline-block', 'vertical-align': 'top'}),
    html.Div([
        html.H4("Net Contribution"),
        html.P(id='netcon')
    ], style={'width': '15%', 'display': 'inline-block', 'vertical-align': 'top'}),
    html.Div([
        html.H4("Revenue"),
        html.P(id='revenue')
    ], style={'width': '15%', 'display': 'inline-block', 'vertical-align': 'top'}),
    html.Div([
        html.H4("Variable Cost"),
        html.P(id='vc')
    ], style={'width': '15%', 'display': 'inline-block', 'vertical-align': 'top'}),
    html.Div([
        html.H4("Fixed Cost"),
        html.P(id='fc')
    ], style={'width': '15%', 'display': 'inline-block', 'vertical-align': 'top'}),
],)

# graph
graph = html.Div(children=[
    html.Div([
        dcc.Graph(id='supply-utilization'),
    ], style={'width': '50%', 'display': 'inline-block'}),
    html.Div([
        dcc.Graph(id='route-map')
    ], style={'width': '40%', 'display': 'inline-block'})
],)

# select supply
select_supply = html.Div(children=[
    html.H4(id='select-supply')
],)

# data table
data_table = html.Div(children=[
    html.Div([
        html.P("Product"),
        html.Div(id='table-product'),
    ], style={'width': '40%', 'display': 'inline-block'}),
    html.Div([
        html.P("Route"),
        html.Div(id='table-route'),
    ], style={'width': '50%', 'display': 'inline-block'})
],)

# invisible data
user_inv = html.Div(children=[
    html.Div([
        html.H6(id='user', style={'display': 'none'}),
    ]),
],)

tab = [refresh_btn, header, graph, select_supply, data_table, user_inv]


def set_callbacks(app):
    @app.callback(
        [Output('plot-user', 'children'),
         Output('plot-filename', 'children'),
         Output('plot-datetime', 'children'),
         Output('netcon', 'children'),
         Output('revenue', 'children'),
         Output('vc', 'children'),
         Output('fc', 'children'),
         Output('supply-utilization', 'figure')],
        [Input('refresh', 'n_clicks')])
    def refresh_data(click):
        viz = visualize.Visualize(request.authorization['username'])
        if viz.file is None:
            plot_user = ""
            plot_filename = ""
            plot_datetime = ""
            header_netcon = None
            header_totalrev = None
            header_totalvc = None
            header_totalfc = None
            fig_utilization = go.Figure()
        else:
            plot_status = viz.plot_status
            header = viz.gen_header()
            plot_user = plot_status['upload_user']
            plot_filename = plot_status['upload_filename']
            plot_datetime = plot_status['upload_datetime']
            header_netcon = header['total_netcon']
            header_totalrev = header['total_rev']
            header_totalvc = header['total_vc']
            header_totalfc = header['total_fc']
            fig_utilization = viz.plt_utilization()
        plot_user = "User: " + plot_user
        plot_filename = "Filename: " + plot_filename
        plot_datetime = "Datetime: " + plot_datetime
        return plot_user, plot_filename, plot_datetime, header_netcon, header_totalrev, header_totalvc, header_totalfc, fig_utilization

    @app.callback(
        [Output('select-supply', 'children'),
         Output('table-product', 'children'),
         Output('table-route', 'children'),
         Output('route-map', 'figure')],
        [Input('supply-utilization', 'clickData')])
    def update_supply(click):
        viz = visualize.Visualize(request.authorization['username'])
        if viz.file is None:
            supply_txt = "Supply: "
            table_product = None
            table_route = None
            fig_routemap = go.Figure()
        else:
            supply_txt = viz.gen_supply(click)
            table_product, table_route = viz.gen_supplytable(click)
            fig_routemap = viz.plt_routemap(click)
        return supply_txt, table_product, table_route, fig_routemap
