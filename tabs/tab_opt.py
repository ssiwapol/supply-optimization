import base64
import datetime

from pytz import timezone
import dash_core_components as dcc
import dash_html_components as html
from dash.dependencies import Input, Output, State
from flask import request

from models import optimize
import mod

# get config data
p = mod.PathFile()

# dash components
upload_btn = dcc.Upload(
    id='upload-data',
    children=html.Div([
        'Drag and Drop or ',
        html.A('Select Files')
    ]),
    style={
        'width': '100%',
        'height': '60px',
        'lineHeight': '60px',
        'borderWidth': '1px',
        'borderStyle': 'dashed',
        'borderRadius': '5px',
        'textAlign': 'center',
        'margin': '10px'
    })

upload_col = html.Div([
    html.H3('Upload'),
    html.P(id='upload-user'),
    html.P(id='upload-ip'),
    html.P(id='upload-filename'),
    html.P(id='upload-datetime'),
    html.P(id='upload-filetype'),
],
    style={'width': '30%', 'display': 'inline-block', 'vertical-align': 'top'},
)

validate_col = html.Div([
    html.H3('Validation'),
    html.P(id='validate-sheet'),
    html.P(id='validate-feas'),
    html.P(html.A('download error', href="/download/error", id='error')),
],
    style={'width': '30%', 'display': 'inline-block', 'vertical-align': 'top'},
)

solve_col = html.Div([
    html.H3('Solve ', style={'display': 'inline-block', 'padding-right': '15px'}),
    html.Button(id='solve', children='Solve'),
    dcc.RadioItems(
        id='solver-engine',
        options=[{'label': 'CBC', 'value': 'cbc'}, {'label': 'GLPK', 'value': 'glpk'}, ],
        value='cbc',
        labelStyle={'display': 'inline-block'},
        style={'margin-top': '0px'}
    ),
    html.P(id='optimize-start'),
    html.P(id='optimize-end'),
    html.P(id='optimize-total'),
    html.P(id='optimize-status'),
    html.P(id='optimize-condition'),
    html.P(html.A('download output', href="/download/output", id='output')),
],
    style={'width': '30%', 'display': 'inline-block', 'vertical-align': 'top'},
)

# dash tab
tab = [
    html.H1('WELCOME TO SUPPLY OPTIMIZATION'),
    html.P('Please upload your input data'),
    html.A("input template", href="/download/input_template"),
    upload_btn,
    html.Div([upload_col, validate_col, solve_col],
             style={"margin-top": "10px"}),
]


# call backs
def set_callbacks(app):
    @app.callback([Output('upload-user', 'children'),
                   Output('upload-ip', 'children'),
                   Output('upload-filename', 'children'),
                   Output('upload-datetime', 'children'),
                   Output('upload-filetype', 'children'),
                   Output('validate-sheet', 'children'),
                   Output('validate-feas', 'children'),
                   Output('error', 'style'),
                   Output('solve', 'style'), ],
                  [Input('upload-data', 'contents')],
                  [State('upload-data', 'filename')])
    def start_optimize(content, filename):
        user = request.authorization['username']
        upload_datetime = datetime.datetime.now(
            timezone('Asia/Bangkok')).strftime("%Y-%m-%d %H:%M:%S")
        opt = optimize.Optimize(user)
        # no file upload
        if filename is None:
            upload_user_txt = ""
            upload_ip_txt = ""
            upload_filename_txt = ""
            upload_datetime_txt = ""
            upload_filetype_txt = ""
            validate_sheet_txt = ""
            validate_feas_txt = ""
            error_style = {'display': 'none'}
            solve_style = {'display': 'none'}
        # upload file with wrong format
        elif filename[-5:] != ".xlsx" and filename[-4:] != ".xls":
            upload_user_txt = user
            upload_ip_txt = str(request.remote_addr)
            upload_filename_txt = str(filename)
            upload_datetime_txt = upload_datetime
            upload_filetype_txt = "ERROR - please upload in excel format"
            validate_sheet_txt = ""
            validate_feas_txt = ""
            error_style = {'display': 'none'}
            solve_style = {'display': 'none'}
        else:
            upload_user_txt = user
            upload_ip_txt = str(request.remote_addr)
            upload_filename_txt = str(filename)
            upload_datetime_txt = upload_datetime
            upload_filetype_txt = "PASS"
            upload_status = {'upload_user': user, 'upload_ip': str(
                request.remote_addr), 'upload_filename': filename, 'upload_datetime': upload_datetime}
            # get content from upload file and validate sheet
            content_type, content_string = content.split(',')
            decoded = base64.b64decode(content_string)
            validate_sheet_status = opt.validate_sheet(decoded, upload_status)
            if sum([x['error'] for x in validate_sheet_status.values()]) > 0:
                validate_sheet_error = [(i, {'column': v['column'], 'duplicate': v['duplicate']})
                                        for i, v in validate_sheet_status.items() if validate_sheet_status[i]['error'] > 0]
                validate_sheet_txt = str(validate_sheet_error)
                validate_feas_txt = ""
                error_style = {'display': 'none'}
                solve_style = {'display': 'none'}
            else:
                validate_sheet_txt = "PASS"
                # validate feasible
                opt.import_data()
                validate_feas_status = opt.validate_feas()
                if sum(validate_feas_status.values()) > 0:
                    validate_feas_error = [x for x, val in validate_feas_status.items() if val > 0]
                    validate_feas_txt = "ERROR - " + str(validate_feas_error)
                    error_style = {'display': 'inline'}
                    solve_style = {'display': 'none'}
                else:
                    validate_feas_txt = "PASS"
                    error_style = {'display': 'inline'}
                    solve_style = {'display': 'inline'}
        upload_user_txt = "User: " + upload_user_txt
        upload_ip_txt = "IP Address: " + upload_ip_txt
        upload_filename_txt = "Filename: " + upload_filename_txt
        upload_datetime_txt = "Date/Time: " + upload_datetime_txt
        upload_filetype_txt = "File Format: " + upload_filetype_txt
        validate_sheet_txt = "- Sheet: " + validate_sheet_txt
        validate_feas_txt = "- Feasible: " + validate_feas_txt
        return upload_user_txt, upload_ip_txt, upload_filename_txt, upload_datetime_txt, upload_filetype_txt, validate_sheet_txt, validate_feas_txt, error_style, solve_style

    @app.callback([Output('optimize-start', 'children'),
                   Output('optimize-end', 'children'),
                   Output('optimize-total', 'children'),
                   Output('optimize-status', 'children'),
                   Output('optimize-condition', 'children'),
                   Output('output', 'style'), ],
                  [Input('solve', 'n_clicks'),
                   Input('solver-engine', 'value')],)
    def solve(click, solver_engine):
        user = request.authorization['username']
        if click is None:
            optimize_start_txt = ""
            optimize_end_txt = ""
            optimize_total_txt = ""
            optimize_status_txt = ""
            optimize_condition_txt = ""
            output_style = {'display': 'none'}
        else:
            opt = optimize.Optimize(user)
            opt.import_data()
            opt_status = opt.optimize(solve_engine=solver_engine)
            opt.gen_output(opt_status)
            opt.gen_plot(opt_status)
            optimize_start_txt = opt_status['optimize_start_time']
            optimize_end_txt = opt_status['optimize_end_time']
            optimize_total_txt = str(opt_status['optimize_solvetime_sec'])
            optimize_status_txt = opt_status['optimize_solver_status']
            optimize_condition_txt = opt_status['optimize_termination_condition']
            output_style = {'display': 'inline'}
        optimize_start_txt = "Start Time: " + optimize_start_txt
        optimize_end_txt = "End Time: " + optimize_end_txt
        optimize_total_txt = "Total Time(secs): " + optimize_total_txt
        optimize_status_txt = "Status: " + optimize_status_txt
        optimize_condition_txt = "Condition: " + optimize_condition_txt
        return optimize_start_txt, optimize_end_txt, optimize_total_txt, optimize_status_txt, optimize_condition_txt, output_style
