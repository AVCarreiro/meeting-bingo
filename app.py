import random
import os
from os.path import basename
import numpy as np
import base64
import datetime
import io

import pandas as pd

import dash
import dash_core_components as dcc
import dash_html_components as html
from dash.dependencies import Input, Output, State
import dash_table

import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

EMAIL_SENDER = 'bingomeeting@gmail.com'
EMAIL_PWD = 'BingoMeetingCOVID19'
EMAIL_SUBJECT = 'Meeting Bingo Card'
EMAIL_BODY = "Please find your Meeting Bingo card attached!\n Have fun on those meetings!"

def main():
    external_stylesheets = ['https://codepen.io/chriddyp/pen/bWLwgP.css']

    app = dash.Dash(__name__, external_stylesheets=external_stylesheets)

    colors = {
    'background': '#FFFFFF',
    'text': '#000000'
    }

    input_style = {
        'width': '50%',
        'height': '40px',
        'lineHeight': '60px',
        'borderWidth': '1px',
        'borderRadius': '5px',
        'textAlign': 'center',
        'margin': '10px',
    }

    label_style = {
        'margin': '10px'
    }

    upload_style = {
        'width': '60%',
        'height': '60px',
        'lineHeight': '60px',
        'borderWidth': '1px',
        'borderStyle': 'dashed',
        'borderRadius': '5px',
        'textAlign': 'center',
        'margin': '10px'
    }

    result_list_style = {
        'width': '30%', 
        'display': 'inline-block',
        'textAlign':'center',
        'margin':'30px'
    }


    app.layout = html.Div(style={'backgroundColor': colors['background']}, children=[
        html.H1(
            children='Meeting Bingo',
            style={
                'textAlign': 'center',
                'color': colors['text']
            }
        ),

        html.Div(children='A bingo card generator for your Meeting Bingo! ', style={
            'textAlign': 'center',
            'color': colors['text']
        }),

        html.Div( #Div breaking page in 3
            [
                html.Div(
                    [
                        html.Div([
                            html.Label('Number of Rows', style=label_style),
                            dcc.Input(id='n_rows', type='number',value=4, min=1, max=15, style=input_style)
                            ],           
                            style= {'display': 'inline-block'}
                        ),

                        html.Div([
                            html.Label('Number of Columns', style=label_style),
                            dcc.Input(id='n_cols',type='number', value=4, min=1, max=15, style=input_style),
                            ],           
                            style= {'display': 'inline-block'}
                        ),      

                        html.Label('Bingo Entries File', style=label_style),
                        dcc.Upload(
                            id='upload_entries',
                            children=html.Div([
                                'Drag and Drop or ',
                                html.A('Select Files')
                            ]),
                            style=upload_style
                        ),
                        html.Div(id="entries_filename"),

                        html.Label('E-mail Addresses List File', style=label_style),
                        dcc.Upload(
                            id='upload_emails',
                            children=html.Div([
                                'Drag and Drop or ',
                                html.A('Select Files')
                            ]),
                            style=upload_style
                        ),
                        html.Ul(id="emails_filename"),

                        html.Button(id='submit_button', n_clicks=0, children='Submit', style=label_style),
                        html.Div(id="result"),
                    ],
                    style = {'width': '30%', 'display': 'inline-block'}
                ),

                html.Div([
                        #html.Label('Bingo Entries', style=label_style),
                        html.Div(id='entries_list')
                    ],
                    style = result_list_style
                ),

                html.Div([
                        #html.Label('Emails', style=label_style),
                        html.Div(id = 'emails_list')
                    ],
                    style = result_list_style
                )
            ],
            style = {'display':'flex'}
        )

        
          
    ])

    @app.callback(
        [Output("entries_filename", "children"), Output("entries_list", "children")],
        [Input("upload_entries", "filename"), Input("upload_entries", "contents")],
    )
    def process_entries_file(filename, contents):
        if contents is None:
            return [], []

        entries = parse_contents(contents)

        df = pd.DataFrame(entries, columns=['Bingo Entries'])
        table = dash_table.DataTable(
            id='entries_filled',
            columns=[{"name": i, "id": i} for i in df.columns],
            data=df.to_dict("rows"),
            style_cell={'textAlign': 'left'},
        )

        return [html.H6(filename)],[table]

    @app.callback(
        [Output("emails_filename", "children"), Output("emails_list", "children")],
        [Input("upload_emails", "filename"), Input("upload_emails", "contents")],
    )
    def process_emails_file(filename, contents):
        if contents is None:
            return [],[]
        
        entries = parse_contents(contents)

        df = pd.DataFrame(entries, columns=['E-mail Addresses'])
        table = dash_table.DataTable(
            id='emails_filled',
            columns=[{"name": i, "id": i} for i in df.columns],
            data=df.to_dict("rows"),
            style_cell={'textAlign': 'left'},
        )

        return [html.H6(filename)], [table]


    @app.callback(
        Output("result", "children"),
        [   Input("submit_button", "n_clicks")],
        [
            State("upload_entries", "filename"), State("upload_entries", "contents"),
            State("upload_emails", "filename"), State("upload_emails", "contents"),
            State("n_rows", "value"), State("n_cols", "value")
        ]
    )
    def submit(n_clicks, entries_filename, entries_contents, emails_filename, emails_contents, n_rows, n_cols):
        if entries_contents is None or emails_contents is None:
            return []

        entries = parse_contents(entries_contents)

        emails = parse_contents(emails_contents)        
        
        cards = []
        for email in emails:
            card = create_bingo_card(email, entries, n_rows, n_cols, send=True)
            cards.append(card)

        return [
            html.H5("Read {:d} bingo entries and {:d} e-mail addresses!".format(len(entries), len(emails))),
        ]
        
    
    app.run_server(debug=True)


def parse_contents(contents):
    if contents is None:
        return []

    try:
        content_type, content_string = contents.split(',')
        decoded = base64.b64decode(content_string)
        stream = io.StringIO(decoded.decode('utf-8'))
        entries = stream.getvalue().splitlines()
        stream.close()
    except Exception as e:
        print(e)
        return [
            'There was an error processing this file. Please provide a .txt or .csv with one entry per line'
        ]

    return entries

def create_bingo_card(email, entries, n_rows, n_cols, send=True):
    total_entries = n_rows * n_cols
    if total_entries > len(entries):
        print("[ERROR] Number of given entries (in file) is lower than total entries requested (rows x columns).")
        return None

    name = email.split('@')[0]
    chosen_entries = random.sample(entries, k=total_entries)
    #chosen_entries = np.array(chosen_entries).reshape((n_rows, n_cols))

    table = generate_table(chosen_entries, n_rows, n_cols)
    table_html = generate_table_raw(chosen_entries, n_rows, n_cols, name)
    
    out_file = name + ".html"
    with open(out_file, 'w') as of:
        of.write(table_html)

    try:
        if send:
            send_email(EMAIL_SENDER, EMAIL_PWD, [email], EMAIL_SUBJECT, EMAIL_BODY, [out_file])
    except Exception as e:
        print(e)
        return 'There was a problem sending the bingo cards in e-mail'

    return table


# Generates an HTML table representation of the bingo card for entries
def generate_table_raw(entries, n_rows, n_cols, name):
    head = """
        <!DOCTYPE html>
        <html>
        <head>
	        <title>Meeting BINGO</title>
	        <style>
            body {align: center; background-color: #FFFFFF; font-weight: 400; font-family: "Open Sans", "HelveticaNeue", "Helvetica Neue", Helvetica, Arial, sans-serif;}
            h1   {color: black; text-align: center;}
            table {align: center; margin: 0 auto; border: 3px solid #003333; width: 80%; text-align: center; border-radius: 10px}
            tr {height: 150px;}
            td {border: 5px solid #006666; width: 150px; vertical-align: center; border-radius: 10px;}
            </style>
        </head>
        <body>"""

    head += "\n<h1> Meeting BINGO: {:s}</h1>".format(name)

    res = "<table>\n"
    for i, entry in enumerate(entries):
        if i % n_cols == 0:
            res += "\t<tr>\n"
        res += "\t\t<td>" + entry + "</td>\n"
        if i>0 and i % n_cols == n_cols - 1:
            res += "\t</tr>\n"
    res += "</table>\n"
    res += "</body></html>"
    return head + res

def generate_table(entries, n_rows, n_cols):
    entries_matrix = np.array(entries).reshape((n_rows, n_cols))
    return html.Table(
        [html.Tr([
            html.Td(entries_matrix[row][col], 
                style= {
                    'border': '5px solid #006666',
  	                #'margin': '10px',
                    'width': '150px',
                    'text-align': 'center'
                }) for col in range(n_cols)
        ], 
        style={
            'height': '150px' 
        }) for row in range(n_rows)],
        style={
            'border': '1px solid #006666',
            'width': '30%',
  	        'text-align': 'center',
  	        'vertical-align': 'center'
        }
    )

def send_email(from_address, pwd, to_address, subject, body, attachments):
    msg = MIMEMultipart()

    msg['From'] = from_address
    msg['To'] = to_address if not isinstance(to_address, list) else ','.join(to_address)
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    for f in attachments or []:
        with open(f, "rb") as fil:
            part = MIMEApplication(
                fil.read(),
                Name=basename(f)
            )
        # After the file is closed
        part['Content-Disposition'] = 'attachment; filename="%s"' % basename(f)
        msg.attach(part)

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(from_address, pwd)
    text = msg.as_string()
    server.sendmail(from_address, to_address, text)
    server.quit()

if __name__ == '__main__':
    main()