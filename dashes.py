from dash import dcc, html, Dash, Input, Output, dash_table
from dash.dependencies import Input, Output, State

import pandas as pd
import openpyxl

def get_dates(date, dicf):
    ret = pd.DataFrame()
    class BreakOutOfLoops(Exception):
        pass

    for name in dicf.keys():
        df = dicf[name]
        dates = df['Dates'].tolist()
        try:
            for elem in dates:
                for d in elem:
                    if date in elem:
                        row = df[df['Dates'] == elem].reset_index(drop=True)
                        row['Unnamed: 0'] = name
                        ret = pd.concat([ret, row])
                        raise BreakOutOfLoops
        except BreakOutOfLoops:
            pass
    ret = ret.sort_values(by='Average Volume Fraction', ascending=False)
    ret = ret.rename(columns={ret.columns[0]: 'Ticker'})
    return ret

def is_third_friday(date):
    # Function to check if a date is the third Friday of the month
    # Implement your logic here and return True or False accordingly
    # Example implementation assuming `date` is a string in 'YYYY-MM-DD' format:
    parsed_date = pd.to_datetime(date)
    return (parsed_date.weekday() == 4) and (15 <= parsed_date.day <= 21)

def format_date_cell(value):
    # Custom formatting function to apply styling to individual elements in the 'Dates' column
    if isinstance(value, list):
        return html.Ul([
            html.Li(html.Span(date, style={'color': 'red'})) if is_third_friday(date) else html.Li(date)
            for date in value
        ])
    return value

app = Dash(__name__)
server = app.server
xlsx_file = 'ticks.xlsx'
xlsx_data = pd.read_excel(xlsx_file, sheet_name=None)
tickers = list(xlsx_data.keys())

app.layout = html.Div([
    html.H1("Ticker Selector"),
    html.H2("This application allows you to select a ticker and retrieve relevant data."),
    dcc.Dropdown(
        id='ticker-selector',
        options=[{'label': ticker, 'value': ticker} for ticker in tickers],
        value=tickers[0]
    ),
    html.Div(id='table-container'),
    dcc.Input(id='date-input', type='text', placeholder='Enter a date'),
    html.Button('Get Dates', id='get-dates-button'),
    html.Div(id='date-table-container'),
])

@app.callback(
    Output('table-container', 'children'),
    Input('ticker-selector', 'value')
)
def display_table(selected_ticker):
    df = xlsx_data[selected_ticker]
    return dash_table.DataTable(
        data=df.to_dict('records'),
        columns=[
            {'name': i, 'id': i, 'type': 'text', 'presentation': 'markdown', 'format': format_date_cell} if i == 'Dates' else {'name': i, 'id': i, 'type': 'text'}
            for i in df.columns
        ]
    )

@app.callback(
    Output('date-table-container', 'children'),
    Input('get-dates-button', 'n_clicks'),
    State('date-input', 'value')
)
def display_date_table(n_clicks, date_input):
    if n_clicks and date_input:
        df = get_dates(date_input, xlsx_data)
        return dash_table.DataTable(
            data=df.to_dict('records'),
            columns=[{'name': i, 'id': i} for i in df.columns]
        )

if __name__ == '__main__':
    app.run_server(debug=True)

