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


app = Dash(__name__)
server = app.server
xlsx_file = 'ticks.xlsx'
xlsx_data = pd.read_excel(xlsx_file, sheet_name=None)
tickers = list(xlsx_data.keys())

app.layout = html.Div([
    html.H1("Ticker Selector"),
    html.H2("Here you can select a ticker and view all days where abnormal daily trading volume occured at the open, hinting at a potential unlock. Average Trading Volume shows the fraction of daily volume traded at the open. Average Percent Change shows the average change from the day before's close to the high volume day's open, with the standard deviation as well. The units are percent. Below you can input a date and find the highest fraction stocks for that date. Please enter dates in the MM-DD format.", style={'font-size': '19px', 'font-weight': 'normal'}),
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
        columns=[{'name': i, 'id': i} for i in df.columns]
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
