from dash import dcc, html, Dash, Input, Output, dash_table
import pandas as pd
import openpyxl

app = Dash(__name__)
server = app.server
xlsx_file = 'ticks.xlsx'
xlsx_data = pd.read_excel(xlsx_file, sheet_name=None)
tickers = list(xlsx_data.keys())

app.layout = html.Div([
    html.H1("Ticker Selector"),
    dcc.Dropdown(
        id='ticker-selector',
        options=[{'label': ticker, 'value': ticker} for ticker in tickers],
        value=tickers[0]
    ),
    html.Div(id='table-container')
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

if __name__ == '__main__':
    app.run_server(debug=True)
