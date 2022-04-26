import pandas as pd
import numpy as np
import requests
import xlsxwriter
import math
from secrets import IEX_CLOUD_API_TOKEN

stocks = pd.read_csv('sp_500_stocks.csv')

symbol='AAPL'
api_url = f'https://sandbox.iexapis.com/stable/stock/{symbol}/quote?token={IEX_CLOUD_API_TOKEN}'
data = requests.get(api_url).json()

price = data["latestPrice"]
market_cap = data["marketCap"]
columns = ['Ticker', 'Price','Market Capitalization', 'Number Of Shares to Buy']

dataframe = pd.DataFrame(columns = columns)

for stock in stocks["Ticker"][:100]:
    api_url = f'https://sandbox.iexapis.com/stable/stock/{stock}/quote?token={IEX_CLOUD_API_TOKEN}'
    data = requests.get(api_url).json()
    dataf = pd.DataFrame(data = [[stock, data["latestPrice"], data["marketCap"], "N/A"]],
    columns= columns)
    dataframe = pd.concat([dataframe, dataf], 
    axis= 0,
    ignore_index= True)

def chunks(lst, limit):
    for i in range(0, len(lst), limit):
        yield lst[i: i + limit]

symbol_groups = list(chunks(stocks["Ticker"], 100))
symbol_strings = list()
for i in range(0, len(symbol_groups)):
    symbol_strings.append(','.join(symbol_groups[i]))

dataframe = pd.DataFrame(columns= columns)

for symbol_string in symbol_strings[:1]:
    batch_api_call_url = f'https://sandbox.iexapis.com/stable/stock/market/batch/?types=quote&symbols={symbol_string}&token={IEX_CLOUD_API_TOKEN}'
    data = requests.get(batch_api_call_url).json()
    for symbol in symbol_string.split(','):
        dataf = pd.DataFrame(data = [
            [symbol,
            data[symbol]["quote"]["latestPrice"],
            data[symbol]["quote"]["marketCap"],
            "N/A"]
            ],
        columns= columns)
        dataframe = pd.concat([dataframe, dataf],
        axis= 0,
        ignore_index= True)

portfolio_size = input("Enter the value of your portfolio:")

try:
    val = float(portfolio_size)
except ValueError:
    print("That's not a number! \n Try again:")
    portfolio_size = input("Enter the value of your portfolio:")

position_size = float(portfolio_size) / len(dataframe.index)
for i in range(0, len(dataframe['Ticker'])-1):
    dataframe.loc[i, 'Number Of Shares to Buy'] = math.floor(position_size / dataframe['Price'][i])
dataframe

writer = pd.ExcelWriter('recommended_trades.xlsx', engine='xlsxwriter')
dataframe.to_excel(writer, sheet_name='Recommended Trades', index = False)

background_color = '#0a0a23'
font_color = '#ffffff'

string_format = writer.book.add_format(
        {
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

dollar_format = writer.book.add_format(
        {
            'num_format':'$0.00',
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

integer_format = writer.book.add_format(
        {
            'num_format':'0',
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )
column_formats = { 
                    'A': ['Ticker', string_format],
                    'B': ['Price', dollar_format],
                    'C': ['Market Capitalization', dollar_format],
                    'D': ['Number of Shares to Buy', integer_format]
                    }

for column in column_formats.keys():
    writer.sheets['Recommended Trades'].set_column(f'{column}:{column}', 20, column_formats[column][1])
    writer.sheets['Recommended Trades'].write(f'{column}1', column_formats[column][0], string_format)

writer.save()