import os
import win32com.client as win32
import datetime
from pandas_datareader import data

working_dir = os.getcwd()

ExcelApp = win32.Dispatch('Excel.Application')
ExcelApp.visible = True
wbStock = ExcelApp.workbooks.add
wbStock.SaveAs(os.path.join(working_dir, 'Output', 'Stock Price Pull {0}.xlsx'.format(datetime.datetime.now().strftime('%m-%d-%Y %HH_%MM_%SS'))))


# Live Price Pull
tickers = ['MSFT', 'TLSA', 'GOOG', 'AAPL', 'DBX', 'FB', 'AMZN']
live_price_worksheet_Name = 'Live Price'

livePrice = data.get_quote_yahoo(tickers)
livePrice.reset_index(inplace=True)
livePrice.rename(columns={'index': 'Ticker'}, inplace=True)

wsPrice = wbStock.worksheets.add
wsPrice.Name = live_price_worksheet_Name

# Inserting Column Names
wbStock.Worksheets(live_price_worksheet_Name).Range(
    wbStock.Worksheets(live_price_worksheet_Name).cells(1, 1),
    wbStock.worksheets(live_price_worksheet_Name).cells(1, livePrice.shape[1])).value = livePrice.columns.tolist()

# Inserting Price Data
wbStock.Worksheets(live_price_worksheet_Name).Range(
    wbStock.Worksheets(live_price_worksheet_Name).cells(2, 1),
    wbStock.worksheets(live_price_worksheet_Name).cells(livePrice.shape[0] + 1, livePrice.shape[1])).value = livePrice.values.tolist()


"""
Import Historical Prices
"""
start_date, end_date = '2017-1-1', '2019-12-31'

for ticker in tickers:
    historicalPrice = data.DataReader(ticker, start=start_date, end=end_date, data_source='yahoo')

    wb = wbStock.worksheets.add
    wb.name = ticker

    # Inserting Column Names #TODO
    wbStock.Worksheets(ticker).Range(
    wbStock.Worksheets(ticker).cells(1, 2),
    wbStock.worksheets(ticker).cells(1, historicalPrice.shape[1] + 1)).value = historicalPrice.columns.tolist()

    # index
    wbStock.Worksheets(ticker).Range('A1').value = 'Date'

    wbStock.Worksheets(ticker).Range(
    wbStock.Worksheets(ticker).cells(2, 1),
    wbStock.worksheets(ticker).cells(historicalPrice.shape[0] + 1, 1)).value = [[d] for d in historicalPrice.index.to_pydatetime().astype(str)]

    # price data 
    wbStock.Worksheets(ticker).Range(
    wbStock.Worksheets(ticker).cells(2, 2),
    wbStock.worksheets(ticker).cells(historicalPrice.shape[0] + 1, historicalPrice.shape[1] + 1)).value = historicalPrice.values.tolist()

    # formatting prices
    wbStock.Worksheets(ticker).Range(
    wbStock.Worksheets(ticker).cells(2, 2),
    wbStock.worksheets(ticker).cells(historicalPrice.shape[0] + 1, historicalPrice.shape[1] + 1)).numberformat = '#,##0.00'

    # formatting volumes
    wbStock.Worksheets(ticker).Range(
    wbStock.Worksheets(ticker).cells(2, "F"),
    wbStock.worksheets(ticker).cells(historicalPrice.shape[0] + 1, "F")).numberformat = '#,##'  

    # formatting dates
    wbStock.Worksheets(ticker).Range(
    wbStock.Worksheets(ticker).cells(2, 1),
    wbStock.worksheets(ticker).cells(historicalPrice.shape[0] + 1, 1)).numberformat = "m/d/yyyy"    

wbStock.save
wbStock.close
ExcelApp.Quit()