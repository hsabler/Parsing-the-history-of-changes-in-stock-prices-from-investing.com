import investpy
import pandas as pd
from openpyxl import Workbook
from pandas import ExcelWriter

def get_historical_data_to_excel(tickers, start_date, end_date, output_file):
    df_all = pd.DataFrame()
    wb = Workbook()
    ws = wb.active
    ws.title = "Historical Data"

    for ticker in tickers:
        try:
            historical_data = investpy.get_stock_historical_data(stock=ticker, country='united states', from_date=start_date, to_date=end_date)
            historical_data['Ticker'] = ticker
            df_all = pd.concat([df_all, historical_data], ignore_index=True)

        except Exception as e:
            print(f"Error fetching data for {ticker}: {str(e)}")
            
    with ExcelWriter(output_file) as writer:
        df_all.to_excel(writer, sheet_name='Historical Data', index=False)

    print(f"Historical data saved to {output_file}")


if __name__ == "__main__":
    tickers = ['TSLA']
    start_date = '06/05/1992'
    end_date = '29/09/2024'
    output_file = 'путь к файлу для выгрузки результата'

    get_historical_data_to_excel(tickers, start_date, end_date, output_file)
