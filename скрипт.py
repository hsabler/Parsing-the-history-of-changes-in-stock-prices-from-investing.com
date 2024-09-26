import investpy
import pandas as pd
from openpyxl import Workbook
from pandas import ExcelWriter

def get_historical_data_to_excel(tickers, start_date, end_date, output_file):
    # Создать пустой DataFrame для сохранения данных
    df_all = pd.DataFrame()
    # Создать новый эксель-файл и рабочий лист
    wb = Workbook()
    ws = wb.active
    ws.title = "Historical Data"

    # Проходимся по каждому тикеру и получаем исторические данные
    for ticker in tickers:
        try:
            historical_data = investpy.get_stock_historical_data(stock=ticker, country='united states', from_date=start_date, to_date=end_date)

            # Добавить столбец с тикером для идентификации данных
            historical_data['Ticker'] = ticker

            # Добавить полученные данные в общий DataFrame
            df_all = pd.concat([df_all, historical_data], ignore_index=True)

        except Exception as e:
            print(f"Error fetching data for {ticker}: {str(e)}")

    # Записать результат в эксель-файл
    with ExcelWriter(output_file) as writer:
        df_all.to_excel(writer, sheet_name='Historical Data', index=False)

    print(f"Historical data saved to {output_file}")


if __name__ == "__main__":
# Вставить список необходимых для анализа тикеров:
    tickers = ['TSLA']
# Даты начала и окончания торгов по тикерам
    start_date = '06/05/1992'
    end_date = '29/09/2024'
    output_file = 'путь к файлу для выгрузки результата'

    get_historical_data_to_excel(tickers, start_date, end_date, output_file)