import datetime

import gspread

from data.data_flow import collect_categories, add_lists_categories
from files.misc import read_files
from data.insert_data import add_items_to_list, add_supplies_list
from config import CREDENTIALS_PATH, STOCKS_PATH, ASSORTMENT_PATH


def main():
    """Запускает скрипт"""
    print(f'Программа запущена {datetime.datetime.now()}')
    gc = gspread.service_account(CREDENTIALS_PATH)
    # Запрашиваем название Документа в Google Disk
    sheet_name = input('Введите название файла Google Sheet:\n')
    sheet = gc.open(sheet_name)
    assortment, stocks = read_files(ASSORTMENT_PATH, STOCKS_PATH)
    category, supplies_storage = collect_categories(stocks)
    add_lists_categories(category, sheet)
    supplies_storage['count'] = 0
    supplies_storage = add_items_to_list(
        stocks,
        assortment,
        supplies_storage.keys(),
        supplies_storage,
        sheet
    )
    add_supplies_list(supplies_storage, sheet)
    print(f'Программа завершена {datetime.datetime.now()}')


if __name__ == "__main__":
    main()
