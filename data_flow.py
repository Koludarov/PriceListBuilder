import time
import textwrap
from typing import List, Any, Dict, Set

import gspread

from worksheets_update import create_list, delete_worksheet


def collect_categories(stocks: List) -> Any:
    """Собирает категории товаров и расходных материалов"""
    category = set()
    supplies = set()
    for stock_item in stocks:
        if 'Расходные материалы' in stock_item['folder']['pathName']:
            supplies.add(stock_item['folder']['name'])
        elif stock_item['folder']['pathName'] == 'Номенклатура':
            category.add(stock_item['folder']['name'])
        elif 'Ламинирование/' in stock_item['folder']['pathName']:
            continue
        elif 'Брови/' in stock_item['folder']['pathName']:
            continue
        else:
            category.add(stock_item['folder']['pathName']
                         .replace('Номенклатура/', '').replace('/', '_').replace(' ', '_'))
    return category, {x: [] for x in supplies}


def get_data(assort_item: Dict, stock_item: Dict) -> List:
    """Получает данные о товаре"""
    wrapper = textwrap.TextWrapper(width=13)

    # получаем данные о цене товара
    prices = assort_item['salePrices']
    retail_price = prices[0]['value']/100
    price_5k = prices[1]['value']/100
    price_15k = prices[2]['value']/100
    price_100k = prices[3]['value']/100

    # Переносим слова, чтобы не выходило за пределы клетки
    name = wrapper.fill(text=stock_item['name'])
    row_data = [name]
    if 'image' in stock_item:
        if 'miniature' in stock_item['image']:
            image_url = stock_item['image']['miniature']['downloadHref'].replace('sel-prod', 'prod')
            row_data.append(f'=IMAGE("{image_url}")')
        else:
            row_data.append('')
    else:
        row_data.append('')

    row_data.extend([retail_price, price_5k, price_15k, price_100k])
    return row_data


def add_lists_categories(category: Set, sheet: gspread.spreadsheet.Spreadsheet) -> None:
    """Добавляет листы в Google Sheets по категориям"""

    exceptions_category = []

    for category_item in category:
        exception_handled = False
        while not exception_handled:
            try:
                create_list(category_item, sheet)
                exception_handled = True
            except gspread.exceptions.APIError as error:
                print(f"Ошибка при создании {category_item}. Описание: {error}")
                exceptions_category.append(category_item)
                delete_worksheet(category_item, sheet)
                print('Ждём 2 минуты')
                time.sleep(121)
        if category_item in exceptions_category:
            exceptions_category.remove(category_item)

    # Ждём минуту, чтобы прошло время с последнего запроса
    time.sleep(121)
    # Удаляем Лист1, который создаётся автоматически
    delete_worksheet('Лист1', sheet)
