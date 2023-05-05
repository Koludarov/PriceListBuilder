from typing import Dict, List

import gspread

from sheets.worksheets_update import create_list, change_row_height
from data.data_flow import get_data


def add_items_to_list(
        stocks: List,
        assortment: List,
        supplies: List,
        supplies_storage: Dict,
        sheet: gspread.spreadsheet.Spreadsheet
) -> Dict:
    """Добавляет товары в листы, соответствующие их категориям"""
    exceptions_rows = []

    for stock_item in stocks:
        external_code = stock_item['externalCode']
        for assortment_item in assortment:
            if external_code == assortment_item['externalCode']:
                if stock_item['folder']['pathName'] == 'Номенклатура':
                    cat = stock_item['folder']['name']
                else:
                    cat = (stock_item['folder']['pathName']
                           .replace('Номенклатура/', '').replace('/', '_').replace(' ', '_'))
                row_data = get_data(assortment_item, stock_item)

                # Заполняем словарь для категории Расходные материалы
                if 'Расходные_материалы' in cat:
                    if stock_item['folder']['name'] in supplies:
                        supplies_storage['count'] += 1
                        supplies_storage[stock_item['folder']['name']] += [row_data]
                    continue
                exception_handled = False
                while not exception_handled:
                    try:
                        worksheet = sheet.worksheet(cat)
                        worksheet.append_row(row_data, value_input_option='USER_ENTERED')
                        exception_handled = True
                    except gspread.exceptions.APIError as error:
                        print(f"Ошибка при создании {row_data[0]}. Описание: {error}")
                        exceptions_rows.append(row_data)
                if row_data in exceptions_rows:
                    exceptions_rows.remove(row_data)
                break

    return supplies_storage


def add_supplies_list(
        supplies_storage: Dict,
        sheet: gspread.spreadsheet.Spreadsheet
) -> None:
    """Добавляет лист расходные материалы и наполняет его товарами"""

    create_list('Расходные материалы', sheet)
    worksheet = sheet.worksheet('Расходные материалы')
    amount = supplies_storage.pop('count') + len(supplies_storage.keys()) - 99
    if amount > 0:
        for i in range(amount):
            worksheet.insert_row([], 100)

    row_index = 2  # Индекс строки для группирования
    for sup_cat, sup_items in supplies_storage.items():
        if not sup_items:
            continue
        worksheet.append_row([sup_cat, '', '', '', '', ''], value_input_option='USER_ENTERED')

        for sup_item in sup_items:
            worksheet.append_row(sup_item, value_input_option='USER_ENTERED')

        change_row_height(worksheet, row_index - 1)
        worksheet.add_dimension_group_rows(row_index, row_index + len(sup_items))
        print(f'Создаём группу {row_index, row_index + len(sup_items)}')
        row_index += len(sup_items) + 1
