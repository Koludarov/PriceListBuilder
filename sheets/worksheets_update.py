import gspread


def create_list(category_item, sheet: gspread.spreadsheet.Spreadsheet) -> None:
    """Добавляет лист и заполняет заголовки столбцов"""
    sheet.add_worksheet(title=category_item, rows=100, cols=6)
    worksheet = sheet.worksheet(category_item)
    worksheet.update_cell(1, 1, 'Наименование')
    worksheet.update_cell(1, 2, 'Изображение')
    worksheet.update_cell(1, 3, 'Цена: Розница')
    worksheet.update_cell(1, 4, 'Цена: от 5 т.р.')
    worksheet.update_cell(1, 5, 'Цена: от 15 т.р.')
    worksheet.update_cell(1, 6, 'Цена: от 100 т.р.')
    change_size(worksheet)



def change_size(worksheet: gspread.spreadsheet.Worksheet) -> None:
    """Задаёт высоту и ширину клеток в 200px"""
    sheet_id = worksheet._properties['sheetId']

    # Меняем ширину второго столбца
    requests = [
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": 1,
                    "endIndex": 2
                },
                "properties": {
                    "pixelSize": 200
                },
                "fields": "pixelSize"
            }
        }
    ]

    # Меняем высоту всех строк начиная со второй
    requests += [
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "startIndex": 1
                },
                "properties": {
                    "pixelSize": 200
                },
                "fields": "pixelSize"
            }
        }
    ]

    # Отправляем запрос на изменение размеров столбца и строк
    worksheet.spreadsheet.batch_update({'requests': requests})
    print(f"Изменена ширина второго столбца и высота всех строк начиная со второй на 200px на листе {worksheet.title}.")


def change_row_height(worksheet: gspread.spreadsheet.Worksheet, row: int) -> None:
    """Задаёт высоту строки в 50px"""
    sheet_id = worksheet._properties['sheetId']

    # Меняем высоту строк в диапазоне
    requests = [
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "startIndex": row,
                    "endIndex": row + 1
                },
                "properties": {
                    "pixelSize": 100
                },
                "fields": "pixelSize"
            }
        }
    ]

    # Отправляем запрос на изменение размеров столбца и строк
    worksheet.spreadsheet.batch_update({'requests': requests})
    print(f"Изменена высота {row} строки на 50px на листе {worksheet.title}.")


def delete_worksheet(category_item, sheet):
    """Удаляет лист"""
    try:
        worksheet = sheet.worksheet(category_item)
        sheet.del_worksheet(worksheet)
    except gspread.exceptions.WorksheetNotFound:
        print(f'Лист {category_item} не был найден')
