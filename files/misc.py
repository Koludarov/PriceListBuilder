import json
from typing import Any


def read_files(assortment_path: str, stocks_path: str) -> Any:
    """Считывает json файлы"""
    with open(assortment_path, encoding='utf-8') as f:
        assortment = json.load(f)

    with open(stocks_path, encoding='utf-8') as f:
        stocks = json.load(f)
        stocks = stocks['rows']

    return assortment, stocks
