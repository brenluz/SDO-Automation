from gspread import Worksheet

from src.podio import Tarefa

meses = {
    1: "Janeiro",
    2: "Fevereiro",
    3: "Março",
    4: "Abril",
    5: "Maio",
    6: "Junho",
    7: "Julho",
    8: "Agosto",
    9: "Setembro",
    10: "Outubro",
    11: "Novembro",
    12: "Dezembro"
}


def findEmpty(worksheet, column):
    col_values = worksheet.col_values(column)
    for idx, value in enumerate(col_values, start=1):
        if value == '':
            return idx
    return len(col_values) + 2


cell_format = {
    "backgroundColor": {
        "red": 0.0,
        "green": 0.5,
        "blue": 0.5
    },
    "textFormat": {
        "foregroundColor": {
            "red": 1.0,
            "green": 1.0,
            "blue": 1.0
        },
        "bold": True
    },
    "borders": {
        "top": {
            "style": "SOLID"
        },
        "bottom": {
            "style": "SOLID"
        },
        "left": {
            "style": "SOLID"
        },
        "right": {
            "style": "SOLID"
        }
    },
    "horizontalAlignment": "CENTER",
    "verticalAlignment": "MIDDLE"
}


def find_last_non_empty_row(worksheet, column):
    column = ord(column) - ord('A') + 1
    col_values = worksheet.col_values(column)
    for idx in range(len(col_values) - 1, -1, -1):
        if col_values[idx] != '':
            return idx + 3  # Return the next row after the last non-empty cell
    return 1  # If the column is completely empty, return the first row


column_map = {
    range(1, 8): 'B',
    range(8, 15): 'H',
    range(15, 22): 'N',
    range(22, 29): 'T',
    range(29, 32): 'Z'
}


def findArea(worksheet, dia):
    area = None
    for day_range, col in column_map.items():
        if int(dia) in day_range:
            first_empty_row = find_last_non_empty_row(worksheet, col)
            area = f"{col}{first_empty_row}:{chr(ord(col) + 4)}{first_empty_row}"
            break
    return area


def get_column_for_day(dia):
    for day_range, col in column_map.items():
        if int(dia) in day_range:
            return col
    raise KeyError(f"Day {dia} is not in any range")
