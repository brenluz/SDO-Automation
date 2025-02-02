from gspread import Worksheet, worksheet

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

column_map = {
    range(1, 8): 'B',
    range(8, 15): 'H',
    range(15, 22): 'N',
    range(22, 29): 'T',
    range(29, 32): 'Z'
}

cell_format = {
    "backgroundColor": {
        "red": 0.0,
        "green": 0.502,
        "blue": 0.502
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


def findEmpty(planilha: Worksheet, column):
    col_values = planilha.col_values(column)
    for idx, value in enumerate(col_values, start=1):
        if value == '':
            return idx
    return len(col_values) + 2


def find_last_non_empty_row(planilha: Worksheet, column):
    coluna = ord(column) - ord('A') + 1
    texto = "SDO - PUNIÇÕES EM ABERTO NO PODIO NO DIA xx/xx"

    if planilha.find(texto, in_column=coluna) is None:
        col_values = planilha.col_values(coluna)
        for idx in range(len(col_values) - 1, -1, -1):
            if col_values[idx] == '0':
                return idx + 3
            if col_values[idx] != '':
                return idx + 4  # Return the next row after the last non-empty cell
    elif planilha.find(texto) is not None:
        return planilha.find(texto).row


def findArea(planilha: Worksheet, dia):
    area = None
    for day_range, col in column_map.items():
        if int(dia) in day_range:
            first_empty_row = find_last_non_empty_row(planilha, col)
            # print(f"First empty row for day {dia} is {first_empty_row}")
            last_col = chr(ord(col) + 4)
            if ord(last_col) > ord('Z'):
                last_col = 'A' + chr(ord(last_col) - 26)
            area = f"{col}{first_empty_row}:{last_col}{first_empty_row}"
            break
    return area


def get_column_for_day(dia):
    for day_range, col in column_map.items():
        if int(dia) in day_range:
            if ord(col) > ord('Z'):
                col = 'A' + chr(ord(col) - 26)
            return col
    raise KeyError(f"Day {dia} is not in any range")


def add_leading_zero(value: str):
    return value.zfill(2)


def addInitialText(planilha: Worksheet, dia: int, mes: int, area: str, colunaInicio, colunaEnd, linha):
    diaText = add_leading_zero(str(dia))
    mesText = add_leading_zero(str(mes))
    textoDia = f"SDO - PUNIÇÕES EM ABERTO NO PODIO NO DIA {diaText}/{mesText}"

    if planilha.find(textoDia) is None:
        print("Texto inicial não encontrado, inserindo")
        planilha.merge_cells(area)
        planilha.format(area, cell_format)
        planilha.update_acell(area.split(":")[0], textoDia)

        modelo_texto = ["Nº", "DATA DA ENTREGA", "MEMBRO", "ATIVIDADE", "PUNIÇÃO"]
        planilha.format(f"{colunaInicio}{linha}" + ":" + f"{colunaEnd}{linha}", cell_format)

        for i, col in enumerate(range(ord(colunaInicio), ord(colunaEnd) + 1)):
            texto = modelo_texto[i]
            planilha.update_acell(f"{chr(col)}{linha}", texto)
        print("Texto  inicial inserido")

        for i in range(0, 5):
            planilha.format(f"{chr(ord(colunaInicio) + i)}{linha + 1}", {
                "borders":
                    {
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
                    }
            })
        return linha
    else:
        print("Texto inicial encontrado")
        celula = planilha.find(textoDia)
        return int(celula.row) + 1


# Looks for already existing ids, if none adds initial id text
def initialIdAmount(planilha: worksheet):
    linha = 50
    if planilha.find('Tarefas adicionadas') is None:
        planilha.update_acell('A' + str(linha), 'Tarefas adicionadas')
    else:
        linha = len(planilha.col_values(1))
    linha += 1
    return linha
