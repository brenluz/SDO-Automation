import os
from datetime import datetime
from dotenv import load_dotenv

from gspread import WorksheetNotFound, Worksheet

import sheets
import podio
from authSheets import getSheet
from podio import Tarefa

load_dotenv()


def main():
    urlSDO = os.getenv("teste")
    sdo = getSheet(urlSDO)
    data = datetime.now()
    ano, mes, dia = str(data).split(" ")[0].split("-")

    try:
        planilhaMes: Worksheet = sdo.worksheet(sheets.meses[int(mes)])
    except WorksheetNotFound as e:
        print("Nome da pagina incorreto, verifique se a planilha do mes existe")
        print(e)
        exit(1)

    diaAtual = int(dia)
    area = sheets.findArea(planilhaMes, diaAtual)
    colunaInicio = sheets.get_column_for_day(diaAtual)
    colunaEnd = chr(ord(colunaInicio) + 4)
    linhaDia = int(area.split(":")[0][1:])

    textoDia = "SDO - PUNIÇÕES EM ABERTO NO PODIO NO DIA " + dia + "/" + mes

    print(planilhaMes.find(textoDia))

    if planilhaMes.find(textoDia) is None:
        sheets.addInitialText(planilhaMes, diaAtual, mes, area, colunaInicio, colunaEnd, linhaDia + 1, textoDia)
        linhaDia = linhaDia + 2  # Pula as duas linhas do texto inicial
    else:
        linhaDia = linhaDia - 3  # Se o texto inicial ja existe a linhaDia e a primeira linha vazia
        print("Texto inicial encontrado")

    linhaId = 50
    if planilhaMes.find('Tarefas adicionadas') is None:
        planilhaMes.update_acell('A' + str(linhaId), 'Tarefas adicionadas')
    linhaId += 1

    sumario_tarefas = podio.get_tasks_in_space()
    tarefas = []

    for task in sumario_tarefas['overdue']['tasks']:
        nome_tarefa = task['text']
        task_id = task['task_id']
        due_date = task['due_on']
        tarefaAno, tarefaMes, tarefaDia = due_date.split(" ")[0].split("-")
        tarefa_data = f"{tarefaDia}/{tarefaMes}"
        vitima = task['responsible']['name']
        tarefas.append(Tarefa(task_id, tarefa_data, vitima, nome_tarefa, "Atenção"))

    addedTarefas = 1
    for tarefa in tarefas:
        contentUpdate = tarefa.returnAtrributes()
        contentUpdate.insert(0, addedTarefas)

        if planilhaMes.cell(linhaId, 1).value != str(tarefa.taskId):
            if tarefa.data == f"{diaAtual}/{mes}":
                taskLinha = int(linhaDia) + addedTarefas
                cell_num = planilhaMes.cell(taskLinha, ord(colunaInicio) - ord('A') + 1)
                if not cell_num.value or cell_num.value == '0':
                    planilhaMes.format(f"{colunaInicio}{taskLinha}" + ":" + f"{colunaEnd}{taskLinha}", {
                        "borders": {
                            "top": {"style": "SOLID"},
                            "bottom": {"style": "SOLID"},
                            "left": {"style": "SOLID"},
                            "right": {"style": "SOLID"}
                        },
                        "textFormat": {
                            "foregroundColor": {"blue": 0.0, "green": 0.0, "red": 0.0}},
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE"
                    })

                    planilhaMes.update_acell('A' + str(linhaId), tarefa.taskId)
                    linhaId += 1
                    for i, col in enumerate(range(ord(colunaInicio), ord(colunaEnd) + 1)):
                        coluna = chr(col)
                        planilhaMes.update_acell(f"{coluna}{taskLinha}", contentUpdate[i])
                    addedTarefas += 1
