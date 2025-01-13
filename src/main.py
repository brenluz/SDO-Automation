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
    # pega a data atual
    data = datetime.now()
    ano, mes, dia = str(data).split(" ")[0].split("-")
    diaAtual = int(dia)
    # procura o url da planilha
    urlSDO = os.getenv("teste")
    sdo = getSheet(urlSDO)

    try:
        planilhaMes: Worksheet = sdo.worksheet(sheets.meses[int(mes)])
    except WorksheetNotFound as e:
        print("Nome da pagina incorreto, verifique se a planilha do mes existe")
        print(e)
        exit(1)

    #  Define a area da planilha que serao inseridas as informacoes
    area = sheets.findArea(planilhaMes, diaAtual)
    colunaInicio = sheets.get_column_for_day(diaAtual)
    colunaEnd = chr(ord(colunaInicio) + 4)
    linhaDia = int(area.split(":")[0][1:])

    # Adiciona o texto inicial na planilha
    linhaDia = sheets.addInitialText(planilhaMes, diaAtual, mes, area, colunaInicio, colunaEnd, linhaDia + 1)
    linhaDia += 2

    # Pega as tarefas do podio
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

    # Pega os ids das tarefas ja adicionadas
    linhaId = sheets.initialIdAmount(planilhaMes)
    ids = planilhaMes.get('A' + str(50) + ':A')

    addedTarefas = 1

    for tarefa in tarefas:
        taskNotAdded = True
        for i in ids:
            if i[0] == str(tarefa.taskId):
                taskNotAdded = False
                print('task already added')
                break

        if taskNotAdded:
            stringDia = sheets.add_leading_zero(str(diaAtual))
            if tarefa.data == f"{stringDia}/{mes}":
                print('date is right')
                tarefaNotchecked = True

                contentUpdate = tarefa.returnAtrributes()
                contentUpdate.insert(0, addedTarefas)

                while tarefaNotchecked:
                    taskLinha = int(linhaDia) + addedTarefas - 2
                    contentUpdate[0] = addedTarefas
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
                        for i, col in enumerate(range(ord(colunaInicio), ord(colunaEnd) + 1)):
                            coluna = chr(col)
                            planilhaMes.update_acell(f"{coluna}{taskLinha}", contentUpdate[i])
                        addedTarefas += 1
                        linhaId += 1
                        tarefaNotchecked = False
                    else:
                        addedTarefas += 1


main()
