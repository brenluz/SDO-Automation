import os
from datetime import datetime

from gspread import WorksheetNotFound, Worksheet

from src import sheets
from src import podio
from src.authSheets import getSheet
from src.podio import Tarefa

if __name__ == '__main__':

    urlSDO = os.getenv("teste")
    sdo = getSheet(urlSDO)
    data = datetime.now()
    ano, mes, diaAtual = str(data).split(" ")[0].split("-")

    try:
        planilhaMes: Worksheet = sdo.worksheet(sheets.meses[int(mes)])
    except WorksheetNotFound as e:
        print("Nome da pagina incorreto, verifique se a planilha do mes existe")
        print(e)
        exit(1)

    textoDia = "SDO - PUNIÇÕES EM ABERTO NO PODIO NO DIA" + diaAtual + "/" + mes
    diaAtual = int(diaAtual) - 1

    area = sheets.findArea(planilhaMes, diaAtual)
    colunaInicio = sheets.get_column_for_day(diaAtual)
    colunaEnd = chr(ord(colunaInicio) + 4)
    linhaDia = int(area.split(":")[0][1:])

    if planilhaMes.find(textoDia) is None:
        print("Texto inicial não encontrado, inserindo")

        planilhaMes.merge_cells(area)
        planilhaMes.format(area, sheets.cell_format)
        planilhaMes.update_acell(area.split(":")[0], textoDia)

        modelo_texto = ["Nº", "DATA DA ENTREGA", "MEMBRO", "ATIVIDADE", "PUNIÇÃO"]
        planilhaMes.format(f"{colunaInicio}{linhaDia + 1}" + ":" + f"{colunaEnd}{linhaDia + 1}", sheets.cell_format)
        for i, col in enumerate(range(ord(colunaInicio), ord(colunaEnd) + 1)):
            texto = modelo_texto[i]
            print("linhadia", linhaDia)
            planilhaMes.update_acell(f"{chr(col)}{linhaDia + 1}", texto)
        print("Texto  inicial inserido")
        linhaDia = linhaDia + 2
    else:
        linhaDia = linhaDia - 1

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

    addedTarefas = 0
    for tarefa in tarefas:
        print(tarefa.nome)
        print(tarefa.data)
        print(f"{diaAtual}/{mes}")

        if tarefa.data == f"{diaAtual}/{mes}":
            print("e")
            taskLinha = int(linhaDia) + addedTarefas
            cell_atividade = planilhaMes.cell(taskLinha, ord(colunaInicio) - ord('A') + 4)
            cell_membro = planilhaMes.cell(taskLinha, ord(colunaInicio) - ord('A') + 3)
            cell_id = planilhaMes.cell(taskLinha, ord(colunaInicio) - ord('A') + 1)

            if cell_atividade.value != tarefa.punicao:
                print("b")
                if cell_membro.value != tarefa.gerente:
                    print("c")
                    if not planilhaMes.cell(taskLinha, ord(colunaInicio) - ord('A') + 1).value:
                        print("d")

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

                        for i, col in enumerate(range(ord(colunaInicio), ord(colunaEnd) + 1)):
                            coluna = chr(col)
                            planilhaMes.update_acell(f"{coluna}{taskLinha}", tarefa.returnAtrributes()[i])
        addedTarefas += 1

