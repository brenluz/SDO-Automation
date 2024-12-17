import os
import requests
from dotenv import load_dotenv

client_id = os.getenv("CLIENT_ID_PROD")
client_secret = os.getenv("CLIENT_SECRET_PROD")
username = os.getenv("USER")
password = os.getenv("PASSWORD")
space_id = os.getenv("WORKSPACE")
podio_domain = "https://api.podio.com"


def authenticate():
    url = "https://api.podio.com/oauth/token"
    payload = {
        'grant_type': 'password',
        'client_id': client_id,
        'client_secret': client_secret,
        'username': username,
        'password': password
    }
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    response = requests.post(url, data=payload, headers=headers)
    response.raise_for_status()
    return response.json()['access_token']


def get_tasks_in_space():
    access_token = authenticate()
    url = f"{podio_domain}/task/space/{space_id}/summary"
    headers = {'Authorization': f'OAuth2 {access_token}'}
    params = {
        'limit': 50
    }
    response = requests.get(url, headers=headers, params=params)
    response.raise_for_status()
    return response.json()


class Tarefa:
    def __init__(self, taskId: str, data: str, vitima: str, nome: str, punicao: str):
        self.taskId = taskId
        self.data = data
        self.gerente = vitima
        self.atividade = "Não marcou a atividade no Podio"
        self.nome = nome
        self.punicao = punicao

    def returnAtrributes(self):
        return [self.data, self.gerente, self.atividade, self.punicao]
