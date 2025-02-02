# Podio.com to Google Sheets Script

![GitHub repo size](https://img.shields.io/github/repo-size/brenluz/ConsoleGame?style=for-the-badge)
![GitHub language count](https://img.shields.io/github/languages/count/brenluz/ConsoleGame?style=for-the-badge)
![GitHub forks](https://img.shields.io/github/forks/brenluz/ConsoleGame?style=for-the-badge)
![Bitbucket open issues](https://img.shields.io/bitbucket/issues/brenluz/ConsoleGame?style=for-the-badge)
![Bitbucket open pull requests](https://img.shields.io/bitbucket/pr-raw/brenluz/ConsoleGame?style=for-the-badge)



> Esse projeto visa automatizar a passagem de informações do site: [Podio](https://podio.com/), para uma planilha do Google Sheets que armazena o sistema de punições da Eletrojr
> Esse projeto foi desenvolvido para a Eletrojr, empresa júnior de engenharia elétrica da UFBA, com o intuito de facilitar o gerenciamento do cumprimento das tarefas dos membros da empresa.

## 💻 Pré-requisitos

Se voce pretende modificar o projeto antes de começar, verifique se você atendeu aos seguintes requisitos:

- Você instalou a versão mais recente de `python`
- Você tem uma máquina `Windows/Linux/Mac`.
- Voce criou uma venv para o projeto.
- Você leu o arquivo `requirements.txt` e instalou as bibliotecas necessárias.
- Você tem uma conta no Google Cloud e criou uma api key para o Google Sheets.
- 

## 🚀 Executando o projeto

- Clone o repositório

- Crie uma venv com o comando:
```
python -m venv venv
```

- Instale as bibliotecas constadas no arquivo requirements.txt com o comando:
```
pip install -r requirements.txt
```
- Crie uma api key no google cloud e baixe o arquivo json com as credenciais, nomeando o arquivo de `credentials.json` e coloque na pasta raiz do projeto
- De acesso à planilha do Google Sheets para o e-mail que está no arquivo `credentials.json`

- Va a seu perfil no podio e crie uma api key fornecendo o local onde a aplicação irá rodar e o nome da aplicação.

- Crie um arquivo `.env` na raiz do projeto e adicione as seguintes variáveis:
```
CLIENT_ID=seu_client_id_do_podio
CLIENT_SECRET=seu_client_secret_do_podio
PODIO_USER=seu_usuario_do_podio
PODIO_PASS=sua_senha_do_podio
WORKSPACE_ID=id_do_workspace

SPREADSHEET_ID=id_da_planilha
```

- Execute o arquivo `main.py` com o comando:
```
python main.py
```

## 📫 Contribuindo para o Projeto

Para contribuir com o jogo, siga estas etapas:

1. Bifurque este repositório.
2. Crie um branch: `git checkout -b <nome_branch>`.
3. Faça suas alterações e confirme-as: `git commit -m '<mensagem_commit>'`
4. Envie para o branch original: `git push origin <nome_do_projeto> / <local>`
5. Crie a solicitação de pull.

Como alternativa, consulte a documentação do GitHub em [como criar uma solicitação pull](https://help.github.com/en/github/collaborating-with-issues-and-pull-requests/creating-a-pull-request).


