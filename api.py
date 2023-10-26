from __future__ import print_function

import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# SCOPES E CREDENTIALS

# Se quiser que o scope dê para ler e editar, remova o '.readonly' no fim
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
# ID do excel desejado
# Range dos campos que deseja ler
SAMPLE_SPREADSHEET_ID = '1qdnGd2oxNYMz6U6oXCkieOvLo8CFJmiTj31lYn2LXgw'
SAMPLE_RANGE_NAME = 'Sheet1!A2:C13'


def main():
    # Login
    creds = None

    # Check if file '.json' exists. This file used for save the credentials of user
    # If file exists, save the credentials using 'Credentials' class
    # If not exists, request to the user for to login with Google Acc
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:

        # After take the credentials, create a service to access Google API using 'build' function
        # Read the specific SpreadSheet using 'values().get()' function
        # That values will saved 'values'
        service = build('sheets', 'v4', credentials=creds)

        # Aqui ele faz a leitura do excel que eu quero (passando o ID especifico) e do range que eu quero
        # (passando o range especifico na variavel global acima)
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range=SAMPLE_RANGE_NAME).execute()
        values = result['values']

        valores_adicionar = [
            ['Imposto'],
        ]

        if not values:
            print('No data found.')
            return

        for linha in values:
            vendas = linha[1]
            vendas = float(vendas.replace("$", "").replace(",", ""))

            imposto = vendas * 0.10
            valores_adicionar.append([imposto])

        result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                       range="Sheet1!C1:C13",
                                       valueInputOption="USER_ENTERED",
                                       body={'values': valores_adicionar}).execute()

        # ADICIONAR / EDITAR VALORES MANUALMENTE
        # valores_adicionar = [
        #     ['Resultado'],
        #     ['True'],
        #     ['False'],
        #     ['True'],
        #     ['False'],
        #     ['True'],
        #     ['True'],
        #     ['True'],
        #     ['True'],
        #     ['True'],
        #     ['True'],
        #     ['True'],
        #     ['True'],
        # ]
        # result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
        #                                range="Sheet1!C1:C13",
        #                                valueInputOption="USER_ENTERED",
        #                                body={'values': valores_adicionar}).execute()

    except HttpError as err:
        print(err)


if __name__ == '__main__':
    main()
