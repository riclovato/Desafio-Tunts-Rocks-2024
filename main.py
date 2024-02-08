import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Se modificar esses escopos, exclua o arquivo token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

def main():
    creds = None

    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # Se não houver credenciais (válidas) disponíveis, permita que o usuário faça login.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "client_secret.json", SCOPES
            )
            creds = flow.run_local_server(port=0)
        # Salve as credenciais para a próxima execução
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        service = build("sheets", "v4", credentials=creds)

        # Chame a API do Sheets
        sheet = service.spreadsheets()

        # Obter valores em C4 até C27
        result = (
            sheet.values()
            .get(spreadsheetId='1isECBYDUcIMScuE28ZU2d_agCtFEbGYRn4-mqSRoX1Y', range="engenharia_de_software!C4:C27")
            .execute()
        )
        values_c = result.get("values", [])


        for i, value in enumerate(values_c, start=4):
            if value and float(value[0]) > 0.25 * 60:

                sheet.values().update(
                    spreadsheetId='1isECBYDUcIMScuE28ZU2d_agCtFEbGYRn4-mqSRoX1Y',
                    range=f"engenharia_de_software!G{i}",
                    valueInputOption="RAW",
                    body={"values": [["Reprovado por Frequência"]]}
                ).execute()
            else:
                # Se não for "Reprovado por Frequência", calcular a média e atualizar a coluna G
                result_avg = (
                    sheet.values()
                    .get(spreadsheetId='1isECBYDUcIMScuE28ZU2d_agCtFEbGYRn4-mqSRoX1Y', range=f"engenharia_de_software!D{i}:F{i}")
                    .execute()
                )
                values_avg = result_avg.get("values", [])

                # Calcula a média, arredonda para zero casas decimais e atualiza os valores na coluna G
                if values_avg and len(values_avg[0]) >= 3:  # Verifica se há pelo menos três colunas
                    average = round(sum(map(float, values_avg[0])) / len(values_avg[0]))  # Arredonda para zero casas decimais

                    # Atualiza a coluna H com zero para aqueles que não estão em "Exame Final"
                    sheet.values().update(
                        spreadsheetId='1isECBYDUcIMScuE28ZU2d_agCtFEbGYRn4-mqSRoX1Y',
                        range=f"engenharia_de_software!H{i}",
                        valueInputOption="RAW",
                        body={"values": [[0]]}
                    ).execute()

                    # Verifica se está em "Exame Final" e calcula a fórmula
                    result_situation = (
                        sheet.values()
                        .get(spreadsheetId='1isECBYDUcIMScuE28ZU2d_agCtFEbGYRn4-mqSRoX1Y', range=f"engenharia_de_software!G{i}")
                        .execute()
                    )
                    situation = result_situation.get("values", [])
                    if situation and situation[0][0] == "Exame Final":
                       # (100 - average) é a nota mínima necessária para 5 <= (m + naf)/2
                        naf = (100 - average)

                        sheet.values().update(
                            spreadsheetId='1isECBYDUcIMScuE28ZU2d_agCtFEbGYRn4-mqSRoX1Y',
                            range=f"engenharia_de_software!H{i}",
                            valueInputOption="RAW",
                            body={"values": [[naf]]}
                        ).execute()

    except HttpError as err:
        print(err)

if __name__ == "__main__":
    main()
