import pandas as pd
import win32com.client as win32

def send_emails_from_excel(file_path):
    # Leer el archivo Excel
    df = pd.read_excel(file_path)

    # Iterar sobre cada fila del DataFrame
    for index, row in df.iterrows():
        # Extraer el correo electrónico y el nombre del banco
        email = row['Correo']
        banco = row['Banco']

        # Crear un objeto de correo electrónico
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)

        # Configurar el asunto y el cuerpo del correo electrónico
        mail.Subject = f'Homologación - Debin Recurente - {banco}'
        mail.Body = 'Cuerpo del correo\n\nSaludos.'

        # Configurar el destinatario
        mail.To = email

        # Enviar el correo electrónico
        mail.Send()

# Llamar a la función con la ruta del archivo Excel
send_emails_from_excel('template1.xlsx')