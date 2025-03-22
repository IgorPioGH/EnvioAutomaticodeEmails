import win32com.client as win32
import pandas as pd


df = pd.read_csv('Nome_email.csv', sep=";")

def enviar_email(nome, email_remetente):
    # Criar integracao com Outlook
    outlook = win32.Dispatch('outlook.application')

    # Criar e-mail
    email = outlook.CreateItem(0)

    # Configurar informacoes de e-mail
    email.To = email_remetente
    email.Subject = 'Automático Python'
    # email.Body
    email.HTMLBody = f"""
    <html>
      <body>
        <p>Olá {nome},</p>
        <p>Eu sou um e-mail automático.</p>
        <p>Não liga pra mim não <br> tô só de passagem.</p>
        <p>Att,<br>Igor</p>
      </body>
    </html>
    """

    # Enviar email
    email.Send()
    print(f'E-mail enviado para {nome} no email: {email_remetente}')

for row in df.itertuples():
    enviar_email(row.Nome, row.Email)