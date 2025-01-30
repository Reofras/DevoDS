import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from io import BytesIO

# Função de autenticação e acesso à planilha
def acessar_planilha():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name('delta-era-447715-h6-d2cb872e89f3.json', scope)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_url('https://docs.google.com/spreadsheets/d/1D-H85Q3i9WfVFXGe1spjBY993ilfyLdG3uTzUt88duc/edit?gid=0#gid=0')
    return spreadsheet.worksheet('Packlist')

# Função para preencher a Packlist
def preencher_packlist(dados):
    packlist_worksheet = acessar_planilha()
    for i, dado in enumerate(dados):
        row = [
            dado['Número da NFe'],
            dado['Part Number'],
            dado['Volume'],
            dado['Call Number'],
            dado['Condição da Peça']
        ]
        packlist_worksheet.insert_row(row, 14 + i)

# Função para exportar a planilha para Excel e enviar por e-mail
def exportar_e_enviar_email():
    # Acessa a planilha e a converte em DataFrame
    packlist_worksheet = acessar_planilha()
    data = packlist_worksheet.get_all_records()
    df = pd.DataFrame(data)

    # Salva o DataFrame em um arquivo Excel em memória
    file_stream = BytesIO()
    df.to_excel(file_stream, index=False)
    file_stream.seek(0)

    # Envia o e-mail com a planilha anexada
    from_email = "davisneider2@gmail.com"
    to_email = "davisneider@matec.com"
    subject = "Planilha Packlist"
    body = "Segue em anexo a planilha Packlist."

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    # Anexar o arquivo Excel em memória
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(file_stream.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f"attachment; filename=packlist.xlsx")
    msg.attach(part)

    # Enviar o e-mail
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(from_email, "Sneider9540*")  # Use uma senha de aplicativo
        server.sendmail(from_email, to_email, msg.as_string())
        server.quit()
        print("E-mail enviado com sucesso!")
    except Exception as e:
        print(f"Erro ao enviar e-mail: {e}")

# Exemplo de dados
dados = [
    {'Número da NFe': '934543', 'Part Number': '3JK6G', 'Volume': 5, 'Call Number': '457029324', 'Condição da Peça': 'USADA'},
    {'Número da NFe': '1003434', 'Part Number': '0K1OP', 'Volume': 3, 'Call Number': '457029325', 'Condição da Peça': 'NOVA'}
]

# Preencher a planilha e enviar o e-mail
preencher_packlist(dados)
exportar_e_enviar_email()
