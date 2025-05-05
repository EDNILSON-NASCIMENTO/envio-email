import csv
import smtplib
from email.message import EmailMessage
from jinja2 import Template

# Configurações do Outlook
SMTP_SERVER = 'smtp.office365.com'
SMTP_PORT = 587
OUTLOOK_USER = 'notificacao@bancodeviagens.com'
OUTLOOK_PASS = 'Pon31742'

# Carrega o template HTML
with open('template.html', 'r', encoding='utf-8') as file:
    template_html = Template(file.read())

# Lê a lista de contatos
with open('lista.csv', newline='', encoding='utf-8') as csvfile:
    reader = csv.DictReader(csvfile)
    for row in reader:
        nome = row['nome']
        email = row['email']

        # Renderiza o template com dados do destinatário
        corpo_email = template_html.render(nome=nome, email=email)

        # Monta a mensagem
        msg = EmailMessage()
        msg['Subject'] = f"Olá {nome}, acesso ao sistema de viagens corporativas da CBS ADVOGADOS"
        msg['From'] = OUTLOOK_USER
        msg['To'] = email
        msg.set_content("Seu e-mail não suporta HTML.")
        msg.add_alternative(corpo_email, subtype='html')

        # Envia o e-mail
        try:
            with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
                server.starttls()
                server.login(OUTLOOK_USER, OUTLOOK_PASS)
                server.send_message(msg)
                print(f"E-mail enviado para {nome} ({email})")
        except Exception as e:
            print(f"Erro ao enviar para {email}: {e}")
