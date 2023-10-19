import smtplib
import ssl
from email.message import EmailMessage
import os
import mimetypes
from datetime import datetime

# Dados de acesso do e-mail corporativo
email_senha = open("C:\\Users\\heinr\\Área de Trabalho\\Solutio\\Projeto SolutioPy\\passwords\\token", 'r', encoding='utf-8').read()
email_origem = 'fulano@xpi.com.br'  # Seu endereço de e-mail

# Valor de fechamento da bolsa
fechamento_bolsa = 120.058  # Valor de fechamento da bolsa definido manualmente

# Valor do CDI no mês anterior
cdi_mes_anterior = 1.57  # Valor do CDI no mês anterior definido manualmente

# Lista de destinatários e suas informações correspondentes
destinatarios = [
    {
        'nome': 'Cliente1',
        'email': 'contador1@empresa.com.br',
        'anexos': ['Caminho para o Anexo'],
        'conta': 'FULANO PARTICIPACOES LTDA',
        'rentabilidade': 0.00,
    },
    {
        'nome': 'Cliente2',
        'email': 'contador2@empresa.com.br',
        'anexos': ['Caminho para o Anexo'],
        'conta': 'CICLANO PARTICIPACOES S.A',
        'rentabilidade': 0.00,
    },
    # Adicione mais destinatários conforme necessário
]

# Configuração do servidor SMTP
smtp_server = 'smtp.office365.com'
smtp_port = 587

# Função para ler o conteúdo do arquivo de texto do corpo do e-mail
def ler_corpo_email(caminho_arquivo):
    with open(caminho_arquivo, 'r', encoding='utf-8') as file:
        return file.read()

# Função para ler o conteúdo do arquivo HTML da assinatura do e-mail
def ler_assinatura_outlook(caminho_arquivo):
    with open(caminho_arquivo, 'r', encoding='utf-8') as file:
        return file.read()

# Iterar sobre a lista de destinatários
for i, destinatario in enumerate(destinatarios):
    # Texto do e-mail
    assunto = f'Fechamento Mensal - {destinatario["conta"]}'
    saudacao = f'Olá, {destinatario["nome"]}!'
    corpo_arquivo = "C:\\Users\\heinr\\Área de Trabalho\\Solutio\\Projeto SolutioPy\\corpo_email.txt"
    corpo_existente = ler_corpo_email(corpo_arquivo)
    assinatura_arquivo = "C:\\Users\\heinr\\Área de Trabalho\\Solutio\\Projeto SolutioPy\\assinatura_outlook.html"
    assinatura_html = ler_assinatura_outlook(assinatura_arquivo)

    # Obter a data atual
    data_atual = datetime.now().strftime('%d/%m/%Y')

    # Formatar os dados da bolsa de valores
    dados_bolsa = f'Fechamento do Ibovespa no mês: {fechamento_bolsa:.3f}\n'
    dados_bolsa += f'CDI no mês: {cdi_mes_anterior:.2f}%\n'
    dados_bolsa += f'Rentabilidade da Carteira: {destinatario["rentabilidade"]:.2f}%\n'

    # Combinar o conteúdo existente do arquivo de texto com os dados da bolsa de valores
    corpo_email = corpo_existente.replace("[DADOS_BOLSA]", dados_bolsa)
    corpo_email = corpo_email.strip()  # Remover espaços em branco extras
    corpo_email = corpo_email.replace('\n', '<br>\n')  # Converter quebras de linha em tags <br>
    corpo_email += '<br>\n' + assinatura_html

    corpo = f'{saudacao}<br>\n<br>\n{corpo_email}'

    # Criar objeto de mensagem de e-mail
    mensagem = EmailMessage()
    mensagem['Subject'] = assunto
    mensagem['From'] = email_origem
    mensagem['To'] = destinatario['email']
    mensagem.set_content(corpo, subtype='html')

    # Adicionar os anexos específicos para o destinatário
    for anexo_path in destinatario['anexos']:
        anexo_arquivo = os.path.basename(anexo_path)
        mime_type, _ = mimetypes.guess_type(anexo_path)
        mime_type, mime_subtype = mime_type.split('/', 1)

        with open(anexo_path, 'rb') as ap:
            mensagem.add_attachment(ap.read(), maintype=mime_type, subtype=mime_subtype, filename=os.path.basename(anexo_path))

    # Enviar o e-mail
    contexto_ssl = ssl.create_default_context()
    with smtplib.SMTP(smtp_server, smtp_port) as smtp:
        smtp.starttls(context=contexto_ssl)
        smtp.login(email_origem, email_senha)
        smtp.send_message(mensagem)
        print(f'E-mail {i+1} enviado para {destinatario["email"]}')



