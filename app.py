
import smtplib
import email.message
import pandas as pd

dados = pd.read_excel(r'caminho para o arquivo', sheet_name='Nome da planilha ')


def enviar_email(nome,valor,data_vencimento,endereco_email):  

    corpo_email = f"""
    <p>Olá {nome}!</p>
    <p>A sua conta no valor de {valor} vence no dia: {data_vencimento}</p>
    <p>Link do Boleto: </p>
    <p>https://github.com/jeffilvu</p> 
    """

    # Escreva login e senha do email que irá enviar os emails(precisa permitir nas configurações de segurança do google)
    
    msg = email.message.Message()
    msg['Subject'] = "Cobrança"
    msg['From'] = 'remetente'
    msg['To'] = f'{endereco_email}' #destinatario capturado da planilha
    password = 'senha' #gerada pelo senhas de app do google
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email)

    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print('Email enviado')


#lê os dados da planilha utilizando pandas para mudar o texto do email

for indice, linha in dados.iterrows():

    nome = linha['Nome']
    valor = linha['Valor']
    data_vencimento = linha['Data de Vencimento']
    endereco_email = linha['Email']
    
    if nome == '':
        break
    
    enviar_email(nome,valor,data_vencimento,endereco_email)
    
    