import pandas as pd
import win32com.client as win32
import smtplib
import email.message

# importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

# visualizar a base de dados
pd.set_option('display.max_columns', None)
print(tabela_vendas)

# Faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

# Quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

print('-' * 50)
# Ticket médio por produto em cada loja
tickt_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
print(tickt_medio)

# enviar um email com relatorio
def enviar_email():
    corpo_email = f"""
<p>Prezados,</p>

<p>Segue o Relatório de vendas por cada Loja</p>

<p>Faturamanto:</p>
{faturamento.to_html()}

<p>Quantidade vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio dos Produtos de  cada LOja</p>
{tickt_medio.to_html()}

<p>Qualquer dúvida estou a disposição.</p>

<p>Att,</p>
<p>Ricardo e Emille Gerente Geral</p>
    """

    msg = email.message.Message()
    msg['Subject'] = "Testando meu Robô"
    msg['From'] = 'www.ttw.dev@gmail.com'
    msg['To'] = 'www.ttw.dev@gmail.com'
    password = 'rbhxonzhfudsbodw'
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email)

    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    # Login Credentials for sending the mail
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print('Email enviado')

enviar_email()


