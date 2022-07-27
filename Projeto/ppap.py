import pandas as pd

# IMPORTAR BASE DE DADOS
tabela_vendas = pd.read_excel('Vendas.xlsx')
print(tabela_vendas)

# Criar email (para transformar em tabela usar to.frame())
    #instalar o pywin32 (pip3 install pywin32)
import win32com.client as win32
outlook = win32.Dispath('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'para.email(s)'
mail.Subject = 'assunto.email'
mail.HTMLBody = f'''
<p>Prezados,</p>

segue resultado texto texto texto texto texto texto:
{tabela_vendas.to_html()}
'''

mail.Send()
print('email enviado')


"""
# VISUALIZAR BASE DE DADOS
pd.set_option('display.max_columns', None)
print(tabela_vendas)

# FATURAMENTO POR LOJA
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

# QUANTIDADE DE PRODUTOS VENDIDOS POR LOJA
quantidade = tabela_vendas[['ID Loja',  'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

# TICKET MÉDIO POR PRODUTO EM CADA LOJA
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
print(ticket_medio)
"""
