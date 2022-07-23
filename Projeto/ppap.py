import pandas as pd
import smtplib
import numpy as np

# IMPORTAR BASE DE DADOS
tabela_vendas = pd.read_excel('Vendas.xlsx')

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

