import pandas as pd
import win32com.client

# importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

# visualizar a base de dados
pd.set_option('display.max_columns', None)
print(tabela_vendas)
print('-' *50)

# faturamento por loja
faturamento = tabela_vendas[['ID Loja','Valor Final']].groupby('ID Loja').sum()
print(faturamento)
print('-' *50)

# quantidad de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja','Quantidade']].groupby('ID Loja').sum()
print(quantidade)
print('-' *50)

# ticket médio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(round(ticket_medio,2))

# enviar um email com relatorio
import win32com.client as win32
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'brunoflemos5@gmail.com'
mail.Subject = 'relatorio de Vendas por Loja'
mail.HTMLBody = f'''

<p>Prezados,</p>

<p>Segue o realatorio de vendas por cada loja como solicitado.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor final': 'R${:,.2f}'.format})}

<p>Quantidade vendida:</p>
{quantidade.to_html()}

<p>Ticket médio dos produtos em cada loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer duvida fico a disposição.</p>

<p>Cordialmente,</p>
<p>Bruno Daniel</p>

'''

mail.Send()

print('Email enviado!')

