import pandas as pd
import win32com. client as win32

#importar base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

#visualizar base de dados
pd.set_option('display.max_columns', None)
print('tabela_vendas')

#faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

#quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print (quantidade)

print('-' * 50)

#ticket medio cada loja
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

#enviar email com relatorio
outlook = win32.Dispatch('outloook.application')
mail = outlook.CreateItem(0)
mail.To = 'affonsosn@gmail.com'
mail.Subject = 'Relatorio de vendas por loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o relatório de vendas por loja</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final':'R${:,.2f}'.format})}

<p>Quantidade vendida:</p>
{quantidade.to_html()}

<p>ticket médio por loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio':'R${:,.2f}'.format})}

<p>Qualquer dúvida estou a disposição.</p>

<p>Att.,</p>
<p>Affonso</p>
'''

mail.Send()
print('email enviado')