import win32com.client as win32
import pandas as pd


# importar a base de dados

tabela_vendas = pd.read_excel('Vendas.xlsx')


# visualizar a base de dados

pd.set_option('display.max_columns', None)

print(tabela_vendas)


# faturamento por loja

faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby(
    'ID Loja').sum()
print(faturamento)
print('-' * 50)

# quantidade de produtos vendidos por loja

quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)
print('-' * 50)
# ticket medio por produto em cada loja

ticket_medio = (faturamento['Valor Final'] /
                quantidade['Quantidade']).to_frame()

# renomeando a coluna, nome anterior '0' novo nome 'ticket medio'
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Medio'})

print(ticket_medio)


outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'amaurypb845@gmail.com'
mail.Subject = 'Relatorio de vendas por loja'
mail.HTMLBody = f'''
Prezados, 

<p>Segue o relatorio de vendas de cada loja.</p>

<p><strong>FATURAMENTO:</strong></p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p><strong>QUANTIDADE VENDIDA:</strong></p>
{quantidade.to_html()}

<p><strong>TICKET MEDIO POR LOJA:</strong></p>
{ticket_medio.to_html()}

<p>Qualquer duvida estou a disposição.</p>
<p>Att..</p>

</p>Amaury Brito</p>
'''

mail.Send()

print('\nEmail enviado com sucesso! :)')
