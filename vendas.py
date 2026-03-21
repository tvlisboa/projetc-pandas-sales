import pandas as pd
import win32com.client as win32

'''
Passo a passo do workflow:
1 - Importar a base de dados
2 - Faturamento por lojas
3 - Quantidade de produtos vendidos por loja
4 - Ticket médio por produto em cada loja
5 - Fechamento de relatório e envio por email
'''

print('\n')
print("*" * 25, "TABELA DE VENDAS", "*" * 25)
print("*" * 68)
print('\n')

# 1 - Importar a base de dados
pd.set_option('display.max_columns', None)
tabela_vendas = pd.read_excel("vendas.xlsx")
print(tabela_vendas)

# 2 - Faturamento por loja
print("*" * 68)
print("Faturamento por LOJAS")
print('\n')
faturamento = tabela_vendas[["id_loja", "valor_unitario", "valor_final", "quantidade_vendida"]].groupby("id_loja").sum()
print(faturamento)

# 3 - Quantidade de produtos vendidos por loja
print("*" * 68)
print("Quantidade de produtos vendidos por LOJAS")
print('\n')
faturamento_lojas = tabela_vendas[["id_loja", "quantidade_vendida"]].groupby("id_loja").sum()
print(faturamento_lojas)

# 4 - Ticket médio por produto em cada loja
print("*" * 68)
print("Ticket médio por produto em lojas")
print('\n')
ticket_medio = (faturamento["valor_final"] / faturamento_lojas["quantidade_vendida"]).to_frame("ticket_medio")
ticket_medio = ticket_medio.rename(columns={0: "ticket_medio"})
print(ticket_medio)

print('\n')
print("*" * 25, "FECHAMENTO DE SISTEMA", "*" * 25)
print("*" * 73)
print('\n')

# 5 - Envio do relatório por email
outlook = win32.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)
mail.To = "{email - usuario ; email - usuario}"
mail.Subject = "{Assunto do email - relatorio}"
mail.HTMLBody = f'''
<h1>Prezados,</h1>

<p>Segue em anexo o Relatório de <b>Faturamento Semanal</b> de Cada Loja do Grupo</p>

<p><b>Faturamento:</b></p>
{faturamento.to_html(formatters={"valor_final": "R${:,.2f}".format()})}

<p><b>Quantidade de produtos vendidos por lojas:</b></p>
{faturamento_lojas.to_html()}

<p><b>Ticket Médio por produto em lojas:</b></p>
{ticket_medio.to_html(formatters={"ticket_medio": "R${:,.2f}".format()})}

<h3>Qualquer dúvida, estou à disposição.</h3>
<h3>Att - Gerência.</h3>
'''

mail.Send()
##confirmacao de email enviado
print("E-mail enviado com sucesso!")