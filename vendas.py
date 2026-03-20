import pandas as pd

'''
passo a passo do wokflow

1 - importar a base de dados
2 - visualizar a base de dados
3 - fazer o tratamento dos meus dados - faturamento por lojas
4 - quantidade de produtos vendidos por lojas
5 - calcular o ticket medio de produto por lojas
6 - fechamento de relatorio e envio por email
'''

print("##########  TABELA DE VENDAS ##########")


#1 - importar a base de dados
pd.set_option('display.max_columns', None)
tabela_vendas = pd.read_excel("vendas.xlsx")
print(tabela_vendas)


#2 - Faturamento por loja
print("Faturamento por LOJAS")
faturamento = tabela_vendas[["id_loja","valor_unitario", "valor_final"]].groupby("id_loja").sum()
print(faturamento)

#3 - Quantidade de produtos vendidos por loja
print("Quantidade de produtos vendidos por LOJAS")
faturamento_lojas = tabela_vendas[["id_loja", "quantidade_vendida"]].groupby("id_loja").sum()
print(faturamento_lojas)


#4 - Ticket medio por produto em casa loja


#5 - Envio do relatorio por email



print("#####  FECHAMENTO DO SISTEMA - VENDAS PYTHON #####")
