import pandas as pd

'''
passo a passo
1 - Importar a base de dados
1.2 - tabela de instalacao de aparelhos
2 - visualizar a minha base de dados
2.1 - acerto da base de dados - tratamento dos dados
'''

# 1.2 - tabela de instalacao de aparelhos
tabela_instalacao = pd.read_excel("teste_erictel.xlsx")



# 2 - Visualizar a base de dados
# 2.1 - acerto dos dados, mostrar todas as colunas da base
pd.set_option("display.max_columns", None)
print(tabela_instalacao)
print("debug - passou aqui")



#filtrar colunas de uma tabela utilizando o pandas
# tabela_instalacao.groupby("cod_tecnico").sum()
print("chegou aqui")



#calcular o total de aparelhos instalados por tecnico, agrupa pelo codigo do tecnico, e soma a quantidade de aparelhos
instalacao = tabela_instalacao[["cod_tecnico", "qtd_aparelhos_instalados"]].groupby("cod_tecnico").sum()
print(instalacao)
print("passou aqui")


#calcular a quantidade de produtos por loja
#ticket medio por loja faturamento / quantidade vendido por loja
#enviar o email com o relatorio semanal