import pandas as pd

#passo 1 - importar base de dados
tabela_instalacao = pd.read_excel('teste_erictel.xlsx')

#passo 2 - visualizar a base de dados
pd.set_option('display.max_columns', None)
print(tabela_instalacao)

#quantidade total de aparelhos instalados
tabela_total_aparelhos = tabela_instalacao.groupby("cod_tecnico").sum()
print(tabela_total_aparelhos)

#quantidade total por tecnico

#ticket medio de instalacao semanal por tecnico

#quantidade descriminada por telefones instalados

#enviar relatorio por email



print("26 de marco de 2026")
