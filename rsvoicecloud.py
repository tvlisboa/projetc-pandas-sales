import pandas as pd

#passo 1 - importar base de dados
tabela_instalacao = pd.read_excel('teste_erictel.xlsx')

#passo 2 - visualizar a base de dados
pd.set_option('display.max_columns', None)
print(tabela_instalacao)

#quantidade total de aparelhos instalados
#a anotacao .sum(numeric_only=True) instrui ao pandas ignorar as strings, usando somente numerics.
tabela_total_aparelhos = tabela_instalacao.drop(columns='patrimonio_aparelho').groupby(['cod_tecnico']).sum(numeric_only=True)
print(tabela_total_aparelhos)

#quantidade total instalado por tecnico
quantidade_instalado = tabela_instalacao[['cod_tecnico', 'qtd_aparelhos_instalados']]


#ticket medio de instalacao semanal por tecnico

#quantidade descriminada por telefones instalados

#enviar relatorio por email
print("Chegou ate aqui")
