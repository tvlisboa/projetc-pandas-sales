import pandas as pd


print('\n')
print("*" * 25, "TABELA DE INSTALACAO - MACAE", "*" * 13)
print("*" * 68)
print('\n')

#passo 1 - importar base de dados
tabela_instalacao = pd.read_excel('teste_erictel.xlsx')
#
#
#
#passo 2 - visualizar a base de dados
pd.set_option('display.max_columns', None)
print("Leitura da base de dados - planilha de instalacao de aparelhos")
print(tabela_instalacao)
#
#
#
#quantidade total de aparelhos instalados
#a anotacao .sum(numeric_only=True) instrui ao pandas ignorar as strings, usando somente numerics.
tabela_total_aparelhos = tabela_instalacao.drop(columns='patrimonio_aparelho').groupby(['cod_tecnico']).sum(numeric_only=True)
print("Quantidade total de aparelhos instalados")
print(tabela_total_aparelhos)
#
#
#
#quantidade total instalado por tecnico
qtd_insta_tecnico = tabela_instalacao[['cod_tecnico', 'qtd_aparelhos_instalados']].groupby('cod_tecnico').sum()
print("Quantidade de aparelhos instalados por tecnico ")
print(qtd_insta_tecnico)
#
#
#
#ticket medio de instalacao semanal por tecnico
ticket_medio = tabela_instalacao.drop(columns='patrimonio_aparelho').groupby(['cod_tecnico']).median(numeric_only=True)
print("Ticket medio por tecnico")
print(ticket_medio)
#
#
#
#quantidade descriminada por telefones instalados

#enviar relatorio por email


print('\n')
print("*" * 25, "FECHAMENTO DE SISTEMA - INSTALACAO MACAE", "*" * 25)
print("*" * 73)
print('\n')
