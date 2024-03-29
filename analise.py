# Biblioiteca utilizada - 
# pip install pandas
# pip install openpyxl
# graficos - pip install plotly

import pandas as pd
import plotly.express as px

# Definindo para nao apresentar numeros grandes como notação cientifica
pd.options.display.float_format = '{:.2f}'.format

#importando cada aba da planilha para uma variavel
df_principal = pd.read_excel("tabelas/tabela.xlsx", sheet_name="Principal")
df_total_de_acoes = pd.read_excel("tabelas/tabela.xlsx", sheet_name="Total_de_acoes")
df_ticker = pd.read_excel("tabelas/tabela.xlsx", sheet_name="Ticker")
df_GPT = pd.read_excel("tabelas/tabela.xlsx", sheet_name="GPT")



#selecionando as colunas que eu quero manter
# copio as colunas desejadas da df_principal
df_principal = df_principal[['Ativo','Data','Ultimo (R$)','Var. Dia (%)']].copy()

# Renomear colunas
df_principal = df_principal.rename(columns={'Ultimo (R$)':'valor_final', 'Var. Dia (%)':'var_dia_pct'}).copy()
#print(df_principal)

# Criando uma nova coluna com a porcentagem de var_dia_pct
df_principal['var_pct'] = df_principal['var_dia_pct'] / 100
df_principal['valor_inicial'] = df_principal['valor_final'] / (df_principal['var_pct'] + 1)

#usando o procv
#.merge(aba onde procurar, coluna1 , coluna2, coluna base para procurar na outra coluna)
df_principal = df_principal.merge(df_total_de_acoes, left_on='Ativo', right_on='Código', how='left')

#removendo uma coluna
df_principal = df_principal.drop(columns=['Código'])

# Mais operaçoes
df_principal['variacao_rs'] = (df_principal['valor_final'] - df_principal['valor_inicial']) * df_principal['Qtde. Teórica']
df_principal = df_principal.rename(columns={'Qtde. Teórica': 'qtd_teorica'}).copy()


# O método apply() em Pandas é usado para aplicar uma função ao longo de um eixo do DataFrame ou de uma série. 
# Quando aplicado a uma série, ele aplica a função a cada elemento da série. Quando aplicado a um DataFrame, 
# ele aplica a função a cada coluna (ou linha, dependendo do eixo especificado) do DataFrame.

#df_principal['Resultado'] = df_principal['Variacao_rs'].apply(lambda x: 'Subiu' if x > 0 else ('Desceu' if x < 0 else 'Estável'))
def define_resultado(variacao):
    if variacao > 0:
        return 'Subiu'
    elif variacao < 0:
        return 'Desceu'
    else:
        return 'Estável'

df_principal['Resultado'] = df_principal['variacao_rs'].apply(define_resultado)

# Procv
df_principal = df_principal.merge(df_ticker, left_on='Ativo', right_on='Ticker', how='left')
# removendo uma coluna
df_principal = df_principal.drop(columns=['Ticker'])

df_principal = df_principal.merge(df_GPT, left_on='Nome', right_on='Empresa', how='left')
# removendo uma coluna
df_principal = df_principal.drop(columns=['Empresa'])
# Preenchendo valores ausentes com zero na coluna 'Idade'
df_principal['Idade'] = df_principal['Idade'].fillna(0).astype(int)

df_principal['Cat_idade'] = df_principal['Idade'].apply(lambda x: 'Mais de 100' if x > 100 else ('Menos de 50' if x < 50 else 'Entre 50 e 100'))

# Claculos

# Calculando o maior valor
maior = df_principal['variacao_rs'].max()

# Calculando o menor valor
menor = df_principal['variacao_rs'].min()

# Calculando a média
media = df_principal['variacao_rs'].mean()

# Calculando a média de quem subiu
media_subiu = df_principal[df_principal['Resultado'] == 'Subiu']['variacao_rs'].mean()

# Calculando a média de quem desceu
media_desceu = df_principal[df_principal['Resultado'] == 'Desceu']['variacao_rs'].mean()

# Imprimindo os resultados
print(f"Maior\tR$ {maior:,.2f}")
print(f"Menor\tR$ {menor:,.2f}")
print(f"Media\tR$ {media:,.2f}")
print(f"Media de quem subiu\tR$ {media_subiu:,.2f}")
print(f"Media de quem desceu\tR$ {media_desceu:,.2f}")

df_principal_subiu = df_principal[df_principal['Resultado'] == 'Subiu']

# grupby(O que vc vai agrupar)['O que vamos somar'].sum.//nao transforma em um vetor e mantem a estrutura de dataframe
df_analise_segmento = df_principal_subiu.groupby('Setor')['variacao_rs'].sum().reset_index()

df_analise_saldo = df_principal.groupby('Resultado')['variacao_rs'].sum().reset_index()

# Mostrando graficos
fig = px.bar(df_analise_saldo, x='Resultado', y='variacao_rs', text='variacao_rs', title='Variação Reais por Resultado')
#fig.show()

# Exportar o DataFrame para um arquivo Excel
df_principal.to_excel('nome_do_arquivo.xlsx', index=False)