# Biblioiteca utilizada - 
# pip install pandas
# pip install openpyxl

import pandas as pd

def_principal = pd.read_excel("tabelas/tabela.xlsx", sheet_name="Principal")

print(def_principal.head(3))
