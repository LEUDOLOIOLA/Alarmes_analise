import pandas as pd

dados = pd.read_excel("alarmes.xlsx", engine="openpyxl")

print(dados.info())