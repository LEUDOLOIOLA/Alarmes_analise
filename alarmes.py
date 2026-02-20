import pandas as pd

dados = pd.read_excel("alarmws.xlsx", engine="openpyxl")

print(dados.info())