import pandas as pd

df = pd.read_excel("alarmes.xlsx", engine="openpyxl")

dias_semana = [
	"segunda-feira",
	"terça-feira",
	"quarta-feira",
	"quinta-feira",
	"sexta-feira",
	"sábado",
	"domingo",
]

meses = [
	"janeiro",
	"fevereiro",
	"março",
	"abril",
	"maio",
	"junho",
	"julho",
	"agosto",
	"setembro",
	"outubro",
	"novembro",
	"dezembro",
]

coluna_data = None

for coluna in df.columns:
	serie_data = pd.to_datetime(df[coluna], errors="coerce")
	if serie_data.notna().any():
		coluna_data = serie_data
		break

if coluna_data is not None:
	df["nome_dia"] = coluna_data.dt.weekday.map(lambda x: dias_semana[x] if pd.notna(x) else "")
	df["nome_mes"] = coluna_data.dt.month.map(lambda x: meses[int(x) - 1] if pd.notna(x) else "")
else:
	df["nome_dia"] = ""
	df["nome_mes"] = ""

df["criticidade"] = ""

print(df.info())

filtro_alarmes = ["Travamento BBC-UTR01", "Travamento BBC-UTR02"]
resumo = df[df["ALARME"].isin(filtro_alarmes)].copy()

resumo["DURAÇÃO"] = pd.to_timedelta(resumo["DURAÇÃO"], errors="coerce")
resumo = resumo.groupby("ALARME", as_index=False)["DURAÇÃO"].sum()
resumo["DURAÇÃO"] = resumo["DURAÇÃO"].astype(str)

arquivos_saida = ["alarmes_novo.xlsx", "alarmes_novo_1.xlsx", "alarmes_novo_2.xlsx"]
arquivo_gerado = None

for arquivo_saida in arquivos_saida:
	try:
		with pd.ExcelWriter(arquivo_saida, engine="openpyxl") as writer:
			df.to_excel(writer, sheet_name="Dados", index=False)
			resumo.to_excel(writer, sheet_name="Resumo_Travamentos", index=False)
		arquivo_gerado = arquivo_saida
		break
	except PermissionError:
		continue

if arquivo_gerado is not None:
	print(f"Planilha gerada: {arquivo_gerado}")
else:
	print("Não foi possível gerar o arquivo. Feche a planilha aberta e execute novamente.")