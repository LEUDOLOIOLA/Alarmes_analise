import pandas as pd
import streamlit as st

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

st.title("Análise de Alarmes")

try:
	df = pd.read_excel("alarmes.xlsx", engine="openpyxl")
except FileNotFoundError:
	st.error("Arquivo 'alarmes.xlsx' não encontrado. Verifique se o arquivo está no diretório correto.")
	st.stop()
except Exception as e:
	st.error(f"Erro ao ler o arquivo: {e}")
	st.stop()

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

st.header("Dados Completos")
st.dataframe(df, use_container_width=True)

st.header("Alarmes por Tipo")
if "ALARME" not in df.columns:
	st.warning("Coluna 'ALARME' não encontrada nos dados.")
else:
	contagem_alarmes = df["ALARME"].value_counts().reset_index()
	contagem_alarmes.columns = ["ALARME", "QUANTIDADE"]
	st.bar_chart(contagem_alarmes.set_index("ALARME"))

if df["nome_dia"].any():
	st.header("Alarmes por Dia da Semana")
	contagem_dia = df["nome_dia"].value_counts().reindex(dias_semana).dropna().reset_index()
	contagem_dia.columns = ["DIA", "QUANTIDADE"]
	st.bar_chart(contagem_dia.set_index("DIA"))

if df["nome_mes"].any():
	st.header("Alarmes por Mês")
	contagem_mes = df["nome_mes"].value_counts().reindex(meses).dropna().reset_index()
	contagem_mes.columns = ["MÊS", "QUANTIDADE"]
	st.bar_chart(contagem_mes.set_index("MÊS"))

st.header("Resumo de Travamentos")
filtro_alarmes = ["Travamento BBC-UTR01", "Travamento BBC-UTR02"]
if "ALARME" not in df.columns or "DURAÇÃO" not in df.columns:
	st.warning("Colunas 'ALARME' ou 'DURAÇÃO' não encontradas nos dados.")
else:
	resumo = df[df["ALARME"].isin(filtro_alarmes)].copy()
	resumo["DURAÇÃO"] = pd.to_timedelta(resumo["DURAÇÃO"], errors="coerce")
	resumo = resumo.groupby("ALARME", as_index=False)["DURAÇÃO"].sum()
	resumo["DURAÇÃO"] = resumo["DURAÇÃO"].astype(str)
	st.dataframe(resumo, use_container_width=True)
