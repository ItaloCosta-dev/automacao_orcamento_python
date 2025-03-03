import os
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import smtplib
import mimetypes
from email.message import EmailMessage

# Configurações
PASTA_PLANILHA = "planilhas"
PASTA_RELATORIO = "../relatorios"
NOME_ARQUIVO_SAIDA = "relatorio_orcamento.pdf"
EMAIL_REMETENTE = "italosiqueiradacosta@gmail.com"
SENHA_EMAIL = "sua_app_gmail" # Configurar depois

def processor_planilha():
    arquivos = [f for f in os.listdir(PASTA_PLANILHA) if f.endswith(".xlsx") or f.endswith(".csv")]

    if not arquivos:
        print("Nenhuma planilha encontrada.")
        return None

    caminho_arquivo = os.path.join(PASTA_PLANILHA, arquivos[0])
    print(f"Processando arquivo: {caminho_arquivo}")

    # Lendo planilha
    if arquivos[0].endswith(".xlsx"):
        df = pd.read_excel(caminho_arquivo)
    else:
        df = pd.read_csv(caminho_arquivo, delimiter=";")

    # Normalizar os valores
    df.columns = ["Descrição", "Valor"]
    df["Valor"] = df["Valor"].astype(str).str.replace(",", ".").astype(float)

    # Somar valores duplicados e ordenação
    df = df.groupby("Descrição", as_index=False).sum()
    df = df.sort_values(by="Valor", ascending=False) 

    return df       
