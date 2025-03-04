import os
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import smtplib
import mimetypes
from email.message import EmailMessage

# Configurações
PASTA_PLANILHA = "planilhas"
PASTA_RELATORIO = "relatório"
NOME_ARQUIVO_SAIDA = "relatorio_orcamento.pdf"
EMAIL_REMETENTE = "italosiqueiradacosta@gmail.com"
SENHA_EMAIL = "sua_app_gmail"  # Configurar depois

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

def gerar_pdf(df):
    if not os.path.exists(PASTA_RELATORIO):
        os.makedirs(PASTA_RELATORIO)

    caminho_pdf = os.path.join(PASTA_RELATORIO, NOME_ARQUIVO_SAIDA)

    c = canvas.Canvas(caminho_pdf, pagesize=A4)
    c.setFont("Helvetica", 12)

    c.drawString(200, 800, "Relatório de orçamento mensal.")
    c.drawString(50, 780, "Descrição")
    c.drawString(400, 780, "Valor")
    y = 760

    for _, row in df.iterrows():
        c.drawString(50, y, row["Descrição"])
        c.drawString(400, y, f"R${row['Valor']:.2f}")
        y -= 20

    c.save()
    print(f"Relatório salvo em: {os.path.abspath(caminho_pdf)}")
    return caminho_pdf

# Execução do script
print("Iniciando processamento da planilha...")
df = processor_planilha()

if df is not None:
    print("DataFrame gerado com sucesso!")
    caminho_pdf = gerar_pdf(df)
    print(f"PDF gerado e salvo em: {os.path.abspath(caminho_pdf)}")
else:
    print("Nenhuma planilha encontrada ou erro ao processar os dados.")
