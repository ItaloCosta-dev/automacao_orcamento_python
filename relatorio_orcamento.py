import os
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import smtplib
import mimetypes
from email.message import EmailMessage
from dotenv import load_dotenv  

# Carregar as variáveis do arquivo .env
load_dotenv()

# Configurações
PASTA_PLANILHA = "planilhas"
PASTA_RELATORIO = "relatório"
NOME_ARQUIVO_SAIDA = "relatorio_orcamento.pdf"
EMAIL_REMETENTE = os.getenv("EMAIL_REMETENTE")
SENHA_EMAIL = os.getenv("SENHA_EMAIL")

def processor_planilha():
    arquivos = [f for f in os.listdir(PASTA_PLANILHA) if f.endswith(".xlsx") or f.endswith(".csv")]
    if not arquivos:
        print("Nenhuma planilha encontrada.")
        return None
    caminho_arquivo = os.path.join(PASTA_PLANILHA, arquivos[0])
    print(f"Processando arquivo: {caminho_arquivo}")
    df = pd.read_excel(caminho_arquivo) if arquivos[0].endswith(".xlsx") else pd.read_csv(caminho_arquivo, delimiter=";")
    df.columns = ["Descrição", "Valor"]
    df["Valor"] = df["Valor"].astype(str).str.replace(",", ".").astype(float)
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

def enviar_email(destinatario, caminho_pdf):
    msg = EmailMessage()
    msg["Subject"] = "Relatório de Orçamento Mensal"
    msg["From"] = EMAIL_REMETENTE
    msg["To"] = destinatario
    msg.set_content("Segue em anexo o relatório de orçamento mensal.")
    with open(caminho_pdf, "rb") as f:
        tipo, _ = mimetypes.guess_type(caminho_pdf)
        msg.add_attachment(f.read(), maintype="application", subtype="pdf", filename=NOME_ARQUIVO_SAIDA)
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(EMAIL_REMETENTE, SENHA_EMAIL)
            server.send_message(msg)
        print(f"E-mail enviado para {destinatario}!")
    except Exception as e:
        print(f"Erro ao enviar e-mail: {e}")

if __name__ == "__main__":
    print("Iniciando processamento da planilha...")
    df_processado = processor_planilha()
    if df_processado is not None:
        print("DataFrame gerado com sucesso!")
        caminho_pdf = gerar_pdf(df_processado)
        print(f"PDF gerado e salvo em: {os.path.abspath(caminho_pdf)}")
        email_destinatario = input("Digite o e-mail para envio do relatório: ")
        enviar_email(email_destinatario, caminho_pdf)
    else:
        print("Nenhuma planilha encontrada ou erro ao processar os dados.")
