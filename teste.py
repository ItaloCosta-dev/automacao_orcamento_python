import os
import pandas as pd

PASTA_PLANILHA = "./planilhas"

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

# Teste rápido
resultado = processor_planilha()
if resultado is not None:
    print("Resultado do processamento:")
    print(resultado)
