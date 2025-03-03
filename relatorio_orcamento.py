import os
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import smtplib
import mimetypes
from email.message import EmailMessage

# Configurações
PASTA_PLANILHA = "../planilha/"
PASTA_RELATORIO = "../relatorios"
NOME_ARQUIVO_SAIDA = "relatorio_orcamento.pdf"
EMAIL_REMETENTE = "italosiqueiradacosta@gmail.com"
SENHA_EMAIL = "sua_app_gmail" # Configurar depois
