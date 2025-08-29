"""Trabalhando com PDFs
Objetivo do dia

Criar um PDF fictício (exemplo de fatura) com reportlab

Ler PDF com PyPDF2 (estrutura, páginas)

Extrair texto com pdfplumber (conteúdo real)"""

import pandas as pd
import pdfplumber
from PyPDF2 import PdfReader
from reportlab.pdfgen import canvas
# Criar um PDF fictício ----------------------------
# Criar PDF
"""nome_pdf = "fatura_exemplo.pdf"
c = canvas.Canvas(nome_pdf)

# cabeçalho

c.setFont("Helvetica-Bold", 16)
c.drawString(200, 800, "FATURA DE COMPRA")

# dados da fatura
c.setFont("Helvetica", 12)
c.drawString(50, 750, "Cliente: Maria Silva")
c.drawString(50, 730, "Data: 29/08/2025")
c.drawString(50, 710, "Produto: Notebook")
c.drawString(50, 690, "Quantidade: 1")
c.drawString(50, 670, "Preço Unitario: R$3.500,00")
c.drawString(50, 650, "Total: R$ 3.500,00")

c.save()
print(f"PDF criado:{nome_pdf}")

# Ler PDF com PyPDF2-----------------------------------
reader = PdfReader("fatura_exemplo.pdf")
print(f"Numero de paginas: {len(reader.pages)}")

# ler conteudo da primeira pagina
pagina = reader.pages[0]
print(pagina.extract_text())

# Extrair texto com pdfplumber------------------
with pdfplumber.open("fatura_exemplo.pdf") as pdf:
    pagina = pdf.pages[0]
    texto = pagina.extract_text()
    print("conteudo do PDF:")
    print(texto)
"""

# Mini-Projeto do Dia – Fatura em PDF---------------
# criar PDF
nome_pdf = "fatura_exemplo.pdf"
c = canvas.Canvas(nome_pdf)

# Cabeçalho
c.setFont("Helvetica-Bold", 16)
c.drawString(200, 800, "FATURA DE COMPRA")

# Dados
c.setFont("Helvetica", 12)
c.drawString(50, 750, "Cliente: Maria Silva")
c.drawString(50, 730, "Data: 29/08/2025")
c.drawString(50, 710, "Produto: Notebook")
c.drawString(50, 690, "Quantidade: 1")
c.drawString(50, 670, "Preço Unitário: R$ 3.500,00")
c.drawString(50, 650, "Total: R$ 3.500,00")

c.save()
print(f" PDF criado: {nome_pdf}")

# ler PDF

reader = PdfReader("fatura_exemplo.pdf")
print(f"Número de páginas: {len(reader.pages)}")

pagina = reader.pages[0]
print(" Conteúdo (PyPDF2):")
print(pagina.extract_text())

# extrair texto

with pdfplumber.open("fatura_exemplo.pdf") as pdf:
    pagina = pdf.pages[0]
    texto = pagina.extract_text()
    print(" Conteúdo (pdfplumber):")
    print(texto)

# salvar em um DataFrame

# Separar linhas do PDF
linhas = texto.split("\n")

# Criar dicionário com pares chave:valor
dados = {}
for linha in linhas:
    if ":" in linha:
        chave, valor = linha.split(":", 1)
        dados[chave.strip()] = valor.strip()

# Transformar em DataFrame
df = pd.DataFrame([dados])
print(" DataFrame gerado:")
print(df)

# Salvar em Excel
df.to_excel("fatura_extraida.xlsx", index=False)
print(" Fatura salva em fatura_extraida.xlsx")
