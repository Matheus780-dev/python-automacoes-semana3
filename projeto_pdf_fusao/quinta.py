"""Projeto: Fusão e Leitura de PDFs:
Unir múltiplos PDFs em um único arquivo (PyPDF2.PdfMerger)

Ler todos os PDFs de uma pasta

Extrair os dados de cada fatura

Consolidar tudo em um único Excel"""

from datetime import datetime
import pandas as pd
import pdfplumber
import re
from PyPDF2 import PdfMerger
from pathlib import Path
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4

# 01 Gerar várias faturas PDF

# pasta de trabalho
pasta = Path("faturas")
pasta.mkdir(exist_ok=True)

# funçao para formatar moeda BR no texto do PDF


def brl(v):
    s = f"R$ {v:,.2f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


# função para criar um PDF de fatura
def criar_fatura_pdf(destino, cliente, data, produto, quantidade, preco_unit):
    total = quantidade * preco_unit
    c = canvas.Canvas(str(destino), pagesize=A4)

    # cabeçalho
    c.setFont("Helvetica-Bold", 16)
    c.drawString(180, 800, "FATURA DE COMPRA")

    # dados
    c.setFont("Helvetica", 12)
    y = 750
    linhas = [
        f"Cliente: {cliente}",
        f"Data: {data}",
        f"Produto: {produto}",
        f"Quantidade: {quantidade}",
        f"Preço Unitário: {brl(preco_unit)}",
        f"Total: {brl(total)}",
    ]
    for linha in linhas:
        c.drawString(50, y, linha)
        y -= 20

    c.save()


# dados de exemplo
faturas = [
    {"cliente": "Maria Silva", "data": "30/08/2025",
        "produto": "Notebook", "quantidade": 1, "preco": 3500.00},
    {"cliente": "João Souza",  "data": "30/08/2025",
        "produto": "Mouse",    "quantidade": 2, "preco":   80.00},
    {"cliente": "Ana Lima",    "data": "30/08/2025",
        "produto": "Teclado",  "quantidade": 1, "preco":  120.00},
    {"cliente": "Pedro Alves", "data": "30/08/2025",
        "produto": "Monitor",  "quantidade": 1, "preco":  900.00},
    {"cliente": "Carla Mendes", "data": "30/08/2025",
        "produto": "Headset",  "quantidade": 3, "preco":  150.00},
]

# criar os PDFs
paths_pdfs = []
for i, f in enumerate(faturas, start=1):
    nome = f"fatura_{i}_{f['cliente'].split()[0].lower()}.pdf"
    destino = pasta / nome
    criar_fatura_pdf(destino, f["cliente"], f["data"],
                     f["produto"], f["quantidade"], f["preco"])
    paths_pdfs.append(destino)


print("faturas criadas em:", pasta.resolve())
for p in paths_pdfs:
    print(" -", p.name)

# 02 Mesclar todas as faturas em um único PDF
saida_merged = pasta / "faturas_merged.pdf"
merger = PdfMerger()
for p in sorted(paths_pdfs):
    merger.append(str(p))
merger.write(str(saida_merged))
merger.close()

print("PDF mesclado criado", saida_merged.name)

# 03 Extrair texto e estruturar os dados


def analizar_moeda_br(s: str):
    if s is None:
        return None
    s = s.replace("R$", "").replace(" ", "")
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except:
        return None


def analizar_int(s: str):
    if s is None:
        return None
    m = re.search(r"\d+", s)
    return int(m.group()) if m else None


def analizar_data_br(s: str):
    try:
        return datetime.strptime(s.strip(), "%d/%m/%Y").date()
    except:
        return None


registros = []

for pdf_path in paths_pdfs:
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        texto = page.extract_text() or ""

    # Mapa chave:valor a partir das linhas "Campo: Valor"
    dados = {}
    for linha in texto.splitlines():
        if ":" in linha:
            chave, valor = linha.split(":", 1)
            dados[chave.strip()] = valor.strip()

    registro = {
        "Arquivo": pdf_path.name,
        "Cliente": dados.get("Cliente"),
        "Data": analizar_data_br(dados.get("Data", "")),
        "Produto": dados.get("Produto"),
        "Quantidade": analizar_int(dados.get("Quantidade", "")),
        "Preço Unitário (BRL)": analizar_moeda_br(dados.get("Preço Unitário",
                                                            "")),
        "Total (BRL)": analizar_moeda_br(dados.get("Total", "")),
    }

    # Recalcular total para conferência
    if registro["Quantidade"] is not None and registro["Preço Unitário (BRL)"]\
            is not None:
        registro["Total Calculado (BRL)"] = registro["Quantidade"] * \
            registro["Preço Unitário (BRL)"]
    else:
        registro["Total Calculado (BRL)"] = None

    # Checagem simples (tolerância para arredondamento)
    if registro["Total (BRL)"] is not None and registro["Total Calculado"
                                                        " (BRL)"] is not None:
        registro["Total OK?"] = abs(
            registro["Total (BRL)"] - registro["Total Calculado (BRL)"]) < 0.01
    else:
        registro["Total OK?"] = False

    registros.append(registro)

df = pd.DataFrame(registros)
print(" Extração concluída. Registros lidos:", len(df))
print(df)

# 04 Consolidar em Excel

arquivo_excel = pasta / \
    f"consolidado_faturas_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.xlsx"

resumo_cliente = (
    df.groupby("Cliente", as_index=False)[
        ["Total (BRL)", "Total Calculado (BRL)"]].sum()
)

resumo_produto = (
    df.groupby("Produto", as_index=False)[
        ["Total (BRL)", "Total Calculado (BRL)"]].sum()
)

with pd.ExcelWriter(arquivo_excel, engine="openpyxl") as w:
    df.to_excel(w, sheet_name="Consolidado", index=False)
    resumo_cliente.to_excel(w, sheet_name="Resumo por Cliente", index=False)
    resumo_produto.to_excel(w, sheet_name="Resumo por Produto", index=False)

print("excel gerado:", arquivo_excel.name)
