"""
Aprender a ler, manipular e salvar planilhas Excel com pandas e openpyxl.
"""
import pandas as pd
from datetime import datetime

"""# criando dataset
dados = {
    "Data": [datetime(2025, 8, 1), datetime(2025, 8, 2), datetime(2025, 8, 2)],
    "Cliente": ["Maria", "João", "Ana"],
    "Produto": ["Notebook", "Mouse", "Teclado"],
    "Quantidade": [1, 2, 1],
    "Preço": [3500, 80, 120]
}

df = pd.DataFrame(dados)
df.to_excel("vendas.xlsx", index=False)"""

"""# Exercícios Práticos-----------

# ler uma planilha
df = pd.read_excel("vendas.xlsx")
print(df.head())

# adicionar uma coluna total
df["total"] = df["Quantidade"] * df["Preço"]
print(df)

# salvar em um novo arquivo
df.to_excel("relatorio_com_total.xlsx", index=False)"""
# ------------------------------------------------------------------------
# Mini-Projeto: Relatório de Vendas em Excel

# dataset fictício de vendas
dados = {
    "Data": [
        datetime(2025, 8, 1),
        datetime(2025, 8, 2),
        datetime(2025, 8, 3),
        datetime(2025, 8, 3),
        datetime(2025, 8, 4),
    ],
    "Cliente": ["Maria", "João", "Ana", "Pedro", "Carla"],
    "Produto": ["Notebbook", "Mouse", "Teclado", "Monitor", "Headset"],
    "Quantidade": [1, 2, 1, 1, 3],
    "Preço": [3500, 80, 120, 900, 150],
}

df = pd.DataFrame(dados)

# criar coluna "Total"
df["Total"] = df["Quantidade"] * df["Preço"]

# ordenar pelo maior valor de vendas
df = df.sort_values(by="Total", ascending=False)

# salvar
df.to_excel("relatorios_final.xlsx", index=False)

print("relatorio criado com sucesso: relatorio_final.xlsx")
