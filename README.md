# 🐍 Automação com Python – Semana 3 (Excel, PDFs e Relatórios)

Este repositório contém os projetos desenvolvidos durante a **Semana 3 do plano de estudos de Automação com Python + UiPath**.  
O foco desta semana foi **automação de Excel, manipulação de PDFs e consolidação de relatórios**.

---

## 📂 Projetos da Semana

### 📊 1. Relatório de Vendas em Excel
- Script: `projeto_excel_simples/segunda.py`
- Gera `relatorio_final.xlsx` com cálculo automático do campo **Total**.

---

### 📑 2. Relatório Agrupado em Excel
- Script: `projeto_excel_agrupado/terça.py`
- Gera `relatorio_YYYY-MM-DD.xlsx` com três abas:
  - **Por Cliente**
  - **Por Produto**
  - **Detalhado**
- Inclui **formatação monetária em R$**.

---

### 📄 3. Extração de Dados de PDF
- Script: `projeto_pdf_extracao/quarta.py`
- Cria `fatura_exemplo.pdf` e extrai informações para `fatura_extraida.xlsx`.

---

### 📦 4. Fusão e Consolidação de Múltiplos PDFs
- Script: `projeto_pdf_fusao/quinta.py`
- Cria várias faturas PDF em `/faturas`
- Mescla todas em `faturas_merged.pdf`
- Gera `consolidado_faturas.xlsx` com:
  - **Consolidado**
  - **Resumo por Cliente**
  - **Resumo por Produto**

---

## 🚀 Como Executar
1. Clone o repositório:
   ```bash
   git clone  https://github.com/Matheus780-dev/python-automacoes-semana3.git
   cd automacao-python-semana3
   ```

2. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```

3. Execute o script desejado:
   ```bash
   python projeto_excel_simples/segunda.py
   ```

---

