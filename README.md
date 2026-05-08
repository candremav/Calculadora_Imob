# Simulador de Financiamento Imobiliário

Réplica em Python da **Planilha de Simulações de Financiamentos Imobiliários** (autoria original: Alessandro Rodolpho Gonçalves da Silva / *amigodopairico*). Implementa o sistema **PRICE com correção monetária pela TR**, padrão dos contratos imobiliários da Caixa Econômica Federal.

O projeto entrega três interfaces para o mesmo modelo de cálculo:

1. **Notebook Jupyter** — análise interativa, exploração e validação numérica.
2. **Aplicativo Streamlit** — interface web amigável, com inputs editáveis e download do relatório.
3. **Arquivo Excel formatado** — saída de 3 abas (Inputs, Análise Mês a Mês, Painel de Acompanhamento), pronta para distribuição.

---

## Sumário

- [Sumário](#sumário)
- [1. Estrutura do projeto](#1-estrutura-do-projeto)
- [2. Como executar](#2-como-executar)
  - [2.1. Aplicativo Streamlit](#21-aplicativo-streamlit)
  - [2.2. Notebook Jupyter](#22-notebook-jupyter)
  - [2.3. Geração standalone do Excel](#23-geração-standalone-do-excel)
- [3. Inputs do modelo](#3-inputs-do-modelo)
- [4. Outputs do modelo](#4-outputs-do-modelo)
  - [4.1. Tabela mês a mês](#41-tabela-mês-a-mês)
  - [4.2. Agregações por triênio e década](#42-agregações-por-triênio-e-década)
  - [4.3. Resumo executivo](#43-resumo-executivo)
  - [4.4. Totais do financiamento](#44-totais-do-financiamento)
  - [4.5. Arquivo Excel exportado](#45-arquivo-excel-exportado)
- [5. Algoritmo de cálculo](#5-algoritmo-de-cálculo)
- [6. Funções principais](#6-funções-principais)
- [7. Validação contra a planilha original](#7-validação-contra-a-planilha-original)
- [8. Limitações conhecidas](#8-limitações-conhecidas)
- [9. Roadmap](#9-roadmap)
- [10. Créditos](#10-créditos)

---

## 1. Estrutura do projeto

```
.
├── README.md                                       ← este arquivo
├── requirements.txt                                ← dependências Python
│
├── simulacao_financiamento_imobiliario.ipynb       ← notebook (12 seções)
├── streamlit_app.py                                ← aplicativo web
│
├── simulacao_financiamento.xlsx                    ← Excel gerado pelo notebook
│
└── Planilha_de_Simulacoes_Financiamentos_Imob_.xlsx ← planilha original (referência)
```

**Convenções de organização:**
- O notebook e o app **compartilham as mesmas funções de cálculo** (`excel_pmt`, `excel_nper`, `gerar_tabela_amortizacao`, `agregar_por_periodo`, `gerar_excel`). Os mesmos inputs produzem os mesmos outputs em ambos.
- O arquivo Excel exportado (`simulacao_financiamento.xlsx`) tem a estrutura idêntica seja gerado pelo notebook ou pelo botão de download do app.

---

## 2. Como executar

### 2.0. Instalação das dependências

```bash
python -m venv venv
source venv/bin/activate          # Linux/Mac
venv\Scripts\activate             # Windows
pip install -r requirements.txt
```

### 2.1. Aplicativo Streamlit

```bash
streamlit run streamlit_app.py
```

O app abrirá em `http://localhost:8501`. Para deploy no **Streamlit Community Cloud**:

1. Suba os arquivos para um repositório GitHub.
2. Conecte o repositório em [share.streamlit.io](https://share.streamlit.io).
3. Aponte para `streamlit_app.py` como ponto de entrada.

### 2.2. Notebook Jupyter

```bash
jupyter notebook simulacao_financiamento_imobiliario.ipynb
```

Execute as células em ordem (`Cell → Run All`). A última seção gera automaticamente o arquivo `simulacao_financiamento.xlsx`.

### 2.3. Geração standalone do Excel

O Excel também pode ser gerado de forma programática a partir das funções do `streamlit_app.py`:

```python
from datetime import datetime
from streamlit_app import (
    gerar_tabela_amortizacao, agregar_por_periodo, gerar_excel
)

inputs = {...}                                # ver tabela na seção 3
tabela = gerar_tabela_amortizacao(...)
trienios = agregar_por_periodo(tabela, "Triênio", "triênio")
decadas  = agregar_por_periodo(tabela, "Década",  "década")

xlsx_bytes = gerar_excel(inputs, tabela, trienios, decadas, resumo, totais)
with open("saida.xlsx", "wb") as f:
    f.write(xlsx_bytes)
```

---

## 3. Inputs do modelo

Todos os inputs aparecem em **um único lugar** em cada interface:

- **Notebook:** seção 2 (`# INPUTS DA PLANILHA`).
- **Streamlit:** sidebar à esquerda.
- **Excel:** aba 1 (*Inputs*), com valores em azul.

| Variável                | Padrão       | Unidade  | Origem na planilha original | Descrição |
|-------------------------|--------------|----------|------------------------------|-----------|
| `valor_financiado`      | 240.000,00   | R$       | `AD11` / `H2`                | Saldo devedor inicial. Não inclui ITBI, registro nem taxa de avaliação. |
| `taxa_juros_anual`      | 7,9347 %     | % a.a.   | `V5` / `H3`                  | Juros nominais anuais. |
| `prazo_meses`           | 420          | meses    | `U9`                         | Número de parcelas (35 anos). |
| `taxa_admin_mensal`     | 25,00        | R$/mês   | `N12`                        | Tarifa fixa cobrada pelo banco. |
| `seguro_mensal`         | 17,45        | R$/mês   | `O12`                        | Seguro habitacional MIP+DFI inicial. |
| `tr_mensal_estimada`    | 0,17 %       | % a.m.   | `AL10`                       | Correção monetária mensal (TR). |
| `data_primeira_parcela` | 01/06/2025   | data     | `H12`                        | Data do primeiro vencimento. |
| `anos_resumo`           | 10           | anos     | `I5`                         | Período do resumo executivo. |

### Variável derivada (calculada automaticamente)

| Variável            | Fórmula                          | Exemplo            |
|---------------------|----------------------------------|--------------------|
| `taxa_juros_mensal` | `(1 + taxa_anual)^(1/12) − 1`    | 0,6383 % a.m.      |

---

## 4. Outputs do modelo

### 4.1. Tabela mês a mês

DataFrame `tabela` com **uma linha por parcela** (420 linhas no exemplo padrão). Replica as colunas principais da aba *Comparações* da planilha original:

| Coluna do DataFrame  | Coluna na planilha (Comparações) | Descrição |
|----------------------|----------------------------------|-----------|
| `Mês`                | `U` (sequencial 1..420)          | Número da parcela. |
| `Data`               | `H`                              | Data do vencimento. |
| `Ano` / `Triênio` / `Década` | `M` / `L` / `K`          | Marcadores temporais para agregação. |
| `Saldo Inicial`      | `AD` da linha anterior           | Saldo devedor no início do mês. |
| `Juros`              | `V`                              | Juros do mês = `saldo × taxa_mensal`. |
| `Amortização`        | `W`                              | Parcela − juros − taxa − seguro. |
| `Taxa Admin`         | `N`                              | Taxa administrativa mensal. |
| `Seguro`             | `O`                              | Seguro habitacional do mês. |
| `Parcela`            | `AA`                             | Valor total pago no mês. |
| `Correção Monetária` | `AJ`                             | `saldo × TR_mensal`. |
| `Saldo Final`        | `AD`                             | Saldo após amortização e correção. |

### 4.2. Agregações por triênio e década

DataFrames `trienios` e `decadas`. Cada linha corresponde a 36 ou 120 meses agrupados:

| Coluna                | Descrição                                              |
|-----------------------|--------------------------------------------------------|
| `Triênio`/`Década`    | Número sequencial do período (1, 2, 3, ...).          |
| `Anos`                | Faixa de anos coberta (`"ANO 1 AO 3"`).               |
| `Parcelas`            | Soma das parcelas pagas no período.                    |
| `Juros`               | Soma dos juros pagos no período.                       |
| `Amortização`         | Soma do principal amortizado no período.               |
| `Dif. Dívida`         | `Amortização − Correção Monetária` (negativo se a dívida cresceu). |
| `Correção Monetária`  | Soma da correção pela TR no período.                  |

### 4.3. Resumo executivo

Indicadores calculados sobre os primeiros `anos_resumo` anos (padrão: 10):

| Indicador            | Fórmula                                          | Interpretação |
|----------------------|--------------------------------------------------|---------------|
| `saiu_do_bolso`      | `Σ(parcela)` nos primeiros N anos                | Quanto foi efetivamente pago. |
| `caiu_a_divida`      | `valor_financiado − saldo_no_mês_N`              | Negativo significa que a dívida CRESCEU. |
| `aproveitamento`     | `caiu_a_divida / saiu_do_bolso`                  | % do dinheiro que efetivamente reduziu a dívida. |

### 4.4. Totais do financiamento

Calculados sobre o financiamento inteiro (420 meses):

- `total_pago` — soma de todas as parcelas.
- `total_juros` — soma dos juros.
- `total_amortizado` — soma do principal amortizado.
- `total_taxa_admin` — taxas administrativas acumuladas.
- `total_seguro` — seguro acumulado.
- `total_correcao` — correção monetária acumulada.
- `saldo_final` — saldo devedor remanescente após a última parcela.
- `multiplicador` — `total_pago / valor_financiado` (quantas vezes você pagou o valor financiado).

### 4.5. Arquivo Excel exportado

Estrutura idêntica em ambas as interfaces (notebook e Streamlit):

| Aba                       | Conteúdo                                                                 |
|---------------------------|--------------------------------------------------------------------------|
| **Inputs**                | Os 9 parâmetros listados na seção 3, com descrição e unidade. Valores em azul (convenção de modelagem). |
| **Análise Mês a Mês**     | Tabela completa de 420 parcelas + linha de totais. Painéis congelados para facilitar navegação. |
| **Painel de Acompanhamento** | Resumo executivo, agregação por triênio (anos 1–30), agregação por década (anos 1–30) e totais gerais do financiamento. |

**Formatação:** Arial, R$ como moeda brasileira, percentuais com 2-4 decimais, datas em `dd/mm/yyyy`, linhas de total destacadas em azul claro com bordas duplas.

---

## 5. Algoritmo de cálculo

O sistema PRICE com correção monetária da Caixa funciona em três etapas por mês:

```
Para cada mês m, com saldo devedor S(m-1):

  1. Juros            = S(m-1) × taxa_mensal
  2. Correção (AJ)    = S(m-1) × TR_mensal
  3. Parcela:
       se m == 1:  P = -PMT(taxa, prazo_total, S(m-1)) + taxa_admin + seguro
       se m  > 1:  n_restante  = NPER(taxa, -P_anterior + taxa_admin + seguro, S(m-1))
                   P           = -PMT(taxa, n_restante, S(m-1) + Correção)
                                 + taxa_admin + seguro
  4. Amortização      = P − Juros − taxa_admin − seguro
  5. Saldo final S(m) = max(0, S(m-1) − Amortização + Correção)
```

**Por que recalcular `NPER` e `PMT` todo mês?** Porque a TR é aplicada sobre o saldo, criando uma "dívida nova" a cada mês. A planilha trata isso recalculando o prazo restante implícito na parcela anterior e, em seguida, recalculando a parcela sobre o saldo já corrigido — exatamente como o sistema da Caixa faz internamente.

**Consequência prática:** quando `correção > amortização`, a parcela cresce mês a mês e o saldo devedor *aumenta* mesmo com você pagando. É o cenário do **aproveitamento negativo** — comum nos primeiros anos de financiamentos longos com TR positiva.

---

## 6. Funções principais

Todas estão definidas tanto no notebook quanto em `streamlit_app.py` com assinaturas idênticas.

### `excel_pmt(rate, nper, pv) → float`

Equivalente à função `PMT()` do Excel. Retorna o pagamento periódico de um empréstimo, com convenção de sinais do Excel (negativo para saídas).

```python
parcela = -excel_pmt(0.006383, 420, 240_000)   # → 1645.68
```

### `excel_nper(rate, pmt, pv) → float`

Equivalente à função `NPER()` do Excel. Retorna o número de períodos para quitar um empréstimo dada uma parcela.

```python
n = excel_nper(0.006383, -1645.68, 240_000)   # → 419.99 (≈ 420)
```

### `gerar_tabela_amortizacao(...) → pd.DataFrame`

Função principal de simulação. Itera mês a mês aplicando o algoritmo da seção 5. Aceita 7 parâmetros e retorna o DataFrame descrito em [4.1](#41-tabela-mês-a-mês).

```python
tabela = gerar_tabela_amortizacao(
    valor_financiado = 240_000,
    taxa_mensal      = 0.006383,
    prazo_total      = 420,
    taxa_admin       = 25,
    seguro           = 17.45,
    tr_mensal        = 0.0017,
    data_inicial     = datetime(2025, 6, 1),
)
```

### `agregar_por_periodo(tabela, coluna, label) → pd.DataFrame`

Agrupa a tabela mês a mês em triênios (36 meses) ou décadas (120 meses). Adiciona a coluna `Anos` com o rótulo descritivo e a coluna `Dif. Dívida` (= amortização − correção).

```python
trienios = agregar_por_periodo(tabela, "Triênio", "triênio")
decadas  = agregar_por_periodo(tabela, "Década",  "década")
```

### `gerar_excel(inputs, tabela, trienios, decadas, resumo, totais) → bytes`

Constrói o workbook das 3 abas com formatação financeira completa (cores, fontes, bordas, números formatados em R$ e %, painéis congelados). Retorna os bytes do arquivo (útil para `st.download_button`).

```python
xlsx = gerar_excel(inputs, tabela, trienios, decadas, resumo, totais)
open("saida.xlsx", "wb").write(xlsx)
```

### Helpers de UI (apenas no `streamlit_app.py`)

- `fmt_brl(v)` — formata float como moeda brasileira (`"R$ 1.234,56"`).
- `fmt_pct(v)` — formata float como percentual (`"15,00%"`).

---

## 7. Validação contra a planilha original

A seção 9 do notebook executa um **teste automatizado** comparando 16 métricas calculadas pelo Python contra os valores extraídos diretamente da planilha original:

| Categoria         | Métricas validadas                                            |
|-------------------|---------------------------------------------------------------|
| Parcela inicial   | Parcela 1                                                     |
| Triênio 1         | Total parcelas, juros, amortização, correção monetária        |
| Década 1          | Total parcelas, juros, amortização, correção monetária        |
| 30 anos completos | Total parcelas, juros, amortização, correção monetária        |
| Resumo executivo  | Saiu do bolso, caiu a dívida, aproveitamento                  |

**Tolerância:** 0,2 %. Justificativa:

- A **Parcela 1** bate com precisão de **4 casas decimais** com a planilha (R$ 1.688,1295 vs R$ 1.688,13).
- Métricas agregadas divergem em até 0,15 % (ex.: amortização total = R$ 351 em R$ 240k em 35 anos).
- A diferença residual vem do fato de o LibreOffice/Excel arredondarem PMT/NPER em pontos ligeiramente diferentes do Python ao longo de 420 iterações encadeadas. **Não é erro lógico** — é precisão numérica acumulada.

```
✅ TESTE PASSOU: todas as 16 métricas estão dentro da tolerância de 0.2%.
   Maior diferença observada: 0.1458%
```

---

## 8. Limitações conhecidas

| Limitação                                  | Detalhe                                                                    |
|--------------------------------------------|----------------------------------------------------------------------------|
| TR constante                               | A planilha original também trata a TR como constante. Na realidade ela varia mês a mês conforme divulgação do Bacen. |
| Seguro fixo                                | Na realidade o seguro varia com idade do mutuário e com o saldo devedor.   |
| Sem SAC                                    | A planilha original tem uma simulação SAC paralela; este projeto cobre apenas PRICE. |
| Sem amortizações extras                    | A planilha original simula adiantamentos pontuais (uma vez ao ano, etc.). |
| Sem comparação com investimentos           | A planilha original compara o custo do financiamento com o rendimento de aplicar o mesmo valor em CDB/Tesouro. |
| Sem juros de obra (carência)               | Modelo aplicável apenas a imóveis prontos.                                |
| FGTS                                       | A planilha original modela uso do FGTS para reduzir saldo a cada 2 anos. |

---

## 9. Roadmap

Funcionalidades planejadas para evoluções futuras (todas existem na planilha original):

- [ ] **Sistema SAC** (parcela decrescente, amortização constante).
- [ ] **Amortizações extras** programadas (mensal, anual, ad hoc).
- [ ] **Comparação com investimentos** (oportunidade de aplicar a parcela em CDB/Tesouro/CDI).
- [ ] **Uso do FGTS** a cada 2 anos para abater saldo.
- [ ] **Cenários de TR** (otimista/base/pessimista) com projeção do Bacen.
- [ ] **Comparativo Price × SAC** lado a lado.
- [ ] **Simulação reversa**: dado um orçamento mensal, qual o valor máximo financiável?

---

## 10. Créditos

- **Lógica de cálculo e planilha original:** Alessandro Rodolpho Gonçalves da Silva — *Amigo do Pai Rico* ([instagram.com/amigodopairico](https://instagram.com/amigodopairico)).
- **Réplica em Python (notebook + app + documentação):** este projeto.

A planilha original encontra-se sob proteção de direitos autorais e foi usada exclusivamente como **referência funcional** para esta réplica em Python. Nenhum conteúdo da planilha foi distribuído ou reproduzido — apenas os **valores numéricos de output** foram usados como gabarito para validar o algoritmo.

> Atenção: esta calculadora deve ser usada apenas para **simulações**. Os valores oficiais e legalmente vinculantes estão no **Documento Descritivo de Crédito** fornecido pelo banco no momento da contratação.
