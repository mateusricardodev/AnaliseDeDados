# 📊 Análise da População Residente por Região de Saúde (CIR) — SP

Relatório interativo gerado automaticamente com **Python e Pandas**, a partir de dados públicos do DATASUS/IBGE. O script lê uma planilha `.xlsx`, processa os dados populacionais por faixa etária e gera um arquivo `relatorio.html` completo e interativo.

---

## 🖥️ Demonstração

O relatório gerado inclui:

- **Resumo executivo** — cards com total da população, região mais populosa, faixa etária predominante e número de regiões em alerta de envelhecimento
- **Gráfico geral do estado** — distribuição da população por faixa etária, somando todas as 62 regiões
- **Ranking** — top 5 regiões mais populosas e 5 menos populosas
- **Tabela interativa** — com busca em tempo real, ordenação por coluna e gráfico de barras ao clicar em cada região
- **Índice de envelhecimento** — razão entre idosos e crianças por região, com alerta visual quando idosos superam crianças
- **Exportação para Excel** — download da tabela completa em `.xlsx` direto pelo navegador

---

## 🗂️ Estrutura do projeto

```
AnaliseDeDados/
├── projeto.py           # Script principal — processa os dados e gera o relatório
├── planilha-dados.xlsx  # Fonte de dados (DATASUS/IBGE) — não incluída no repositório
└── relatorio.html       # Relatório gerado (criado ao executar o script)
```

---

## ⚙️ Como executar

**1. Clone o repositório**
```bash
git clone https://github.com/mateusricardodev/AnaliseDeDados.git
cd AnaliseDeDados
```

**2. Instale a dependência**
```bash
pip install pandas openpyxl
```

**3. Adicione a planilha de dados**

Baixe o arquivo de estimativas populacionais no [DATASUS](https://datasus.saude.gov.br/) e salve como `planilha-dados.xlsx` na raiz do projeto.

**4. Execute o script**
```bash
python projeto.py
```

O arquivo `relatorio.html` será gerado na mesma pasta. Abra-o em qualquer navegador.

---

## 🔢 Faixas etárias analisadas

| Faixa | Intervalo |
|---|---|
| Crianças | 0 a 12 anos |
| Adolescentes | 13 a 17 anos |
| Adultos jovens | 18 a 30 anos |
| Adultos | 31 a 50 anos |
| Adultos maduros | 51 a 64 anos |
| Idosos | 65 anos ou mais |

---

## 📐 Índice de envelhecimento

Calculado pela fórmula:

```
Índice = (Idosos ÷ Crianças) × 100
```

- **Índice < 100** → mais crianças que idosos ✅
- **Índice ≥ 100** → mais idosos que crianças ⚠️ alerta para planejamento de saúde

---

## 🛠️ Tecnologias utilizadas

| Tecnologia | Uso |
|---|---|
| Python + Pandas | Leitura, tratamento e análise dos dados |
| Chart.js | Gráficos interativos no relatório HTML |
| SheetJS | Exportação para Excel pelo navegador |
| HTML + CSS + JavaScript | Interface do relatório |

---

## 📁 Fonte dos dados

**DATASUS / IBGE** — Estudo de Estimativas Populacionais por Município, Idade e Sexo · 2000–2025  
Acesso em: [datasus.saude.gov.br](https://datasus.saude.gov.br/)

---

## 👤 Autor

**Mateus Ricardo dos Santos**  
[![LinkedIn](https://img.shields.io/badge/LinkedIn-0077B5?style=flat&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/mateus-ricardo)
