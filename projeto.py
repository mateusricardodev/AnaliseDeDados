import pandas as pd

# Lê a planilha (header na linha 3, índice 0)
tabela = pd.read_excel('planilha-dados.xlsx', header=3)


coluna_regiao = tabela.iloc[0:63, 0] # Pega da linha 3 até a ultima região
#vi em um video do youtube https://youtu.be/IT7zPluDADk?si=g9VvFKdhZBUapXjw

coluna_regiao = coluna_regiao[coluna_regiao != 'Região de Saúde (CIR)']

regioes_separadas = coluna_regiao.str.split(' ', n=1, expand=True)  # expand = True, controla o formato do resultado do split.
regioes_separadas.columns = ['codigo', 'nome']

##print(regioes_separadas)
##regioes_separadas.to_excel('regioes_separadas.xlsx', index=False)

#coluna_idade_1 = tabela.iloc [1, 1:14] # menos de 1 ano à 12 anos
#print(coluna_idade_1.sum())
#coluna_idade_2 = tabela.iloc [1, 14:19] # de 13 à 17 anos
#print(coluna_idade_2.sum())
#coluna_idade_3 = tabela.iloc [1, 19:32] # 18 à 30 anos
#print(coluna_idade_3.sum())
#coluna_idade_4 = tabela.iloc [1, 32:52] # de 31 à 50 anos
#print(coluna_idade_4.sum())
#coluna_idade_5 = tabela.iloc [1, 52: 66] # de 51 à 64 anos
#print(coluna_idade_5.sum())
#coluna_idade_6 = tabela.iloc [1, 66: 82] # 65+ 
#print(coluna_idade_6.sum())
lista = []
for i in range (1, 63):
    nome = regioes_separadas.loc[i, 'nome'] 
    codigo = regioes_separadas.loc[i, 'codigo']
    criancas = tabela.iloc[i, 1:14].sum()
    #print(f'total de criancas na regiao {nome} = {criancas}')
    adolescentes = tabela.iloc[i, 14:19].sum()
    #print(f'total de adolescentes na regiao {nome} = {adolescentes}')
    adultos1 = tabela.iloc[i, 19:32].sum()
    #print(f'total de adultos até 30 anos na regiao {nome} = {adultos1}')
    adultos2 = tabela.iloc[i, 32:52].sum()
    #print(f'total de adultos de 31 à 50 anos na regiao {nome} = {adultos2}')
    adultos3 = tabela.iloc[i, 52:66].sum()
    #print(f'total de adultos - 51 à 64 anos na regiao {nome} = {adultos3}')
    idosos = tabela.iloc[i, 66:82].sum()
    #print(f'total de idosos na regiao {nome} = {idosos}')
    lista.append([codigo, nome, criancas, adolescentes, adultos1, adultos2, adultos3, idosos])
    #print(lista)
    #print("\n\n\n")

colunas = ['Código', 'Região de Saúde', 'Crianças (0–12)', 'Adolescentes (13–17)',
           'Adultos (18–30)', 'Adultos (31–50)', 'Adultos (51–64)', 'Idosos (65+)']
    
df_resultado = pd.DataFrame(lista, columns=colunas)
 
# ── Totais do estado inteiro (soma de todas as regiões) ──────────
total_criancas     = int(df_resultado['Crianças (0–12)'].sum())
total_adolescentes = int(df_resultado['Adolescentes (13–17)'].sum())
total_adultos1     = int(df_resultado['Adultos (18–30)'].sum())
total_adultos2     = int(df_resultado['Adultos (31–50)'].sum())
total_adultos3     = int(df_resultado['Adultos (51–64)'].sum())
total_idosos       = int(df_resultado['Idosos (65+)'].sum())

# ── Ranking: top 5 mais e menos populosas ────────────────────────
# Cria coluna com total geral de cada região (soma de todas as faixas)
df_resultado['Total'] = (
    df_resultado['Crianças (0–12)'] +
    df_resultado['Adolescentes (13–17)'] +
    df_resultado['Adultos (18–30)'] +
    df_resultado['Adultos (31–50)'] +
    df_resultado['Adultos (51–64)'] +
    df_resultado['Idosos (65+)']
)

cinco_maiores      = df_resultado.nlargest(5, 'Total')[['Região de Saúde', 'Total']] 
cinco_menores     = df_resultado.nsmallest(5, 'Total')[['Região de Saúde', 'Total']] 
df_resultado['Indice'] = (df_resultado['Idosos (65+)'] / df_resultado['Crianças (0–12)'] * 100).round(1)
total_populacao    = int(df_resultado['Total'].sum())
regiao_mais_pop    = df_resultado.loc[df_resultado['Total'].idxmax(), 'Região de Saúde']
regiao_menos_pop   = df_resultado.loc[df_resultado['Total'].idxmin(), 'Região de Saúde']

totais_faixas = {
    'Crianças (0–12)':      total_criancas,
    'Adolescentes (13–17)': total_adolescentes,
    'Adultos (18–30)':      total_adultos1,
    'Adultos (31–50)':      total_adultos2,
    'Adultos (51–64)':      total_adultos3,
    'Idosos (65+)':         total_idosos,
}
faixa_predominante = max(totais_faixas, key=totais_faixas.get)
total_estado_geral = total_criancas + total_adolescentes + total_adultos1 + total_adultos2 + total_adultos3 + total_idosos
pct_faixa_pred     = round(totais_faixas[faixa_predominante] / total_estado_geral * 100, 1)
num_regioes_alerta = int((df_resultado['Indice'] >= 100).sum())  # regiões com índice de envelhecimento em alerta


def linhas_ranking(df_rank, destaque_cor):
    html_rank = ""
    for pos, (_, row) in enumerate(df_rank.iterrows(), start=1):
        html_rank += f"""
            <tr class="rank-linha">
                <td class="rank-pos" style="color:{destaque_cor}">{pos}º</td>
                <td class="rank-nome">{row['Região de Saúde']}</td>
                <td class="rank-total">{formatar(row['Total'])}</td>
            </tr>"""
    return html_rank


def formatar(n):
    return f"{int(n):,}".replace(",", ".")

linhas_rank_mais  = linhas_ranking(cinco_maiores,  '#1a3557')
linhas_rank_menos = linhas_ranking(cinco_menores, '#8e5bb5')

linhas_html = ""
for _, row in df_resultado.iterrows():
    linhas_html += f"""
        <tr class="linha-dados"
            data-nome="{row['Região de Saúde']}"
            data-criancas="{int(row['Crianças (0–12)'])}"
            data-adolescentes="{int(row['Adolescentes (13–17)'])}"
            data-adultos1="{int(row['Adultos (18–30)'])}"
            data-adultos2="{int(row['Adultos (31–50)'])}"
            data-adultos3="{int(row['Adultos (51–64)'])}"
            data-idosos="{int(row['Idosos (65+)'])}"
            data-indice="{row['Indice']}">
            <td class="col-codigo">{row['Código']}</td>
            <td class="col-nome">
                <span class="nome-regiao">{row['Região de Saúde']}</span>
                <span class="icone-seta">&#9658;</span>
            </td>
            <td>{formatar(row['Crianças (0–12)'])}</td>
            <td>{formatar(row['Adolescentes (13–17)'])}</td>
            <td>{formatar(row['Adultos (18–30)'])}</td>
            <td>{formatar(row['Adultos (31–50)'])}</td>
            <td>{formatar(row['Adultos (51–64)'])}</td>
            <td>{formatar(row['Idosos (65+)'])}</td>
            <td><span class="badge-indice {'alerta' if row['Indice'] >= 100 else 'normal'}">{row['Indice']}</span></td>
        </tr>
        <tr class="linha-grafico">
            <td colspan="9">
                <div class="grafico-container">
                    <canvas class="grafico-canvas"></canvas>
                </div>
            </td>
        </tr>"""
 

html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Análise da População Residente — SP</title>
  <link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;600&family=Source+Sans+3:wght@300;400;500;600&display=swap" rel="stylesheet">
  <script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
  <style>
    *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
    :root {{
      --azul-escuro: #1a3557;
      --azul-medio: #2a5298;
      --azul-claro: #e8f0fa;
      --dourado: #b8922a;
      --cinza-texto: #3a3a3a;
      --cinza-suave: #f5f5f3;
      --borda: #d0d8e4;
      --fonte-titulo: 'Playfair Display', serif;
      --fonte-corpo: 'Source Sans 3', sans-serif;
    }}
    body {{ background: #fff; font-family: var(--fonte-corpo); color: var(--cinza-texto); }}
 
    .barra-topo {{ background: var(--azul-escuro); height: 6px; }}
    .barra-gov {{ background: var(--azul-medio); padding: 8px 60px; display: flex; align-items: center; gap: 12px; }}
    .barra-gov span {{ font-size: 12px; font-weight: 500; color: rgba(255,255,255,0.85); letter-spacing: 0.04em; text-transform: uppercase; }}
    .barra-gov .sep {{ width: 1px; height: 14px; background: rgba(255,255,255,0.3); }}
    .barra-gov .linkedin-link {{
        margin-left: auto;
        display: flex; align-items: center; gap: 6px;
        color: rgba(255,255,255,0.75);
        text-decoration: none;
        font-size: 11px; font-weight: 500; letter-spacing: 0.04em;
        transition: color 0.15s;
    }}
    .barra-gov .linkedin-link:hover {{ color: #fff; }}
    .barra-gov .linkedin-link svg {{ width: 15px; height: 15px; fill: currentColor; }}
 
    .header {{ padding: 56px 60px 48px; border-bottom: 1px solid var(--borda); position: relative; }}
    .header::before {{ content: ''; position: absolute; left: 0; top: 0; bottom: 0; width: 5px; background: linear-gradient(to bottom, var(--azul-medio), var(--azul-escuro)); }}
    .header-meta {{ display: flex; align-items: center; gap: 10px; margin-bottom: 20px; }}
    .tag {{ font-size: 11px; font-weight: 600; letter-spacing: 0.08em; text-transform: uppercase; color: var(--azul-medio); background: var(--azul-claro); padding: 4px 10px; border-radius: 3px; }}
    .tag-ano {{ font-size: 11px; font-weight: 500; color: #888; letter-spacing: 0.05em; }}
    .header h1 {{ font-family: var(--fonte-titulo); font-size: 32px; font-weight: 600; color: var(--azul-escuro); line-height: 1.3; max-width: 780px; margin-bottom: 8px; }}
    .header h1 em {{ font-style: italic; color: var(--dourado); }}
    .header-subtitulo {{ font-size: 15px; font-weight: 300; color: #666; margin-bottom: 32px; max-width: 680px; line-height: 1.5; }}
    .divider-ornamental {{ display: flex; align-items: center; gap: 12px; margin-bottom: 28px; }}
    .divider-ornamental .linha {{ height: 1px; background: var(--borda); flex: 1; max-width: 120px; }}
    .divider-ornamental .losango {{ width: 7px; height: 7px; background: var(--dourado); transform: rotate(45deg); }}
    .contexto-grid {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 1px; background: var(--borda); border: 1px solid var(--borda); border-radius: 6px; overflow: hidden; max-width: 860px; }}
    .contexto-item {{ background: #fff; padding: 18px 22px; }}
    .contexto-item .rotulo {{ font-size: 10px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.1em; color: #999; margin-bottom: 5px; }}
    .contexto-item .valor {{ font-size: 14px; font-weight: 500; color: var(--azul-escuro); line-height: 1.4; }}
 
    .descricao-section {{ padding: 44px 60px 40px; border-bottom: 1px solid var(--borda); }}
    .secao-titulo {{ font-size: 10px; font-weight: 700; letter-spacing: 0.14em; text-transform: uppercase; color: var(--azul-medio); margin-bottom: 14px; }}
    .descricao-texto {{ font-size: 15px; font-weight: 400; line-height: 1.8; color: #444; max-width: 740px; }}
    .destaque-inline {{ font-weight: 600; color: var(--azul-escuro); }}

    /* ── Resumo executivo ── */
    .resumo-section {{ padding: 44px 60px; border-bottom: 1px solid var(--borda); background: var(--cinza-suave); }}
    .resumo-section .secao-titulo {{ margin-bottom: 20px; }}
    .resumo-grid {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; }}
    .resumo-card {{
        background: #fff;
        border: 1px solid var(--borda);
        border-radius: 8px;
        padding: 20px 22px 18px;
        position: relative;
        overflow: hidden;
    }}
    .resumo-card::before {{
        content: '';
        position: absolute;
        top: 0; left: 0; right: 0;
        height: 3px;
    }}
    .resumo-card.c-azul::before   {{ background: var(--azul-medio); }}
    .resumo-card.c-verde::before  {{ background: #1a8a4a; }}
    .resumo-card.c-dourado::before {{ background: var(--dourado); }}
    .resumo-card.c-alerta::before {{ background: #c0392b; }}
    .resumo-rotulo {{
        font-size: 10px; font-weight: 700; letter-spacing: 0.1em;
        text-transform: uppercase; color: #999; margin-bottom: 8px;
    }}
    .resumo-valor {{
        font-size: 26px; font-weight: 600; line-height: 1.1;
        color: var(--azul-escuro); margin-bottom: 4px;
    }}
    .resumo-card.c-alerta .resumo-valor {{ color: #c0392b; }}
    .resumo-detalhe {{
        font-size: 12px; color: #888; line-height: 1.4; margin-top: 6px;
    }}
 
    /* ── Gráfico geral do estado ── */
    .grafico-estado-section {{ padding: 44px 60px; border-bottom: 1px solid var(--borda); }}
    .grafico-estado-section .secao-titulo {{ margin-bottom: 6px; }}
    .grafico-estado-descricao {{ font-size: 13px; color: #888; margin-bottom: 24px; }}
    .grafico-estado-wrapper {{
        background: var(--cinza-suave);
        border: 1px solid var(--borda);
        border-radius: 8px;
        padding: 28px 32px;
        height: 320px;
        position: relative;
    }}

    /* ── Ranking ── */
    .ranking-section {{ padding: 44px 60px; border-bottom: 1px solid var(--borda); }}
    .ranking-section .secao-titulo {{ margin-bottom: 6px; }}
    .ranking-descricao {{ font-size: 13px; color: #888; margin-bottom: 28px; }}
    .ranking-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 24px; max-width: 900px; }}
    .ranking-card {{ border: 1px solid var(--borda); border-radius: 8px; overflow: hidden; }}
    .ranking-card-header {{
        padding: 14px 20px;
        font-size: 11px; font-weight: 700; letter-spacing: 0.1em; text-transform: uppercase;
        display: flex; align-items: center; gap: 8px;
    }}
    .ranking-card-header.mais  {{ background: #eaf1fb; color: #1a3557; border-bottom: 1px solid var(--borda); }}
    .ranking-card-header.menos {{ background: #f3eefb; color: #5a3580; border-bottom: 1px solid var(--borda); }}
    .ranking-card-header .bolinha {{
        width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0;
    }}
    .mais  .bolinha {{ background: #1a3557; }}
    .menos .bolinha {{ background: #8e5bb5; }}
    table.rank-table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
    tr.rank-linha {{ border-bottom: 1px solid var(--borda); }}
    tr.rank-linha:last-child {{ border-bottom: none; }}
    tr.rank-linha:hover {{ background: var(--cinza-suave); }}
    td.rank-pos  {{ padding: 11px 14px; font-weight: 700; font-size: 13px; width: 36px; }}
    td.rank-nome {{ padding: 11px 8px; color: var(--cinza-texto); }}
    td.rank-total {{ padding: 11px 14px; text-align: right; font-family: monospace; font-size: 12px; color: #666; white-space: nowrap; }}

    /* ── Tabela ── */
    .tabela-section {{ padding: 44px 60px 60px; }}
    .tabela-section .secao-titulo {{ margin-bottom: 20px; }}
    .tabela-wrapper {{ overflow-x: auto; border: 1px solid var(--borda); border-radius: 8px; }}
    table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
    thead tr {{ background: var(--azul-escuro); }}
    thead th {{ padding: 13px 14px; text-align: right; font-weight: 600; font-size: 11px; letter-spacing: 0.05em; text-transform: uppercase; color: rgba(255,255,255,0.85); white-space: nowrap; }}
    thead th.col-codigo, thead th.col-nome {{ text-align: left; }}
 
    /* ── Linhas de dados (clicáveis) ── */
    tr.linha-dados {{ border-bottom: 1px solid var(--borda); cursor: pointer; transition: background 0.15s; }}
    tr.linha-dados:hover {{ background: var(--azul-claro); }}
    tr.linha-dados.aberta {{ background: var(--azul-claro); border-bottom: none; }}
 
    td {{ padding: 11px 14px; color: var(--cinza-texto); text-align: right; }}
    td.col-codigo {{ font-family: monospace; font-size: 12px; color: #888; text-align: left; }}
    td.col-nome {{ text-align: left; font-weight: 500; color: var(--azul-escuro); min-width: 200px; }}
 
    /* ── Seta indicadora ── */
    .nome-regiao {{ margin-right: 8px; }}
    .icone-seta {{ font-size: 9px; color: var(--azul-medio); display: inline-block; transition: transform 0.25s; opacity: 0.6; }}
    tr.linha-dados.aberta .icone-seta {{ transform: rotate(90deg); }}
 
    /* ── Linha do gráfico (oculta por padrão) ── */
    tr.linha-grafico {{ display: none; background: #f0f5fc; border-bottom: 2px solid var(--azul-claro); }}
    tr.linha-grafico td {{ padding: 0; }}
    .grafico-container {{
        padding: 24px 40px 28px;
        height: 260px;
        overflow: hidden;
    }}
    tr.linha-grafico.visivel {{ display: table-row; }}
 
    /* ── Índice de envelhecimento ── */
    .badge-indice {{
        display: inline-block;
        padding: 3px 9px;
        border-radius: 12px;
        font-size: 12px;
        font-weight: 600;
        font-family: monospace;
        white-space: nowrap;
    }}
    .badge-indice.normal {{ background: #e6f4ec; color: #1a6b3a; }}
    .badge-indice.alerta {{ background: #fdecea; color: #a0281e; }}

    /* ── Busca e ordenação ── */
    .tabela-controles {{
        display: flex; align-items: center; gap: 12px;
        margin-bottom: 16px; flex-wrap: wrap;
    }}
    .campo-busca {{
        display: flex; align-items: center; gap: 8px;
        background: var(--cinza-suave); border: 1px solid var(--borda);
        border-radius: 5px; padding: 8px 14px; flex: 1; max-width: 320px;
        transition: border-color 0.15s;
    }}
    .campo-busca:focus-within {{ border-color: var(--azul-medio); }}
    .campo-busca svg {{ width: 14px; height: 14px; color: #aaa; flex-shrink: 0; }}
    .campo-busca input {{
        border: none; background: transparent; outline: none;
        font-family: var(--fonte-corpo); font-size: 13px; color: var(--cinza-texto);
        width: 100%;
    }}
    .campo-busca input::placeholder {{ color: #bbb; }}
    .sem-resultado {{
        text-align: center; padding: 32px; font-size: 13px;
        color: #aaa; display: none;
    }}
    thead th.ordenavel {{ cursor: pointer; user-select: none; }}
    thead th.ordenavel:hover {{ background: #243f6a; }}
    thead th .seta-ordem {{
        display: inline-block; margin-left: 5px; opacity: 0.4; font-size: 10px;
    }}
    thead th.asc  .seta-ordem {{ opacity: 1; content: '▲'; }}
    thead th.desc .seta-ordem {{ opacity: 1; }}

    /* ── Botão exportar ── */
    .btn-exportar {{
        display: inline-flex; align-items: center; gap: 8px;
        background: var(--azul-escuro); color: #fff;
        border: none; border-radius: 5px; cursor: pointer;
        font-family: var(--fonte-corpo); font-size: 12px; font-weight: 600;
        letter-spacing: 0.05em; text-transform: uppercase;
        padding: 9px 18px; margin-bottom: 16px;
        transition: background 0.15s;
    }}
    .btn-exportar:hover {{ background: var(--azul-medio); }}
    .btn-exportar svg {{ width: 14px; height: 14px; }}
 
    .rodape-tabela {{ margin-top: 12px; font-size: 11px; color: #aaa; letter-spacing: 0.03em; }}
    .legenda-indice {{ margin-top: 10px; font-size: 12px; color: #888; display: flex; align-items: center; gap: 6px; flex-wrap: wrap; }}
    .rodape-header {{ padding: 20px 60px; background: var(--cinza-suave); border-top: 1px solid var(--borda); display: flex; justify-content: space-between; align-items: center; }}
    .rodape-header span {{ font-size: 11px; color: #aaa; letter-spacing: 0.04em; }}
  </style>
</head>
<body>
 
  <div class="barra-topo"></div>
  <div class="barra-gov">
    <span>Mateus Ricardo dos Santos</span>
    <a class="linkedin-link" href="https://www.linkedin.com/in/mateus-ricardo" target="_blank">
      <svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
        <path d="M20.447 20.452h-3.554v-5.569c0-1.328-.027-3.037-1.852-3.037-1.853 0-2.136 1.445-2.136 2.939v5.667H9.351V9h3.414v1.561h.046c.477-.9 1.637-1.85 3.37-1.85 3.601 0 4.267 2.37 4.267 5.455v6.286zM5.337 7.433a2.062 2.062 0 0 1-2.063-2.065 2.064 2.064 0 1 1 2.063 2.065zm1.782 13.019H3.555V9h3.564v11.452zM22.225 0H1.771C.792 0 0 .774 0 1.729v20.542C0 23.227.792 24 1.771 24h20.451C23.2 24 24 23.227 24 22.271V1.729C24 .774 23.2 0 22.222 0h.003z"/>
      </svg>
      LinkedIn
    </a>
  </div>
 
  <header class="header">
    <div class="header-meta">
      <span class="tag">Análise Demográfica</span>
      <span class="tag-ano">Período: 2025</span>
    </div>
    <h1>Análise da População Residente por<br>
    <em>Idade Simples</em> segundo Região de Saúde (CIR)</h1>
    <p class="header-subtitulo">Estado de São Paulo — Estudo de Estimativas Populacionais por Município e Idade</p>
    <div class="divider-ornamental">
      <div class="linha"></div><div class="losango"></div><div class="linha"></div>
    </div>
    <div class="contexto-grid">
      <div class="contexto-item"><div class="rotulo">Unidade Federativa</div><div class="valor">São Paulo — SP</div></div>
      <div class="contexto-item"><div class="rotulo">Fonte dos Dados</div><div class="valor">DATASUS / IBGE</div></div>
      <div class="contexto-item"><div class="rotulo">Ano de Referência</div><div class="valor">2025</div></div>
    </div>
  </header>
 
  <section class="resumo-section">
    <div class="secao-titulo">Resumo executivo</div>
    <div class="resumo-grid">

      <div class="resumo-card c-azul">
        <div class="resumo-rotulo">População total do estado</div>
        <div class="resumo-valor">{formatar(total_populacao)}</div>
        <div class="resumo-detalhe">Soma de todas as 62 regiões de saúde · SP · 2025</div>
      </div>

      <div class="resumo-card c-verde">
        <div class="resumo-rotulo">Região mais populosa</div>
        <div class="resumo-valor" style="font-size:17px; padding-top:4px">{regiao_mais_pop}</div>
        <div class="resumo-detalhe">{formatar(int(df_resultado['Total'].max()))} habitantes</div>
      </div>

      <div class="resumo-card c-dourado">
        <div class="resumo-rotulo">Faixa etária predominante</div>
        <div class="resumo-valor" style="font-size:17px; padding-top:4px">{faixa_predominante}</div>
        <div class="resumo-detalhe">{pct_faixa_pred}% da população total do estado</div>
      </div>

      <div class="resumo-card c-alerta">
        <div class="resumo-rotulo">Regiões em alerta de envelhecimento</div>
        <div class="resumo-valor">{num_regioes_alerta} <span style="font-size:14px; font-weight:400; color:#888">de 62</span></div>
        <div class="resumo-detalhe">Regiões com índice de envelhecimento ≥ 100</div>
      </div>

    </div>
  </section>

  <section class="descricao-section">
    <div class="secao-titulo">Sobre este relatório</div>
    <p class="descricao-texto">
      Este relatório apresenta a distribuição da <span class="destaque-inline">população residente</span>
      no estado de São Paulo, organizada por <span class="destaque-inline">idade simples</span> (ano a ano)
      e segmentada segundo as <span class="destaque-inline">Regiões de Saúde (CIR)</span> —
      Comissões Intergestores Regionais definidas pelo Ministério da Saúde.
      Os dados são provenientes do estudo de estimativas populacionais elaborado em conjunto pelo
      DATASUS e pelo IBGE, cobrindo o período de 2025.
      Clique em qualquer linha da tabela para visualizar o gráfico de distribuição por faixa etária.
    </p>
  </section>
 
  <section class="grafico-estado-section">
    <div class="secao-titulo">Visão geral do estado de São Paulo</div>
    <p class="grafico-estado-descricao">Distribuição total da população por faixa etária — soma de todas as 62 regiões de saúde</p>
    <div class="grafico-estado-wrapper">
      <canvas id="grafico-estado"></canvas>
    </div>
  </section>

  <section class="ranking-section">
    <div class="secao-titulo">Ranking de regiões por população total</div>
    <p class="ranking-descricao">Comparativo entre as 5 regiões mais populosas e as 5 menos populosas do estado</p>
    <div class="ranking-grid">

      <div class="ranking-card">
        <div class="ranking-card-header mais">
          <span class="bolinha"></span>
          5 regiões mais populosas
        </div>
        <table class="rank-table">
          <tbody>{linhas_rank_mais}</tbody>
        </table>
      </div>

      <div class="ranking-card">
        <div class="ranking-card-header menos">
          <span class="bolinha"></span>
          5 regiões menos populosas
        </div>
        <table class="rank-table">
          <tbody>{linhas_rank_menos}</tbody>
        </table>
      </div>

    </div>
  </section>

  <section class="tabela-section">
    <div class="secao-titulo">Distribuição por faixa etária e região — clique na linha para ver o gráfico</div>
    <div class="tabela-controles">
      <div class="campo-busca">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
          <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
        </svg>
        <input type="text" id="campo-busca" placeholder="Buscar região..." oninput="filtrarTabela(this.value)">
      </div>
      <button class="btn-exportar" onclick="exportarExcel()">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
          <polyline points="7 10 12 15 17 10"/>
          <line x1="12" y1="15" x2="12" y2="3"/>
        </svg>
        Baixar planilha Excel
      </button>
    </div>
    <div class="tabela-wrapper">
      <table id="tabela-principal">
        <thead>
          <tr>
            <th class="col-codigo">Código</th>
            <th class="col-nome ordenavel" onclick="ordenarTabela(1)">Região de Saúde <span class="seta-ordem">⇅</span></th>
            <th class="ordenavel" onclick="ordenarTabela(2)">Crianças (0–12) <span class="seta-ordem">⇅</span></th>
            <th class="ordenavel" onclick="ordenarTabela(3)">Adolescentes (13–17) <span class="seta-ordem">⇅</span></th>
            <th class="ordenavel" onclick="ordenarTabela(4)">Adultos (18–30) <span class="seta-ordem">⇅</span></th>
            <th class="ordenavel" onclick="ordenarTabela(5)">Adultos (31–50) <span class="seta-ordem">⇅</span></th>
            <th class="ordenavel" onclick="ordenarTabela(6)">Adultos (51–64) <span class="seta-ordem">⇅</span></th>
            <th class="ordenavel" onclick="ordenarTabela(7)">Idosos (65+) <span class="seta-ordem">⇅</span></th>
            <th class="ordenavel" onclick="ordenarTabela(8)">Índice env. <span class="seta-ordem">⇅</span></th>
          </tr>
        </thead>
        <tbody>
          {linhas_html}
        </tbody>
      </table>
      <p class="sem-resultado" id="sem-resultado">Nenhuma região encontrada para "<span id="termo-busca"></span>"</p>
    </div>
    <p class="rodape-tabela">Fonte: DATASUS — Estimativas Populacionais por Município e Idade · Valores absolutos</p>
    <p class="legenda-indice">
      <span class="badge-indice normal">ex: 85,0</span> Índice &lt; 100 — mais crianças que idosos &nbsp;·&nbsp;
      <span class="badge-indice alerta">ex: 120,0</span> Índice ≥ 100 — mais idosos que crianças (alerta de envelhecimento) &nbsp;·&nbsp;
      Fórmula: (Idosos ÷ Crianças) × 100
    </p>
  </section>
 
  <footer class="rodape-header">
    <span>Relatório gerado automaticamente a partir de dados públicos</span>
    <span>DATASUS · Estimativas Populacionais · SP · 2025</span>
  </footer>
 
  <script>
    // ── Busca em tempo real ──────────────────────────────────────
    function filtrarTabela(termo) {{
      const t = termo.toLowerCase().trim();
      let visiveis = 0;
      document.querySelectorAll('tr.linha-dados').forEach(function(tr) {{
        const nome = tr.dataset.nome.toLowerCase();
        const bate = nome.includes(t);
        tr.style.display = bate ? '' : 'none';
        // fecha gráfico se a linha for escondida
        if (!bate) {{
          const lg = tr.nextElementSibling;
          tr.classList.remove('aberta');
          lg.classList.remove('visivel');
          lg.style.display = 'none';
        }}
        tr.nextElementSibling.style.display = bate && tr.nextElementSibling.classList.contains('visivel') ? 'table-row' : (bate ? '' : 'none');
        if (bate) visiveis++;
      }});
      const semRes = document.getElementById('sem-resultado');
      document.getElementById('termo-busca').textContent = termo;
      semRes.style.display = visiveis === 0 && t !== '' ? 'block' : 'none';
    }}

    // ── Ordenação por coluna ─────────────────────────────────────
    let ultimaColuna = -1;
    let ordemAsc = true;

    function ordenarTabela(colIdx) {{
      const tbody = document.querySelector('#tabela-principal tbody');

      // Coleta pares: [linha-dados, linha-grafico]
      const pares = [];
      const rows = Array.from(tbody.querySelectorAll('tr.linha-dados'));
      rows.forEach(function(tr) {{
        pares.push([tr, tr.nextElementSibling]);
      }});

      // Determina direção
      if (ultimaColuna === colIdx) {{
        ordemAsc = !ordemAsc;
      }} else {{
        ordemAsc = true;
        ultimaColuna = colIdx;
      }}

      // Atualiza ícones nos cabeçalhos
      document.querySelectorAll('thead th.ordenavel').forEach(function(th) {{
        th.classList.remove('asc', 'desc');
        th.querySelector('.seta-ordem').textContent = '⇅';
      }});
      const thAtivo = document.querySelectorAll('thead th.ordenavel')[colIdx - 1];
      thAtivo.classList.add(ordemAsc ? 'asc' : 'desc');
      thAtivo.querySelector('.seta-ordem').textContent = ordemAsc ? '▲' : '▼';

      // Ordena
      pares.sort(function(a, b) {{
        let va, vb;
        if (colIdx === 1) {{
          va = a[0].dataset.nome;
          vb = b[0].dataset.nome;
          return ordemAsc ? va.localeCompare(vb, 'pt-BR') : vb.localeCompare(va, 'pt-BR');
        }}
        const mapa = {{2:'criancas', 3:'adolescentes', 4:'adultos1', 5:'adultos2', 6:'adultos3', 7:'idosos', 8:'indice'}};
        va = parseFloat(a[0].dataset[mapa[colIdx]]);
        vb = parseFloat(b[0].dataset[mapa[colIdx]]);
        return ordemAsc ? va - vb : vb - va;
      }});

      // Reinsere na ordem certa e fecha gráficos abertos
      if (graficoAtivo) {{ graficoAtivo.destroy(); graficoAtivo = null; }}
      pares.forEach(function(par) {{
        par[0].classList.remove('aberta');
        par[1].classList.remove('visivel');
        par[1].style.display = 'none';
        tbody.appendChild(par[0]);
        tbody.appendChild(par[1]);
      }});
    }}

    // ── Exportar para Excel ──────────────────────────────────────
    function exportarExcel() {{
      // Coleta os dados de cada linha (lê os data-* para ter números puros, sem formatação)
      const linhas = [];
      document.querySelectorAll('tr.linha-dados').forEach(function(tr) {{
        linhas.push({{
          'Código':               tr.querySelector('.col-codigo').textContent.trim(),
          'Região de Saúde':      tr.dataset.nome,
          'Crianças (0–12)':      parseInt(tr.dataset.criancas),
          'Adolescentes (13–17)': parseInt(tr.dataset.adolescentes),
          'Adultos (18–30)':      parseInt(tr.dataset.adultos1),
          'Adultos (31–50)':      parseInt(tr.dataset.adultos2),
          'Adultos (51–64)':      parseInt(tr.dataset.adultos3),
          'Idosos (65+)':         parseInt(tr.dataset.idosos),
        }});
      }});
 
      // Cria a planilha com o SheetJS
      const planilha  = XLSX.utils.json_to_sheet(linhas);
      const workbook  = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, planilha, 'Regiões de Saúde');
 
      // Dispara o download
      XLSX.writeFile(workbook, 'populacao_regioes_sp_2025.xlsx');
    }}
 
    // ── Gráfico geral do estado ──────────────────────────────────
    const totaisEstado = {{
      labels: [
        'Crianças (0–12)',
        'Adolescentes (13–17)',
        'Adultos (18–30)',
        'Adultos (31–50)',
        'Adultos (51–64)',
        'Idosos (65+)',
      ],
      valores: [
        {total_criancas},
        {total_adolescentes},
        {total_adultos1},
        {total_adultos2},
        {total_adultos3},
        {total_idosos},
      ],
      cores: [
        '#2a7fc9',
        '#3aaa7a',
        '#e8a020',
        '#d95f3b',
        '#8e5bb5',
        '#4a7c9e',
      ]
    }};

    new Chart(document.getElementById('grafico-estado'), {{
      type: 'bar',
      data: {{
        labels: totaisEstado.labels,
        datasets: [{{
          data: totaisEstado.valores,
          backgroundColor: totaisEstado.cores,
          borderRadius: 5,
          borderSkipped: false,
        }}]
      }},
      options: {{
        indexAxis: 'y',
        responsive: true,
        maintainAspectRatio: false,
        plugins: {{
          legend: {{ display: false }},
          tooltip: {{
            callbacks: {{
              label: function(ctx) {{
                return ' ' + ctx.raw.toLocaleString('pt-BR');
              }}
            }}
          }}
        }},
        scales: {{
          x: {{
            grid: {{ color: 'rgba(0,0,0,0.05)' }},
            ticks: {{
              font: {{ size: 11 }},
              callback: function(v) {{
                if (v >= 1000000) return (v / 1000000).toFixed(1) + 'M';
                if (v >= 1000) return (v / 1000).toFixed(0) + 'k';
                return v;
              }}
            }}
          }},
          y: {{
            grid: {{ display: false }},
            ticks: {{ font: {{ size: 13 }}, color: '#444' }}
          }}
        }}
      }}
    }});

    // Guarda referência ao gráfico aberto para destruir antes de criar outro
    let graficoAtivo = null;
 
    // Cores de cada faixa etária
    const cores = [
      '#2a7fc9',
      '#3aaa7a',
      '#e8a020',
      '#d95f3b',
      '#8e5bb5',
      '#4a7c9e',
    ];
 
    document.querySelectorAll('tr.linha-dados').forEach(function(tr) {{
      tr.addEventListener('click', function() {{
        const linhaGrafico = tr.nextElementSibling;
        const jaAberta = tr.classList.contains('aberta');
 
        // Fecha tudo que estiver aberto
        document.querySelectorAll('tr.linha-dados.aberta').forEach(function(outra) {{
          outra.classList.remove('aberta');
          outra.nextElementSibling.classList.remove('visivel');
          outra.nextElementSibling.style.display = 'none';
        }});
        if (graficoAtivo) {{
          graficoAtivo.destroy();
          graficoAtivo = null;
        }}
 
        // Se já estava aberta, só fecha (toggle)
        if (jaAberta) return;
 
        // Abre a linha do gráfico
        tr.classList.add('aberta');
        linhaGrafico.style.display = 'table-row';
        linhaGrafico.classList.add('visivel');
 
        // Lê os dados dos atributos data-* da linha clicada
        const valores = [
          parseInt(tr.dataset.criancas),
          parseInt(tr.dataset.adolescentes),
          parseInt(tr.dataset.adultos1),
          parseInt(tr.dataset.adultos2),
          parseInt(tr.dataset.adultos3),
          parseInt(tr.dataset.idosos),
        ];
 
        const labels = [
          'Crianças (0–12)',
          'Adolescentes (13–17)',
          'Adultos (18–30)',
          'Adultos (31–50)',
          'Adultos (51–64)',
          'Idosos (65+)',
        ];
 
        const canvas = linhaGrafico.querySelector('canvas');
        graficoAtivo = new Chart(canvas, {{
          type: 'bar',
          data: {{
            labels: labels,
            datasets: [{{
              data: valores,
              backgroundColor: cores,
              borderRadius: 4,
              borderSkipped: false,
            }}]
          }},
          options: {{
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            plugins: {{
              legend: {{ display: false }},
              tooltip: {{
                callbacks: {{
                  label: function(ctx) {{
                    return ' ' + ctx.raw.toLocaleString('pt-BR');
                  }}
                }}
              }}
            }},
            scales: {{
              x: {{
                grid: {{ color: 'rgba(0,0,0,0.05)' }},
                ticks: {{
                  font: {{ size: 11 }},
                  callback: function(v) {{
                    if (v >= 1000000) return (v / 1000000).toFixed(1) + 'M';
                    if (v >= 1000) return (v / 1000).toFixed(0) + 'k';
                    return v;
                  }}
                }}
              }},
              y: {{
                grid: {{ display: false }},
                ticks: {{ font: {{ size: 12 }} }}
              }}
            }}
          }}
        }});
      }});
    }});
  </script>
 
</body>
</html>"""

with open('relatorio.html', 'w', encoding='utf-8') as arquivo:
    arquivo.write(html)
 
print("Relatorio gerado: relatorio.html")