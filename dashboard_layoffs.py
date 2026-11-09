# -*- coding: utf-8 -*-
"""
DASHBOARD DE LAYOFFS EM TECNOLOGIA (2022–2026) — versão em Python (Dash + Plotly)
APS 1 — Dashboards em Excel e Python · Henrique Fazzio Badin

Porta para Python o dashboard feito em Excel, seguindo o método das aulas
(Dash + Plotly, um único arquivo, pandas no lugar das tabelas dinâmicas).
As páginas Panorama, Setores e Empresas reproduzem as do Excel (mesmos gráficos,
títulos, cores e disposição). Diferenças intencionais: filtros em listas suspensas,
mapa de bolhas, marcos da IA no gráfico mensal e a página extra "Análises".

COMO RODAR (detalhes no README.md)
1. Instale as bibliotecas uma vez:   pip install -r requirements.txt
2. Execute:                          python dashboard_layoffs.py
3. Abra no navegador:                http://127.0.0.1:8050
   Para encerrar, volte ao terminal e aperte Ctrl+C.

FUNCIONA OFFLINE
- Os dados ficam em dados/layoffs.csv (mesmo CSV usado no Excel).
- O desenho do mapa-múndi fica em assets/topojson/world_110m.json, então o mapa
  não precisa baixar nada da internet.

ROTEIRO (os marcadores # %% criam células no VS Code/Spyder)
1. Bibliotecas e configuração.
2. Carga e tratamento dos dados: o mesmo passo a passo do Power Query do Excel.
3. Funções de filtro, formatação e gráficos (os groupby substituem as tabelas dinâmicas).
4. Layout Dash: cabeçalho com navegação, linha do tempo, filtros, indicadores e as páginas.
5. Callbacks: controles (linha do tempo, cliques, limpar) -> filtros -> indicadores e gráficos.
6. Iniciar o servidor local.

DEFINIÇÕES (iguais às do Excel)
- Demitidos = soma de total_laid_off (só eventos que informaram o número).
- Eventos = linhas da base (cada linha é um anúncio de corte de uma empresa numa data).
- % médio do quadro cortado = média simples de percentage_laid_off entre os eventos que informaram.
- Fechamento = evento em que a empresa demitiu 100% do quadro.
- Capital aberto = estágio Post-IPO; demais estágios = Privada.
"""

# %% 1. Bibliotecas e configuração
import warnings
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from dash import Dash, Input, Output, State, ctx, dash_table, dcc, html, no_update

warnings.filterwarnings("ignore", category=DeprecationWarning)  # avisos internos de versões antigas do pandas

PASTA = Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()
ARQUIVO_CSV = PASTA / "dados" / "layoffs.csv"
PORTA = 8050

# Paleta: a mesma registrada na aba "tabelas" do Excel
COR = {
    "fundo": "#EFEDEA", "cartao": "#FFFFFF", "texto": "#24272B", "destaque": "#CC6A2A",
    "secundaria": "#4A5058", "neutro": "#D9D5CF", "apoio": "#8C9198", "destaque_claro": "#F7EEE7",
}

# Marcos da IA generativa (análise proposta na APS: picos de demissões x lançamentos)
# empresa: "OpenAI", "Anthropic" ou "Outra" (define a cor da linha no gráfico mensal)
MARCOS_IA = [
    ("2022-11-30", "ChatGPT", "OpenAI"),
    ("2023-03-14", "GPT-4 · Claude", "Ambas"),          # GPT-4 e o primeiro Claude saíram no mesmo dia
    ("2024-03-04", "Claude 3", "Anthropic"),
    ("2024-05-13", "GPT-4o", "OpenAI"),
    ("2024-06-20", "Claude 3.5 Sonnet", "Anthropic"),
    ("2025-01-20", "DeepSeek-R1", "Outra"),
    ("2025-05-22", "Claude 4", "Anthropic"),
    ("2025-08-07", "GPT-5", "OpenAI"),
    ("2025-11-24", "Claude Opus 4.5", "Anthropic"),
    ("2026-06-09", "Claude Fable 5", "Anthropic"),
]

# Coordenadas aproximadas (latitude, longitude) para as bolhas do mapa.
# Usamos coordenadas em vez do nome do país para que países pequenos
# (Singapura, Hong Kong, Malta...) também apareçam no mapa.
COORDENADAS = {
    "Argentina": (-34.0, -64.0), "Australia": (-25.3, 133.8), "Austria": (47.5, 14.6), "Bahrain": (26.0, 50.6),
    "Belgium": (50.5, 4.5), "Brazil": (-10.8, -52.9), "Canada": (56.1, -106.3), "Cayman Islands": (19.3, -81.3),
    "Chile": (-35.7, -71.5), "China": (35.9, 104.2), "Colombia": (4.6, -74.3), "Cyprus": (35.1, 33.4),
    "Czech Republic": (49.8, 15.5), "Denmark": (56.3, 9.5), "Egypt": (26.8, 30.8), "Estonia": (58.6, 25.0),
    "Finland": (61.9, 25.7), "France": (46.6, 2.2), "Germany": (51.2, 10.5), "Ghana": (7.9, -1.0),
    "Greece": (39.1, 21.8), "Hong Kong": (22.3, 114.2), "Hungary": (47.2, 19.5), "India": (21.0, 78.0),
    "Indonesia": (-2.5, 118.0), "Ireland": (53.4, -8.2), "Israel": (31.0, 34.9), "Italy": (41.9, 12.6),
    "Japan": (36.2, 138.3), "Kenya": (0.0, 37.9), "Lithuania": (55.2, 23.9), "Luxembourg": (49.8, 6.1),
    "Malaysia": (4.2, 102.0), "Malta": (35.9, 14.4), "Mexico": (23.6, -102.6), "Morocco": (31.8, -7.1),
    "Netherlands": (52.1, 5.3), "New Zealand": (-40.9, 174.9), "Nigeria": (9.1, 8.7), "Norway": (60.5, 8.5),
    "Pakistan": (30.4, 69.3), "Peru": (-9.2, -75.0), "Philippines": (12.9, 121.8), "Poland": (51.9, 19.1),
    "Portugal": (39.4, -8.2), "Romania": (45.9, 25.0), "Russia": (61.5, 105.3), "Saudi Arabia": (23.9, 45.1),
    "Senegal": (14.5, -14.5), "Seychelles": (-4.7, 55.5), "Singapore": (1.35, 103.8), "South Korea": (35.9, 127.8),
    "Spain": (40.5, -3.7), "Sweden": (60.1, 18.6), "Switzerland": (46.8, 8.2), "Thailand": (15.9, 100.99),
    "Turkey": (39.0, 35.2), "Ukraine": (48.4, 31.2), "United Arab Emirates": (23.4, 53.8),
    "United Kingdom": (54.0, -2.5), "United States": (39.5, -98.4), "Uruguay": (-32.5, -55.8), "Vietnam": (14.1, 108.3),
}


# %% 2. Carga e tratamento: o mesmo passo a passo do Power Query
FAIXAS_CAPTACAO = ["1. Até 50", "2. 50 a 200", "3. 200 a 1.000", "4. Acima de 1.000"]


def carregar_dados(caminho):
    if not caminho.is_file():
        raise FileNotFoundError(
            f"Não encontrei o CSV em: {caminho}\n"
            "Ele deve ficar na pasta 'dados', ao lado deste script."
        )

    # Consulta "layoffs_raw_data": ler o CSV no padrão americano (vírgula separa
    # colunas, ponto é decimal, datas mês/dia/ano) e converter os tipos.
    bruto = pd.read_csv(caminho, encoding="utf-8")
    bruto["date"] = pd.to_datetime(bruto["date"], format="%m/%d/%Y", errors="raise")
    for coluna in ["total_laid_off", "percentage_laid_off", "funds_raised"]:
        bruto[coluna] = pd.to_numeric(bruto[coluna], errors="raise")

    # Consulta "layoffs_raw_data (2)" (tratada):
    # 1) remover colunas sem uso; 2) manter 2022 em diante; 3) renomear.
    dados = bruto.drop(columns=["location", "source", "date_added"])
    dados = dados[dados["date"] >= "2022-01-01"].copy()
    dados = dados.rename(columns={
        "company": "Empresa", "total_laid_off": "Total de Demissões", "date": "Data",
        "percentage_laid_off": "Porcentagem de Demissões", "industry": "Indústria",
        "stage": "Estágio de Financiamento", "funds_raised": "Captação (US$ mi)", "country": "País",
    })

    # 4) tirar espaços e preencher vazios com "Unknown" nas colunas de texto
    for coluna in ["Empresa", "Indústria", "Estágio de Financiamento", "País"]:
        dados[coluna] = dados[coluna].astype("string").str.strip().replace("", pd.NA).fillna("Unknown")

    # 5) "UAE" e "United Arab Emirates" são o mesmo país
    dados["País"] = dados["País"].replace({"UAE": "United Arab Emirates"})

    # 6) colunas novas (as mesmas criadas no Power Query)
    dados["Ano"] = dados["Data"].dt.year
    dados["Nº Mês"] = dados["Data"].dt.month
    dados["Mês"] = dados["Data"].dt.to_period("M").dt.to_timestamp()
    captacao = dados["Captação (US$ mi)"]
    dados["Faixa de Captação"] = np.select(
        [captacao.isna(), captacao < 50, captacao < 200, captacao < 1000],
        ["Sem informação", FAIXAS_CAPTACAO[0], FAIXAS_CAPTACAO[1], FAIXAS_CAPTACAO[2]],
        default=FAIXAS_CAPTACAO[3],
    )
    dados["Fechamento"] = (dados["Porcentagem de Demissões"] == 1).astype(int)
    dados["Fechou?"] = np.where(dados["Fechamento"] == 1, "Fechou", "Não fechou")
    dados["Tipo de Empresa"] = np.where(dados["Estágio de Financiamento"] == "Post-IPO", "Capital aberto", "Privada")

    if dados.empty:
        raise ValueError("A base ficou vazia depois do tratamento. Revise o CSV.")
    return dados.sort_values("Data", ascending=False).reset_index(drop=True)


df = carregar_dados(ARQUIVO_CSV)

# Conferência com o Excel (TD_KPIs): 3.965 eventos, 841.703 demitidos, 29,2%
print(f"Base tratada: {len(df):,} eventos | {df['Total de Demissões'].sum():,.0f} demitidos | ".replace(",", ".")
      + f"% médio {df['Porcentagem de Demissões'].mean() * 100:.1f}%".replace(".", ","))


DATA_MIN, DATA_MAX = df["Data"].min().normalize(), df["Data"].max().normalize()
ANOS = sorted(df["Ano"].unique())


# %% 3a. Linha do tempo: períodos por nível (anos, trimestres, meses), como a do Excel
def periodos(nivel):
    """Lista de (rótulo, início, fim) de cada bloco da linha do tempo."""
    freq = {"anos": "YS", "trimestres": "QS", "meses": "MS"}[nivel]
    inicio = pd.Timestamp(f"{DATA_MIN.year}-01-01")
    blocos = []
    for comeco in pd.date_range(inicio, DATA_MAX, freq=freq):
        if nivel == "anos":
            fim, rotulo = comeco + pd.offsets.YearEnd(0), str(comeco.year)
        elif nivel == "trimestres":
            fim, rotulo = comeco + pd.offsets.QuarterEnd(0), f"T{comeco.quarter}/{str(comeco.year)[2:]}"
        else:
            fim, rotulo = comeco + pd.offsets.MonthEnd(0), comeco.strftime("%m/%y")
        blocos.append((rotulo, comeco, min(fim, DATA_MAX)))
    return blocos


def marcas(nivel):
    """Rótulos da linha do tempo (em trimestres e meses, só o primeiro bloco do ano mostra o ano)."""
    resultado = {}
    for i, (rotulo, comeco, _) in enumerate(periodos(nivel)):
        if nivel == "anos":
            texto = rotulo
        elif nivel == "trimestres":
            texto = str(comeco.year) if comeco.quarter == 1 else ""
        else:
            texto = str(comeco.year) if comeco.month == 1 else ""
        resultado[i] = {"label": texto, "style": {"fontSize": "11px", "color": "#5E646B", "whiteSpace": "nowrap"}}
    return resultado


def datas_da_linha(nivel, faixa):
    blocos = periodos(nivel)
    if not faixa:
        return DATA_MIN, DATA_MAX
    i, j = max(0, faixa[0]), min(len(blocos) - 1, faixa[1])
    return blocos[i][1], blocos[j][2]


# %% 3b. Funções de filtro, formatação e gráficos (cada groupby = uma tabela dinâmica do Excel)
def numero_br(valor, casas=0):
    if valor is None or pd.isna(valor):
        return "–"
    return f"{valor:,.{casas}f}".replace(",", "_").replace(".", ",").replace("_", ".")


def pct_br(valor, casas=1):
    return "–" if valor is None or pd.isna(valor) else numero_br(valor * 100, casas) + "%"


def filtrar_dados(dados, inicio, fim, tipos, industrias, estagios, paises):
    # Uma seleção vazia significa "todos", como nos controles da aula.
    mascara = pd.Series(True, index=dados.index)
    if inicio is not None:
        mascara &= dados["Data"] >= pd.Timestamp(inicio)
    if fim is not None:
        mascara &= dados["Data"] < pd.Timestamp(fim) + pd.Timedelta(days=1)
    for coluna, valores in [("Tipo de Empresa", tipos), ("Indústria", industrias),
                            ("Estágio de Financiamento", estagios), ("País", paises)]:
        if valores:
            mascara &= dados[coluna].isin(valores)
    return dados.loc[mascara]


FONTE = "Segoe UI, Helvetica, Arial, sans-serif"


def finalizar_figura(figura, altura, titulo=None, subtitulo=None, legenda=False, topo=None):
    """Estilo comum, imitando os gráficos do Excel: título centralizado, sem grade, fundo branco."""
    texto = f"<b>{titulo}</b>" if titulo else None
    figura.update_layout(
        template="plotly_white", height=altura, separators=",.",
        font=dict(family=FONTE, color=COR["texto"], size=12),
        title=dict(text=texto, x=0.5, xanchor="center", y=0.97, yanchor="top", font=dict(size=15),
                   subtitle=dict(text=subtitulo or "", font=dict(size=12, color=COR["apoio"]))) if texto else None,
        margin=dict(l=10, r=20, t=topo if topo is not None else (70 if subtitulo else 44 if titulo else 10), b=10),
        paper_bgcolor="white", plot_bgcolor="white", showlegend=legenda, legend_title_text="",
        hoverlabel=dict(bgcolor="white", font_size=12, font_family=FONTE),
    )
    figura.update_xaxes(showgrid=False, title=None)
    figura.update_yaxes(showgrid=False, title=None)
    if legenda:
        figura.update_layout(legend=dict(orientation="h", y=-0.12, x=0.5, xanchor="center", traceorder="normal"))
    return figura


def figura_vazia(altura, titulo=None, subtitulo=None):
    figura = go.Figure()
    figura.add_annotation(text="Sem dados para os filtros selecionados", x=0.5, y=0.5,
                          xref="paper", yref="paper", showarrow=False, font=dict(color=COR["apoio"]))
    figura.update_xaxes(visible=False)
    figura.update_yaxes(visible=False)
    return finalizar_figura(figura, altura, titulo, subtitulo)


def barras_horizontais(tabela, coluna, formato, altura, titulo, subtitulo=None, cinza=(), selecionados=None, rotulo_hover="", nota=None):
    """Barras horizontais com o maior em cima (o 'Categorias em ordem inversa' do Excel)."""
    if tabela.empty:
        return figura_vazia(altura, titulo, subtitulo)
    tabela = tabela.iloc[::-1]
    nomes = tabela.index.astype(str).tolist()
    figura = go.Figure(go.Bar(
        x=tabela[coluna], y=nomes, orientation="h",
        marker=dict(color=[COR["secundaria"] if n in cinza else COR["destaque"] for n in nomes],
                    opacity=[1 if not selecionados or n in selecionados else 0.35 for n in nomes]),
        text=[formato(v) for v in tabela[coluna]], textposition="outside", cliponaxis=False, constraintext="none",
        hovertemplate="<b>%{y}</b><br>" + rotulo_hover + "%{text}<extra></extra>",
    ))
    figura.update_layout(bargap=0.33)
    figura.update_xaxes(visible=False, range=[0, tabela[coluna].max() * 1.25])
    figura.update_yaxes(tickfont=dict(size=11), showline=True, linecolor=COR["neutro"])
    figura = finalizar_figura(figura, altura, titulo, subtitulo)
    if nota:
        figura.add_annotation(text=f"<i>{nota}</i>", x=0, xref="paper", y=-0.02, yref="paper", yanchor="top",
                              xanchor="left", showarrow=False, align="left", font=dict(size=10, color=COR["apoio"]))
        figura.update_layout(margin=dict(b=40))
    return figura


def colunas_anuais(serie, formato, titulo, selecionados):
    """Os 3 gráficos do 'Retrato por ano' (colunas com rótulo em todas, sem eixo vertical)."""
    serie = serie.reindex(ANOS)
    figura = go.Figure(go.Bar(
        x=[str(a) for a in serie.index], y=serie.values,
        marker=dict(color=COR["destaque"], opacity=[1 if not selecionados or a in selecionados else 0.35 for a in serie.index]),
        text=[formato(v) for v in serie.values], textposition="outside", cliponaxis=False, constraintext="none",
        hovertemplate="%{x}: <b>%{text}</b><extra></extra>",
    ))
    figura.update_layout(bargap=0.3)
    maior = np.nanmax(serie.values) if serie.notna().any() else 1
    figura.update_yaxes(visible=False, range=[0, maior * 1.2])
    figura.update_xaxes(type="category", showline=True, linecolor=COR["neutro"])
    return finalizar_figura(figura, 330, titulo)


def grafico_mensal(dados, inicio, fim):
    """'Demissões Por Mês': colunas com eixo em dois níveis (mês em cima, ano embaixo) + marcos da IA."""
    titulo = "Demissões Por Mês"
    if dados.empty:
        return figura_vazia(330, titulo)
    meses = pd.date_range(pd.Timestamp(inicio).to_period("M").to_timestamp(),
                          pd.Timestamp(fim).to_period("M").to_timestamp(), freq="MS")
    # Completar meses sem demissões, como o "Mostrar itens sem dados" do Excel
    mensal = dados.groupby("Mês")["Total de Demissões"].sum().reindex(meses, fill_value=0)
    eventos = dados.groupby("Mês").size().reindex(meses, fill_value=0)
    x = [[str(m.year) for m in meses], [str(m.month) for m in meses]]    # eixo em dois níveis
    picos = set(mensal.nlargest(2).index) if mensal.max() > 0 else set()
    figura = go.Figure(go.Bar(
        x=x, y=mensal.values, marker_color=COR["destaque"],
        text=[numero_br(v) if m in picos else "" for m, v in mensal.items()], textposition="outside", cliponaxis=False,
        constraintext="none", textfont=dict(size=12),
        customdata=np.stack([[m.strftime("%Y-%m-%d") for m in meses], [m.strftime("%m/%Y") for m in meses], eventos.values], axis=-1),
        hovertemplate="%{customdata[1]}<br><b>%{y:,.0f}</b> demitidos · %{customdata[2]} eventos<extra></extra>",
    ))
    topo = max(mensal.max(), 1)
    # Marcos da IA generativa: linha na posição do mês do lançamento
    cores = {"OpenAI": COR["secundaria"], "Anthropic": COR["destaque"], "Ambas": COR["texto"], "Outra": COR["apoio"]}
    posicoes = {m: i for i, m in enumerate(meses)}
    visiveis = [(d, n, e) for d, n, e in MARCOS_IA if pd.Timestamp(d).to_period("M").to_timestamp() in posicoes]
    for k, (data, nome, empresa) in enumerate(visiveis):
        i = posicoes[pd.Timestamp(data).to_period("M").to_timestamp()]
        figura.add_shape(type="line", xref="x", yref="paper", x0=i, x1=i, y0=0, y1=1,
                         line=dict(color=cores[empresa], width=1, dash="dot"))
        figura.add_annotation(x=i, xref="x", y=1.0 - (k % 3) * 0.075, yref="paper", text=nome, showarrow=False,
                              xanchor="left", yanchor="top", font=dict(size=10, color=cores[empresa]),
                              bgcolor="rgba(255,255,255,0.85)")
    figura.update_layout(bargap=0.25)
    figura.update_yaxes(range=[0, topo * 1.45], tickformat=",.0f", showline=False, tickfont=dict(size=10, color=COR["apoio"]))
    figura.update_xaxes(tickfont=dict(size=10), showdividers=True, dividercolor=COR["neutro"])
    figura = finalizar_figura(figura, 330, titulo, topo=40)
    figura.add_annotation(text="Linhas pontilhadas: lançamentos de IA generativa — <span style='color:#CC6A2A'>Anthropic</span> · "
                               "<span style='color:#4A5058'>OpenAI</span> · <span style='color:#8C9198'>outras</span>",
                          x=1, xref="paper", y=1.11, yref="paper", xanchor="right", showarrow=False,
                          font=dict(size=10, color=COR["apoio"]))
    return figura


def tabela_paises(dados, por_ano=False):
    chaves = ["Ano", "País"] if por_ano else ["País"]
    tabela = dados.groupby(chaves).agg(Demitidos=("Total de Demissões", "sum"), Eventos=("Empresa", "size")).reset_index()
    tabela = tabela[tabela["País"].isin(COORDENADAS)].copy()
    tabela["lat"] = tabela["País"].map(lambda p: COORDENADAS[p][0])
    tabela["lon"] = tabela["País"].map(lambda p: COORDENADAS[p][1])
    return tabela


def estilo_geo(figura):
    figura.update_geos(projection_type="natural earth", showframe=False, showcoastlines=False, showcountries=True,
                       countrycolor="white", landcolor="#E6E3DE", showocean=False, lataxis_range=[-58, 85], bgcolor="white")
    return figura


def mapa_bolhas(dados, paises_sel):
    titulo = "Demissões por país"
    if dados.empty:
        return figura_vazia(360, titulo)
    tabela = tabela_paises(dados).sort_values("Demitidos", ascending=False)
    maior = max(tabela["Demitidos"].max(), 1)
    textos = [numero_br(v) for v in tabela["Demitidos"]]
    rotulos = [t if i < 3 or p in (paises_sel or []) else "" for i, (t, p) in enumerate(zip(textos, tabela["País"]))]
    figura = go.Figure(go.Scattergeo(
        lat=tabela["lat"], lon=tabela["lon"], text=rotulos, mode="markers+text", textposition="middle center",
        textfont=dict(size=11, color=COR["texto"]),
        customdata=np.stack([tabela["País"], textos, tabela["Eventos"]], axis=-1),
        marker=dict(size=(np.sqrt(tabela["Demitidos"] / maior) * 60).clip(lower=5), color=COR["destaque"],
                    opacity=[0.75 if not paises_sel or p in paises_sel else 0.2 for p in tabela["País"]],
                    line=dict(color="white", width=1)),
        hovertemplate="<b>%{customdata[0]}</b><br>%{customdata[1]} demitidos · %{customdata[2]} eventos<extra></extra>",
    ))
    return finalizar_figura(estilo_geo(figura), 360, titulo)


def mapa_animado(dados):
    if dados.empty:
        return figura_vazia(400)
    tabela = tabela_paises(dados, por_ano=True).sort_values(["Ano", "Demitidos"], ascending=[True, False])
    figura = px.scatter_geo(tabela, lat="lat", lon="lon", size="Demitidos", hover_name="País", animation_frame="Ano",
                            hover_data={"lat": False, "lon": False, "Demitidos": ":,.0f", "Eventos": True, "Ano": False},
                            size_max=55, color_discrete_sequence=[COR["destaque"]])
    return finalizar_figura(estilo_geo(figura), 400)


def grafico_heatmap(dados):
    if dados.empty:
        return figura_vazia(470)
    top = dados.groupby("Indústria")["Total de Demissões"].sum().nlargest(10).index
    tabela = dados[dados["Indústria"].isin(top)].pivot_table(
        index="Indústria", columns="Ano", values="Total de Demissões", aggfunc="sum", fill_value=0)
    tabela = tabela.reindex(index=top, columns=ANOS, fill_value=0)
    percentual = tabela / tabela.sum().replace(0, np.nan)   # "% do Total da Coluna" (entre as 10 maiores)
    figura = go.Figure(go.Heatmap(
        z=percentual.values, x=[str(a) for a in percentual.columns], y=percentual.index.tolist(),
        colorscale=[[0, "#F7F5F2"], [1, COR["destaque"]]], showscale=False, xgap=3, ygap=3,
        text=[[pct_br(v) if pd.notna(v) else "" for v in linha] for linha in percentual.values],
        texttemplate="%{text}", textfont=dict(size=14),
        hovertemplate="<b>%{y}</b> · %{x}<br>%{text} dos demitidos do ano<extra></extra>",
    ))
    figura.update_yaxes(autorange="reversed", tickfont=dict(size=14, color=COR["texto"]))
    figura.update_xaxes(side="top", type="category", tickfont=dict(size=14, color=COR["texto"]))
    return finalizar_figura(figura, 470, topo=30)


def grafico_capital_privado(dados):
    titulo, sub = "Empresas de Capital Aberto × Privado", "% dos demitidos de cada ano, por tipo de empresa"
    if dados.empty:
        return figura_vazia(310, titulo, sub)
    tabela = dados.pivot_table(index="Ano", columns="Tipo de Empresa", values="Total de Demissões", aggfunc="sum", fill_value=0)
    tabela = tabela.reindex(columns=["Capital aberto", "Privada"], fill_value=0)
    percentual = tabela.div(tabela.sum(axis=1).replace(0, np.nan), axis=0)   # "% do Total da Linha"
    figura = go.Figure()
    for tipo, cor in [("Capital aberto", COR["destaque"]), ("Privada", COR["secundaria"])]:
        figura.add_bar(x=[str(a) for a in percentual.index], y=percentual[tipo], name=tipo, marker_color=cor,
                       text=[pct_br(v) for v in percentual[tipo]], textposition="inside", insidetextanchor="middle",
                       textfont=dict(color="white"), hovertemplate="%{x} · " + tipo + ": <b>%{text}</b><extra></extra>")
    figura.update_layout(barmode="stack", bargap=0.35)
    figura.update_yaxes(tickformat=".0%", range=[0, 1], tickvals=[0, 1], tickfont=dict(color=COR["apoio"]))
    figura.update_xaxes(type="category")
    return finalizar_figura(figura, 310, titulo, sub, legenda=True)


def grafico_por_tipo(dados, medida):
    if medida == "eventos":
        titulo, sub, nota = "Número de Eventos por Tipo de Empresa", None, None
    else:
        titulo, sub = "Fechamentos por tipo de empresa", "% dos cortes em que a empresa demitiu 100% do quadro"
        nota = None
    if dados.empty:
        return figura_vazia(310, titulo, sub)
    if medida == "eventos":
        tabela, formato = dados.groupby(["Ano", "Tipo de Empresa"]).size().unstack(fill_value=0), numero_br
    else:
        tabela, formato = dados.groupby(["Ano", "Tipo de Empresa"])["Fechamento"].mean().unstack(), pct_br
    tabela = tabela.reindex(columns=["Capital aberto", "Privada"])
    figura = go.Figure()
    for tipo, cor in [("Capital aberto", COR["destaque"]), ("Privada", COR["secundaria"])]:
        figura.add_bar(x=[str(a) for a in tabela.index], y=tabela[tipo], name=tipo, marker_color=cor,
                       text=[formato(v) for v in tabela[tipo]], textposition="outside", cliponaxis=False, constraintext="none",
                       hovertemplate="%{x} · " + tipo + ": <b>%{text}</b><extra></extra>")
    figura.update_layout(barmode="group", bargap=0.3)
    maior = np.nanmax(tabela.values) if np.isfinite(np.nanmax(tabela.values)) else 1
    figura.update_yaxes(visible=False, range=[0, maior * 1.25])
    figura.update_xaxes(type="category", showline=True, linecolor=COR["neutro"])
    figura = finalizar_figura(figura, 310, titulo, sub, legenda=True)
    if nota:
        figura.add_annotation(text=f"<i>{nota}</i>", x=0.5, xref="paper", y=-0.27, yref="paper", showarrow=False,
                              font=dict(size=10, color=COR["apoio"]))
        figura.update_layout(margin=dict(b=40))
    return figura


def grafico_dispersao(dados):
    privadas = dados[(dados["Tipo de Empresa"] == "Privada") & (dados["Captação (US$ mi)"] > 0)
                     & dados["Porcentagem de Demissões"].notna()]
    if privadas.empty:
        return figura_vazia(400)
    figura = px.scatter(privadas, x="Captação (US$ mi)", y="Porcentagem de Demissões", log_x=True, opacity=0.45,
                        hover_name="Empresa", custom_data=["Indústria", "País"], color_discrete_sequence=[COR["destaque"]])
    figura.update_traces(marker=dict(size=7), hovertemplate=(
        "<b>%{hovertext}</b><br>%{customdata[0]} · %{customdata[1]}<br>"
        "Captação: US$ %{x:,.0f} mi<br>Cortou %{y:.0%} do quadro<extra></extra>"))
    faixas = privadas.groupby("Faixa de Captação").agg(x=("Captação (US$ mi)", "median"), y=("Porcentagem de Demissões", "mean"))
    faixas = faixas.reindex(FAIXAS_CAPTACAO).dropna()
    figura.add_scatter(x=faixas["x"], y=faixas["y"], mode="lines+markers", line=dict(color=COR["texto"], width=2),
                       marker=dict(size=9), hovertemplate="Média da faixa: %{y:.0%}<extra></extra>")
    figura = finalizar_figura(figura, 400)
    figura.update_yaxes(tickformat=".0%", range=[0, 1.05], showgrid=True, gridcolor="#EEE")
    figura.update_xaxes(title="Captação total (US$ milhões, escala log)", showgrid=True, gridcolor="#EEE",
                        tickvals=[1, 10, 100, 1000, 10000, 100000], ticktext=["1", "10", "100", "1 mil", "10 mil", "100 mil"])
    return figura


def chamada_setores(dados):
    """Texto dinâmico — a 'chamada vinculada a célula' do Excel (tabelas!AR2)."""
    if dados.empty:
        return "Sem dados para os filtros selecionados."
    ranking = dados.groupby("Indústria")["Total de Demissões"].sum().sort_values(ascending=False)
    ranking = ranking[ranking.index != "Other"]
    if ranking.empty:
        return "No recorte selecionado só há empresas da categoria Other."
    return (f"Destaque do recorte: {ranking.index[0]} é a indústria com mais demitidos "
            f"({numero_br(ranking.iloc[0])}), sem contar a categoria Other.")


# %% 4. Layout Dash: o que será exibido na página
app = Dash(__name__, assets_folder=str(PASTA / "assets"), title="Layoffs Tech 2022–2026")
server = app.server  # usado para hospedar na web (ex.: Render com gunicorn)

# CSS embutido para manter tudo em um único .py (como no exemplo da aula).
# Sem fontes externas: o dashboard funciona sem internet.
app.index_string = """<!DOCTYPE html>
<html><head>{%metas%}<title>{%title%}</title>{%favicon%}{%css%}
<style>
:root{--Dash-Fill-Interactive-Strong:#CC6A2A;--Dash-Fill-Interactive-Weak:#F7EEE7;--Dash-Fill-Primary-Hover:#B45A1F;--Dash-Fill-Primary-Active:#B45A1F}
*{box-sizing:border-box} body{margin:0;background:#EFEDEA;color:#24272B;font-family:"Segoe UI",Helvetica,Arial,sans-serif}
.pagina{max-width:1600px;margin:auto;padding:0 20px 24px}
.cabecalho{display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:12px;background:white;padding:14px 20px;margin:0 -20px}
h1{margin:0;font-size:22px;letter-spacing:.3em;font-weight:700}
.sub{margin:4px 0 0;color:#8C9198;font-size:12.5px}
.nav .opcao{display:inline-block;margin-left:8px}
.nav input{display:none}
.nav label{display:inline-block;padding:7px 22px;border:1px solid #D9D5CF;border-radius:999px;font-weight:700;font-size:13px;cursor:pointer;background:white;color:#24272B}
.nav label:has(input:checked){background:#24272B;border-color:#24272B;color:white !important}
.nav label:has(input:checked) *{color:white !important}
.filtros{display:grid;grid-template-columns:1.55fr .8fr 1.2fr 1fr 1fr auto;gap:14px;align-items:start;background:white;border-radius:10px;padding:12px 16px;margin-top:14px}
.rotulo-filtro{display:flex;justify-content:space-between;align-items:center;font-size:12px;font-weight:700;margin-bottom:6px;border-bottom:1px solid #CC6A2A;padding-bottom:4px}
.niveis{display:flex;gap:4px}
.niveis input{display:none}
.niveis label{font-size:10.5px;font-weight:700;letter-spacing:.06em;color:#8C9198;padding:1px 6px;border-radius:4px;cursor:pointer}
.niveis label:has(input:checked){background:#F7EEE7;color:#B45A1F}
.limpar-mini{border:0;background:none;color:#8C9198;cursor:pointer;font-size:13px;padding:0 2px}
.limpar-mini:hover{color:#CC6A2A}
.linha-tempo{padding:6px 4px 0}
.linha-tempo .rc-slider-track{background:#CC6A2A}
.dash-range-slider-input,.dash-range-slider-min-input,.dash-range-slider-max-input{display:none !important}
.linha-tempo .rc-slider-handle{border-color:#CC6A2A}
.texto-periodo{font-size:11.5px;color:#B45A1F;font-weight:600;margin-top:2px}
.botao{background:#24272B;color:white;border:0;border-radius:999px;padding:9px 16px;font-weight:600;cursor:pointer;white-space:nowrap;margin-top:22px}
.botao:hover{background:#CC6A2A}
.indicadores{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-top:14px}
.card{background:white;border-radius:10px;padding:10px 18px}
.valor{font-size:28px;font-weight:700;line-height:1.15}
.rotulo{font-size:10.5px;letter-spacing:.14em;text-transform:uppercase;color:#8C9198}
.grade{display:grid;gap:14px;margin-top:14px}
.painel{background:white;border-radius:10px;padding:6px;min-width:0}
.chamada{background:#F7EEE7;color:#B45A1F;font-weight:700;border-radius:10px;padding:11px 16px;font-size:14px}
.nota{font-size:12px;color:#8C9198;font-style:italic;margin:4px 14px 8px}
.cartao-titulo{padding:10px 14px 0}.cartao-titulo h2{font-size:15px;margin:0}.cartao-titulo p{font-size:12px;color:#8C9198;margin:2px 0 0}
.rodape{font-size:12px;color:#8C9198;font-style:italic;margin:10px 2px 0}
@media(max-width:1100px){.filtros{grid-template-columns:1fr 1fr}.g-panorama,.g-setores,.g-empresas,.g-analises,.g-2x2,.g-anuais{grid-template-columns:1fr !important}}
</style></head><body>{%app_entry%}<footer>{%config%}{%scripts%}{%renderer%}</footer></body></html>"""

# O mapa usa o arquivo local assets/topojson/world_110m.json (funciona offline)
CONFIG_GRAFICO = {"displaylogo": False, "topojsonURL": app.get_asset_url("topojson/")}


def grafico(identificador, altura):
    # Altura fixa (como no exemplo da aula) e largura acompanhando o cartão
    return dcc.Graph(id=identificador, responsive=True, config=CONFIG_GRAFICO,
                     style={"height": f"{altura}px", "width": "100%"})


def cartao(identificador, altura, *extra):
    return html.Div([grafico(identificador, altura), *extra], className="painel")


def indicador(rotulo, identificador):
    return html.Div([html.Div(id=identificador, className="valor"), html.Div(rotulo, className="rotulo")], className="card")


def opcoes(coluna):
    valores = sorted(df[coluna].unique(), key=lambda v: (v == "Unknown", v))
    return [{"label": v, "value": v} for v in valores]


def filtro_lista(titulo, identificador, coluna, placeholder):
    return html.Div([html.Div(titulo, className="rotulo-filtro"),
                     dcc.Dropdown(id=identificador, options=opcoes(coluna), multi=True, value=[], placeholder=placeholder)])


NIVEL_INICIAL = "anos"
pagina_panorama = html.Div(id="pagina-panorama", children=[
    html.Div([cartao("g-mensal", 330)], className="grade"),
    html.Div([
        cartao("g-mapa", 360),
        cartao("g-ano-soma", 360), cartao("g-ano-eventos", 360), cartao("g-ano-media", 360),
    ], className="grade g-panorama", style={"gridTemplateColumns": "2.7fr 1fr 1fr 1fr"}),
])

pagina_setores = html.Div(id="pagina-setores", style={"display": "none"}, children=[
    html.Div([
        html.Div([
            html.Div(id="chamada-setores", className="chamada"),
            html.Div([
                cartao("g-ind-soma", 290, html.P("Other: categoria da fonte para empresas de múltiplos ramos ou fora das demais categorias.", className="nota")),
                cartao("g-ind-eventos", 290),
                cartao("g-ind-media", 290), cartao("g-ind-empresas", 290),
            ], className="grade g-2x2", style={"gridTemplateColumns": "1fr 1fr"}),
        ]),
        html.Div([
            html.Div([html.H2("Peso de cada indústria por ano"),
                      html.P("% dos demitidos de cada ano, entre as 10 maiores indústrias")], className="cartao-titulo"),
            grafico("g-heatmap", 470),
            html.P("Cor mais forte = indústria com maior fatia dos demitidos daquele ano. Cada coluna soma 100%.", className="nota"),
            html.P("Other: categoria da fonte para empresas de múltiplos ramos (68% dos demitidos vêm de Microsoft, Oracle, SAP e Ericsson).", className="nota"),
        ], className="painel"),
    ], className="grade g-setores", style={"gridTemplateColumns": "1.55fr 1fr"}),
])

pagina_empresas = html.Div(id="pagina-empresas", style={"display": "none"}, children=[
    html.Div([
        html.Div([cartao("g-empresas", 310), cartao("g-capital", 310)], className="grade", style={"marginTop": 0}),
        html.Div([cartao("g-estagio", 610, html.P("Post-IPO (cinza) = empresas de capital aberto. Series I e J têm menos de 10 eventos cada; a média pode variar bastante.", className="nota"))], className="grade", style={"marginTop": 0}),
        html.Div([cartao("g-eventos-tipo", 310), cartao("g-fechamentos", 286, html.P("Inclui saídas de um país (ex.: Deliveroo Austrália), não só falências.", className="nota"))], className="grade", style={"marginTop": 0}),
    ], className="grade g-empresas", style={"gridTemplateColumns": "1fr 1fr 1fr"}),
])

pagina_analises = html.Div(id="pagina-analises", style={"display": "none"}, children=[
    html.Div([
        html.Div([html.Div([html.H2("Captação × % do quadro cortado"),
                            html.P("Cada ponto é um corte de uma empresa privada · linha = média por faixa de captação")], className="cartao-titulo"),
                  grafico("g-dispersao", 400),
                  html.P("Empresas de capital aberto ficam fora: na fonte, o valor captado delas é inconsistente "
                         "(ex.: Microsoft aparece com US$ 1 mi).", className="nota")], className="painel"),
        html.Div([html.Div([html.H2("Demissões pelo mundo, ano a ano"),
                            html.P("Aperte ▶ para ver a evolução · respeita os filtros")], className="cartao-titulo"),
                  grafico("g-mapa-animado", 400)], className="painel"),
    ], className="grade g-analises", style={"gridTemplateColumns": "1fr 1fr"}),
    html.Div([
        html.Div([html.H2("Eventos recentes"), html.P("Os 15 cortes mais recentes dentro do recorte selecionado")], className="cartao-titulo"),
        dash_table.DataTable(
            id="tabela-recentes", page_size=15, sort_action="native", style_as_list_view=True,
            columns=[{"name": n, "id": n} for n in ["Data", "Empresa", "Indústria", "País", "Estágio", "Demitidos", "% do quadro"]],
            style_table={"padding": "0 14px 10px"},
            style_header={"fontWeight": "700", "color": COR["apoio"], "fontSize": 12, "backgroundColor": "white",
                          "borderBottom": f"1px solid {COR['neutro']}"},
            style_cell={"fontFamily": FONTE, "fontSize": 13, "padding": "6px 10px", "textAlign": "left", "border": "0",
                        "borderBottom": "1px solid #F0EEEA"},
            style_cell_conditional=[{"if": {"column_id": c}, "textAlign": "right"} for c in ["Demitidos", "% do quadro"]],
            style_data_conditional=[{"if": {"column_id": "Empresa"}, "fontWeight": "700"}],
        ),
    ], className="painel", style={"marginTop": 14}),
])

app.layout = html.Div([
    html.Div([
        html.Div([html.H1("LAYOFFS TECH 2022–2026"),
                  html.P("Demissões no setor de tecnologia · jan/2022 a out/2026 · fonte: layoffs.fyi", className="sub")]),
        dcc.RadioItems(id="navegacao", value="panorama", className="nav", labelClassName="opcao-label",
                       options=[{"label": n, "value": v} for n, v in
                                [("Panorama", "panorama"), ("Setores", "setores"), ("Empresas", "empresas"), ("Análises", "analises")]],
                       inline=True),
    ], className="cabecalho"),
    html.Div([
        html.Div([
            html.Div([html.Span("Data"),
                      html.Div([dcc.RadioItems(id="nivel-tempo", value=NIVEL_INICIAL, className="niveis", inline=True,
                                               options=[{"label": "ANOS", "value": "anos"}, {"label": "TRIMESTRES", "value": "trimestres"},
                                                        {"label": "MESES", "value": "meses"}]),
                                html.Button("✕", id="limpar-tempo", className="limpar-mini", title="Limpar filtro de data")],
                               style={"display": "flex", "alignItems": "center", "gap": "6px"})],
                     className="rotulo-filtro"),
            html.Div(dcc.RangeSlider(id="linha-tempo", min=0, max=len(periodos(NIVEL_INICIAL)) - 1, step=1,
                                     value=[0, len(periodos(NIVEL_INICIAL)) - 1], marks=marcas(NIVEL_INICIAL),
                                     allowCross=False, updatemode="mouseup"), className="linha-tempo"),
            html.Div(id="texto-periodo", className="texto-periodo"),
        ]),
        filtro_lista("Tipo de Empresa", "filtro-tipo", "Tipo de Empresa", "Todos"),
        filtro_lista("Indústria", "filtro-industria", "Indústria", "Todas"),
        filtro_lista("Estágio de Financiamento", "filtro-estagio", "Estágio de Financiamento", "Todos"),
        filtro_lista("País", "filtro-pais", "País", "Todos"),
        html.Button("Limpar filtros", id="botao-limpar", className="botao"),
    ], className="filtros"),
    html.Div([
        indicador("Demitidos reportados", "kpi-demitidos"),
        indicador("Eventos de demissão", "kpi-eventos"),
        indicador("% médio do quadro cortado", "kpi-pct"),
    ], className="indicadores"),
    dcc.Store(id="nivel-anterior", data=NIVEL_INICIAL),   # guarda o nível para converter o período ao trocar
    pagina_panorama, pagina_setores, pagina_empresas, pagina_analises,
    html.P(id="resumo-filtro", className="rodape"),
], className="pagina")


# %% 5a. Navegação entre as páginas (os botões do topo, como no Excel)
@app.callback(
    Output("pagina-panorama", "style"), Output("pagina-setores", "style"),
    Output("pagina-empresas", "style"), Output("pagina-analises", "style"),
    Input("navegacao", "value"),
)
def mostrar_pagina(pagina):
    return [{"display": "block" if pagina == p else "none"} for p in ["panorama", "setores", "empresas", "analises"]]


# %% 5b. Callback principal: filtros -> base filtrada -> indicadores e gráficos
@app.callback(
    Output("kpi-demitidos", "children"), Output("kpi-eventos", "children"), Output("kpi-pct", "children"),
    Output("texto-periodo", "children"), Output("resumo-filtro", "children"),
    Output("g-mensal", "figure"), Output("g-mapa", "figure"),
    Output("g-ano-soma", "figure"), Output("g-ano-eventos", "figure"), Output("g-ano-media", "figure"),
    Output("chamada-setores", "children"),
    Output("g-ind-soma", "figure"), Output("g-ind-eventos", "figure"), Output("g-ind-media", "figure"),
    Output("g-ind-empresas", "figure"), Output("g-heatmap", "figure"),
    Output("g-empresas", "figure"), Output("g-estagio", "figure"), Output("g-capital", "figure"),
    Output("g-eventos-tipo", "figure"), Output("g-fechamentos", "figure"),
    Output("g-dispersao", "figure"), Output("g-mapa-animado", "figure"), Output("tabela-recentes", "data"),
    Input("linha-tempo", "value"), State("nivel-tempo", "value"),
    Input("filtro-tipo", "value"), Input("filtro-industria", "value"),
    Input("filtro-estagio", "value"), Input("filtro-pais", "value"),
)
def atualizar_dashboard(faixa, nivel, tipos, industrias, estagios, paises):
    inicio, fim = datas_da_linha(nivel, faixa)
    dados = filtrar_dados(df, inicio, fim, tipos, industrias, estagios, paises)
    periodo_todo = inicio <= DATA_MIN and fim >= DATA_MAX
    anos_sel = [] if periodo_todo else list(range(inicio.year, fim.year + 1))

    # Indicadores (= TD_KPIs)
    kpis = (numero_br(dados["Total de Demissões"].sum()), numero_br(len(dados)), pct_br(dados["Porcentagem de Demissões"].mean()))
    texto_periodo = "Todos os períodos" if periodo_todo else f"{inicio:%d/%m/%Y} a {fim:%d/%m/%Y}"
    resumo = (f"{numero_br(len(dados))} de {numero_br(len(df))} eventos no recorte. Fonte: dados/layoffs.csv (layoffs.fyi via Kaggle). "
              "Clique nos gráficos para filtrar; clique de novo no mesmo item para tirar o filtro.")

    # Panorama — os anuais mostram todos os anos e destacam o período escolhido (como o eixo do Excel)
    todos_anos = filtrar_dados(df, None, None, tipos, industrias, estagios, paises)
    mensal = grafico_mensal(dados, inicio, fim)
    mapa = mapa_bolhas(dados, paises)
    ano_soma = colunas_anuais(todos_anos.groupby("Ano")["Total de Demissões"].sum(), numero_br, "Total Anual De Demissões", anos_sel)
    ano_eventos = colunas_anuais(todos_anos.groupby("Ano").size(), numero_br, "Nº De Eventos De Demissões", anos_sel)
    ano_media = colunas_anuais(todos_anos.groupby("Ano")["Total de Demissões"].mean(), numero_br, "Média Anual De Demissões", anos_sel)

    # Setores — os rankings ignoram o próprio filtro de indústria para destacar a seleção sem esconder o resto
    dados_ind = filtrar_dados(df, inicio, fim, tipos, None, estagios, paises)
    por_ind = dados_ind.groupby("Indústria").agg(Soma=("Total de Demissões", "sum"), Eventos=("Empresa", "size"),
                                                 Media=("Total de Demissões", "mean"), Empresas=("Empresa", "nunique"))
    sub = " · no recorte selecionado"
    ind_soma = barras_horizontais(por_ind.nlargest(10, "Soma"), "Soma", numero_br, 290, "Quantas pessoas", "Soma de demitidos" + sub,
                                  cinza=("Other",), selecionados=industrias, rotulo_hover="Demitidos: ")
    ind_eventos = barras_horizontais(por_ind.nlargest(10, "Eventos"), "Eventos", numero_br, 290, "Quantas vezes",
                                     "Nº de eventos de demissão" + sub, cinza=("Other",), selecionados=industrias, rotulo_hover="Eventos: ")
    ind_media = barras_horizontais(por_ind.dropna(subset=["Media"]).nlargest(10, "Media"), "Media", numero_br, 290, "Tamanho do corte",
                                   "Média de demitidos por evento" + sub, cinza=("Other",), selecionados=industrias, rotulo_hover="Média: ")
    ind_empresas = barras_horizontais(por_ind.nlargest(10, "Empresas"), "Empresas", numero_br, 290, "Quantas empresas",
                                      "Número de empresas em cada indústria" + sub, cinza=("Other",), selecionados=industrias,
                                      rotulo_hover="Empresas: ")
    heatmap = grafico_heatmap(dados)

    # Empresas
    top_emp = dados.groupby("Empresa")["Total de Demissões"].sum().nlargest(10).to_frame("Soma")
    empresas = barras_horizontais(top_emp[top_emp["Soma"] > 0], "Soma", numero_br, 310, "Quem",
                                  "Top 10 empresas por número de demissões" + sub, rotulo_hover="Demitidos: ")
    dados_est = filtrar_dados(df, inicio, fim, tipos, industrias, None, paises)
    est = dados_est.groupby("Estágio de Financiamento")["Porcentagem de Demissões"].mean().dropna().sort_values(ascending=False)
    estagio = barras_horizontais(est.to_frame("Media"), "Media", pct_br, 634, "Intensidade do corte",
                                 "% média de funcionários demitidos a cada corte, por estágio de financiamento",
                                 cinza=("Post-IPO",), selecionados=estagios, rotulo_hover="% médio cortado: ")
    capital = grafico_capital_privado(dados)
    eventos_tipo = grafico_por_tipo(dados, "eventos")
    fechamentos = grafico_por_tipo(dados, "fechamentos")

    # Análises (só no Python)
    recentes = [{"Data": r["Data"].strftime("%d/%m/%Y"), "Empresa": r["Empresa"], "Indústria": r["Indústria"], "País": r["País"],
                 "Estágio": r["Estágio de Financiamento"],
                 "Demitidos": "" if pd.isna(r["Total de Demissões"]) else numero_br(r["Total de Demissões"]),
                 "% do quadro": "" if pd.isna(r["Porcentagem de Demissões"]) else pct_br(r["Porcentagem de Demissões"], 0)}
                for _, r in dados.head(15).iterrows()]

    return (*kpis, texto_periodo, resumo, mensal, mapa, ano_soma, ano_eventos, ano_media, chamada_setores(dados),
            ind_soma, ind_eventos, ind_media, ind_empresas, heatmap, empresas, estagio, capital, eventos_tipo, fechamentos,
            grafico_dispersao(dados), mapa_animado(dados), recentes)


# %% 5c. Controles: nível da linha do tempo, cliques nos gráficos e botões de limpar
GRAFICOS_CLICAVEIS = ["g-mapa", "g-mensal", "g-ano-soma", "g-ano-eventos", "g-ano-media", "g-ind-soma",
                      "g-ind-eventos", "g-ind-media", "g-ind-empresas", "g-heatmap", "g-estagio"]


def alternar(lista, valor):
    """Clicar num item já filtrado tira o filtro; senão, filtra só por ele."""
    return [] if (lista or []) == [valor] else [valor]


def indice_periodo(nivel, data):
    for i, (_, comeco, fim) in enumerate(periodos(nivel)):
        if comeco <= data <= fim:
            return i
    return 0


@app.callback(
    Output("nivel-tempo", "value"), Output("linha-tempo", "max"), Output("linha-tempo", "marks"), Output("linha-tempo", "value"),
    Output("filtro-tipo", "value"), Output("filtro-industria", "value"), Output("filtro-estagio", "value"), Output("filtro-pais", "value"),
    *[Output(g, "clickData") for g in GRAFICOS_CLICAVEIS], Output("nivel-anterior", "data"),
    Input("nivel-tempo", "value"), Input("limpar-tempo", "n_clicks"), Input("botao-limpar", "n_clicks"),
    *[Input(g, "clickData") for g in GRAFICOS_CLICAVEIS],
    State("linha-tempo", "value"), State("nivel-anterior", "data"),
    State("filtro-tipo", "value"), State("filtro-industria", "value"), State("filtro-estagio", "value"), State("filtro-pais", "value"),
    prevent_initial_call=True,
)
def controles(nivel, _limpar_tempo, _limpar_tudo, *args):
    cliques = dict(zip(GRAFICOS_CLICAVEIS, args[:len(GRAFICOS_CLICAVEIS)]))
    faixa, anterior, tipos, industrias, estagios, paises = args[len(GRAFICOS_CLICAVEIS):]
    origem = ctx.triggered_id
    zerar = [None] * len(GRAFICOS_CLICAVEIS)

    def linha(nv, i=None, j=None):
        n = len(periodos(nv))
        i = 0 if i is None else i
        j = n - 1 if j is None else j
        return nv, n - 1, marcas(nv), [i, j]

    def selecionar(nv, data):
        # Clicar de novo no mesmo período volta para "todos os períodos"
        k = indice_periodo(nv, data)
        if nv == nivel and faixa == [k, k]:
            return linha(nv)
        return linha(nv, k, k)

    tempo = (no_update,) * 4
    if origem == "botao-limpar":
        return (*linha("anos"), [], [], [], [], *zerar, "anos")
    if origem == "limpar-tempo":
        tempo = linha(nivel)
    elif origem == "nivel-tempo":
        # Ao trocar o nível, mantém o mesmo intervalo de datas convertido para os novos blocos
        inicio, fim = datas_da_linha(anterior or "anos", faixa)
        tempo = linha(nivel, indice_periodo(nivel, inicio), indice_periodo(nivel, fim))
    elif origem in cliques and cliques[origem]:
        ponto = cliques[origem]["points"][0]
        if origem == "g-mapa":
            paises = alternar(paises, ponto["customdata"][0])
        elif origem == "g-mensal":
            tempo = selecionar("meses", pd.Timestamp(ponto["customdata"][0]))
        elif origem in ("g-ano-soma", "g-ano-eventos", "g-ano-media"):
            tempo = selecionar("anos", pd.Timestamp(f"{ponto['x']}-01-01"))
        elif origem in ("g-ind-soma", "g-ind-eventos", "g-ind-media", "g-ind-empresas"):
            industrias = alternar(industrias, ponto["y"])
        elif origem == "g-heatmap":
            industrias = alternar(industrias, ponto["y"])
            tempo = linha("anos", *(2 * [indice_periodo("anos", pd.Timestamp(f"{ponto['x']}-01-01"))])) if industrias else linha("anos")
        elif origem == "g-estagio":
            estagios = alternar(estagios, ponto["y"])
    else:
        return (no_update,) * 8 + tuple(zerar) + (no_update,)
    return (*tempo, tipos, industrias, estagios, paises, *zerar, tempo[0] if tempo[0] is not no_update else no_update)


# %% 6. Iniciar o servidor local
if __name__ == "__main__":
    print(f"Abra no navegador: http://127.0.0.1:{PORTA}  (Ctrl+C para encerrar)")
    # use_reloader=False evita iniciar o programa duas vezes em IDEs.
    app.run(debug=False, use_reloader=False, host="127.0.0.1", port=PORTA)