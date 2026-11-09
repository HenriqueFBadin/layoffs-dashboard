# Layoffs Tech 2022–2026 — Dashboard em Excel e Python

APS 1 · Dashboards em Excel e Python · Insper 2026/2
Aluno: Henrique Fazzio Badin · Prof.: Rogério da Costa Monteiro

Dashboard sobre demissões (layoffs) no setor de tecnologia entre janeiro de 2022 e outubro de 2026, período que cobre a retomada pós-pandemia e a ascensão da IA generativa. O mesmo dashboard existe em duas versões: **Excel** (tabelas dinâmicas, segmentações e linha do tempo) e **Python** (Dash + Plotly), que porta o Excel e acrescenta interatividade.

Dados: [Layoffs Dataset (Kaggle)](https://www.kaggle.com/datasets/swaptr/layoffs-2022), que consolida os anúncios registrados pelo site layoffs.fyi.

---

## Estrutura da pasta

```
├── dashboard_layoffs.py        ← dashboard em Python (Dash + Plotly)
├── requirements.txt            ← bibliotecas necessárias
├── dados/
│   └── layoffs.csv             ← base original (4.644 linhas, mar/2020 a out/2026)
├── assets/
│   └── topojson/world_110m.json← desenho do mapa-múndi (permite usar o mapa offline)
├── excel/
│   └── (dashboard em Excel)
└── README.md
```

---

## Como rodar o dashboard em Python

Precisa de **Python 3.9 ou mais novo**. Funciona em Windows, macOS e Linux.

1. **Baixe a pasta** (botão *Code → Download ZIP* no GitHub) e descompacte.
2. **Abra um terminal dentro da pasta** (no Windows: clique com o botão direito na pasta → *Abrir no Terminal*).
3. **Instale as bibliotecas** (só na primeira vez; precisa de internet):
   ```
   pip install -r requirements.txt
   ```
4. **Execute:**
   ```
   python dashboard_layoffs.py
   ```
5. **Abra no navegador:** http://127.0.0.1:8050

Para encerrar, volte ao terminal e aperte **Ctrl+C**.

Opcional, para não misturar com outras instalações de Python: crie um ambiente virtual antes do passo 3 com `python -m venv .venv` e ative com `.venv\Scripts\activate` (Windows) ou `source .venv/bin/activate` (macOS/Linux).

### Funciona offline

Depois que as bibliotecas estão instaladas, o dashboard **não precisa de internet**: os dados estão em `dados/`, as bibliotecas JavaScript vêm instaladas com o Dash e o mapa usa o arquivo local em `assets/topojson/`. As fontes usadas são as do sistema.

Para instalar num computador **sem internet**, baixe os pacotes antes num computador conectado (`pip download -r requirements.txt -d pacotes`), leve a pasta `pacotes` junto e instale com `pip install --no-index --find-links pacotes -r requirements.txt`.

### Problemas comuns

| Sintoma | Solução |
|---|---|
| `python` ou `pip` não é reconhecido | No Windows, use `py dashboard_layoffs.py` e `py -m pip install -r requirements.txt` |
| `Address already in use` / porta ocupada | Já existe um dashboard rodando: feche o outro terminal ou mude `PORTA = 8050` no início do script |
| `Não encontrei o CSV` | O arquivo precisa estar em `dados/layoffs.csv`, ao lado do script |
| Página em branco | Espere alguns segundos e recarregue (F5); confira se o terminal ainda está rodando |

---

## Como abrir o dashboard em Excel

Abra o arquivo da pasta `excel/` no **Excel para Windows (Microsoft 365, 2021 ou 2024)**. O dashboard abre na página **Panorama**. Os botões no topo levam às páginas **Setores** e **Empresas**, e os filtros (linha do tempo e segmentações) valem para as três páginas.

- **Funciona offline.** O mapa guarda o desenho dos países dentro do próprio arquivo.
- **Não é preciso clicar em "Atualizar Tudo".** As tabelas dinâmicas já têm os dados. O Power Query lê o CSV de uma pasta do computador do autor; num outro computador, "Atualizar Tudo" daria erro de caminho, mas isso não afeta filtros nem gráficos.
- Para ver os bastidores (dados, Power Query, tabelas dinâmicas): *Arquivo → Opções → Avançado → Mostrar guias de planilha* e reexiba as abas ocultas.

---

## O que o dashboard mostra

| Página | Conteúdo |
|---|---|
| **Indicadores (topo)** | Demitidos reportados · eventos de demissão · % médio do quadro cortado |
| **Panorama** | Demissões por mês · mapa por país · total anual, nº de eventos e média anual de demissões |
| **Setores** | Indústrias por demitidos, eventos, tamanho do corte e nº de empresas · peso de cada indústria por ano (mapa de calor) · destaque dinâmico do recorte |
| **Empresas** | Top 10 empresas · capital aberto × privado · intensidade do corte por estágio · eventos e fechamentos por tipo de empresa |
| **Análises** (só Python) | Captação × % do quadro cortado · mapa animado ano a ano · eventos recentes |

**Filtros:** linha do tempo (data), tipo de empresa, indústria, estágio de financiamento e país — valem para todas as páginas. No Python, **clicar nos gráficos também filtra**: um país no mapa, um mês, um ano, uma indústria, uma célula do mapa de calor ou um estágio. Clicar de novo tira o filtro.

### Excel × Python

As páginas Panorama, Setores e Empresas têm os mesmos gráficos, títulos, cores e disposição nas duas versões. Diferenças intencionais no Python:

- **Linha do tempo:** uma barra deslizante com os níveis anos, trimestres e meses (no Excel, a Linha do Tempo).
- **Filtros em listas suspensas** (no Excel, segmentações em botões).
- **Mapa de bolhas** (no Excel, mapa preenchido por cor).
- **Marcos da IA generativa** no gráfico mensal: ChatGPT (nov/22), GPT-4 e primeiro Claude (mar/23), Claude 3 (mar/24), GPT-4o (mai/24), Claude 3.5 Sonnet (jun/24), DeepSeek-R1 (jan/25), Claude 4 (mai/25), GPT-5 (ago/25), Claude Opus 4.5 (nov/25) e Claude Fable 5 (jun/26). Laranja = Anthropic, cinza-escuro = OpenAI. A lista fica em `MARCOS_IA`, no início do script.
- **Página Análises** (dispersão captação × % cortado, mapa animado e eventos recentes).

### Definições

- **Evento** = um anúncio de corte de uma empresa numa data (uma linha da base).
- **Demitidos** = soma de `total_laid_off`, só dos eventos que informaram o número (cerca de 64%).
- **% médio do quadro cortado** = média simples de `percentage_laid_off` entre os eventos que informaram.
- **Fechamento** = evento em que a empresa demitiu 100% do quadro (inclui saídas de um país).
- **Capital aberto** = estágio Post-IPO; os demais estágios = **Privada**.
- **Other** = categoria genérica da fonte; 68% dos seus demitidos vêm de Microsoft, Oracle, SAP e Ericsson.

### Tratamento dos dados (igual no Power Query e no pandas)

1. Ler o CSV no padrão americano (datas mês/dia/ano, ponto decimal).
2. Remover `location`, `source` e `date_added`.
3. Manter eventos a partir de 01/01/2022 (3.965 eventos).
4. Renomear as colunas para português.
5. Tirar espaços sobrando e preencher vazios de texto com `Unknown`.
6. Unificar `UAE` e `United Arab Emirates`.
7. Criar Ano, Nº Mês, Mês, Faixa de Captação, Fechou?/Fechamento e Tipo de Empresa.

Conferência: **3.965 eventos · 841.703 demitidos · 29,2% de corte médio**, nas duas versões.

---

## Bibliotecas

[Dash](https://dash.plotly.com/) e [Plotly](https://plotly.com/python/) para o dashboard, [pandas](https://pandas.pydata.org/) para o tratamento dos dados. Testado com dash 3.0.4 / plotly 6.0.1 / pandas 2.2.3 e com dash 4.5.0 / plotly 7.1.0 / pandas 3.0.5.
