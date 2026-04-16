"""
Dashboard Financeiro - Federacao Paulista de Golfe
Razao Contabil x Orcamento Aprovado - 2026
Servidor local: http://localhost:8050
"""

import os
import io
import base64
from datetime import datetime
import pandas as pd
import numpy as np
from dash import Dash, html, dcc, dash_table, callback, Input, Output, State, no_update
import plotly.graph_objects as go
import plotly.express as px
from flask import send_file, request as flask_request

# ============================================================
# CARREGAR DADOS
# ============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# --- Razao Contabil (realizado) ---
RAZAO_PATH = os.path.join(BASE_DIR, "razao.xlsx")

df_razao_raw = pd.read_excel(RAZAO_PATH, sheet_name="Planilha1")

# Usar indices para evitar problemas de encoding
_cols = list(df_razao_raw.columns)
COL_EMPRESA = _cols[0]
COL_MES = _cols[1]
COL_CONTA = _cols[2]
COL_CONTA_RES = _cols[3]
COL_DATA = _cols[7]
COL_HIST = _cols[9]
COL_COMPLEMENTO = _cols[10]
COL_FORNECEDOR = _cols[13]
COL_VLR = _cols[18]

# Limpar: remover linhas de resumo e sem conta
df_razao = df_razao_raw.dropna(subset=[COL_CONTA]).copy()
df_razao = df_razao[~df_razao[COL_HIST].astype(str).str.contains("SALDO ANTERIOR|TOTAL CONTA", na=False)]
df_razao = df_razao[df_razao[COL_VLR].notna()].copy()

# Classificar receita vs despesa
df_razao["Tipo"] = df_razao[COL_CONTA].apply(lambda x: "Receita" if str(x).startswith("3") else "Despesa")

# Extrair codigo detalhe (ultimo bloco do codigo conta)
df_razao["CodDetalhe"] = df_razao[COL_CONTA].apply(lambda x: str(x).split(".")[-1] if "." in str(x) else str(x))

# Extrair grupo (5 primeiros niveis)
df_razao["GrupoConta"] = df_razao[COL_CONTA].apply(lambda x: ".".join(str(x).split(".")[:5]) if "." in str(x) else str(x))

# Datas
df_razao[COL_DATA] = pd.to_datetime(df_razao[COL_DATA], errors="coerce")
df_razao["Ano"] = df_razao[COL_DATA].dt.year
df_razao["MesNum"] = df_razao[COL_DATA].dt.month
df_razao["AnoMes"] = df_razao[COL_DATA].dt.strftime("%Y-%m")

# Valor absoluto para receitas (no razao sao negativos = creditos)
df_razao["Valor"] = df_razao.apply(
    lambda r: abs(r[COL_VLR]) if r["Tipo"] == "Receita" else r[COL_VLR], axis=1
)

# --- De/Para: Centro de Custo e Projeto ---
COL_CR = _cols[15]  # CR
DEPARA = {
    100000: ("FPG", "SAPEZAL"), 100001: ("CE", "CENTRO ESPORTIVO"), 100002: ("FPG", "FEDERAÇÃO PAULISTA DE GOLFE"),
    200001: ("FPG", "TORNEIO ABERTURA DO SAPEZAL GOLFE CLUBE"), 200002: ("FPG", "TORNEIO INTERNO DE DUPLAS SAPEZAL GOLFE CLUBE"),
    200003: ("FPG", "CAMPEONATO ABERTO SAPEZAL GOLFE CLUBE"), 200004: ("FPG", "EVENTO DE ENCERRAMENTO SAPEZAL GOLFE CLUBE"),
    200005: ("FPG", "Torneio Interno Sapezal"),
    300000: ("FPG", "CAMPEONATO BANDEIRANTES"), 300001: ("FPG", "CAMPEONATO ABERTO DO ESTADO DE SÃO PAULO"),
    300002: ("FPG", "CAMPEONATO ABERTO SENIORS DO ESTADO DE SP"), 300003: ("FPG", "CAMPEONATO INTERCLUBES SCRATCH MASCULINO"),
    300004: ("FPG", "CAMPEONATO INTERCLUBES HCP INDEX MASCULINO"), 300005: ("FPG", "CAMPEONATO ABERTO DE DUPLAS MASCULINO"),
    300006: ("FPG", "CAMPEONATO INTERCLUBES SCRATCH FEMININO"), 300008: ("FPG", "CAMPEONATO ABERTO DE DUPLAS FEMININAS"),
    300009: ("FPG", "TAÇA ESCUDO"), 300010: ("FPG", "CAMPEONATO PAULISTA MATCH PLAY SCRATCH"),
    300011: ("FPG", "CAMPEONATO PAULISTA MATCH PLAY HCP INDEX"),
    300013: ("JUV", "INTERFEDERACOES"), 300015: ("JUV", "DESENVOLVIMENTO DE GOLFE JUVENIL"),
    300016: ("JUV", "TOUR NACIONAL JUVENIL 1ª ETAPA"), 300017: ("JUV", "TOUR NACIONAL JUVENIL 2ª ETAPA"),
    300018: ("JUV", "TOUR NACIONAL JUVENIL 3ª ETAPA"), 300019: ("JUV", "TOUR NACIONAL JUVENIL 4ª ETAPA"),
    300020: ("JUV", "TOUR NACIONAL JUVENIL 5ª ETAPA"), 300021: ("JUV", "TOUR NACIONAL JUVENIL 4ª ETAPA"),
    300022: ("JUV", "TOUR NACIONAL JUVENIL 6ª ETAPA"),
    300024: ("JUV", "CAMPEONATO BRASILEIRO JUVENIL"),
    300025: ("JUV", "TORNEIO JUVENIL DO ESTADO DE SP - 1ª ETAPA (KIDS)"),
    300026: ("JUV", "TORNEIO JUVENIL DO ESTADO DE SP - 2ª ETAPA (KIDS)"),
    300027: ("JUV", "TORNEIO JUVENIL DO ESTADO DE SP - 3ª ETAPA (KIDS)"),
    300028: ("JUV", "TORNEIO JUVENIL DO ESTADO DE SP - 4ª ETAPA (KIDS)"),
    300029: ("JUV", "TORNEIO JUVENIL DO ESTADO DE SP - 5ª ETAPA (KIDS)"),
    300030: ("JUV", "TORNEIO JUVENIL DO ESTADO DE SP - 6ª ETAPA (KIDS)"),
    300032: ("JUV", "CAMPEONATO DE VERÃO JUVENIL DO ESTADO DE SP"),
    300033: ("JUV", "CAMPEONATO DE INVERNO JUVENIL DO ESTADO DE SP"),
    300034: ("JUV", "CAMPEONATO ABERTO JUVENIL DO ESTADO DE SP"),
    300036: ("JUV", "TORNEIO JUVENIL CORDOBA"), 300037: ("JUV", "TORNEIO JUVENIL CHILE"),
    300038: ("JUV", "TAÇA YOSHITO NOMURA"),
    300040: ("FPG", "CAMPEONATO LATINO AMERICANO"), 300041: ("FPG", "ABERTO NIKKEY"),
    300042: ("FPG", "CASA DA PAZ"), 300044: ("FPG", "GOLFE SOLIDÁRIO- 1ªETAPA"),
    300049: ("FPG", "ZUP JUNIOR PRO GOLF TOUR"),
    400000: ("CE", "EVENTO CYRELA"), 400001: ("CE", "EVENTO APOLO ENERGIA"),
    400002: ("CE", "EVENTO BEM-FIRENZE"), 400003: ("CE", "EVENTO WTC BUSINESS CLUB"),
    400004: ("CE", "EVENTO PARTNERONE"), 400005: ("CE", "EVENTO ADD VALUE"),
    400006: ("CE", "EVENTO CATERPILAR"), 400007: ("CE", "EVENTO AGÊNCIA MAFÊ"),
    400008: ("CE", "EVENTO BRADASCHIA"), 400009: ("CE", "EVENTO MULHERES & GOLFE"),
    400010: ("CE", "EVENTO DIA DAS CRIANÇAS"), 400011: ("CE", "EVENTO IVANILDO DE LIMA"),
    400012: ("CE", "GIOVANA GUIMARAES DA SILVA"), 400013: ("CE", "GMC CONSULTÓRIO"),
    400014: ("CE", "AGILITY-TD SYNNEX"), 400015: ("CE", "SYNGENTA"),
    400016: ("CE", "SAO PAULO FUTEBOL CLUBE"), 400017: ("CE", "LS LEAN SALES"),
    400018: ("CE", "MADI EVENTOS"), 400019: ("CE", "PROLOGIS LOGISTICA"),
    400020: ("CE", "EDER RODRIGUES ROGATTI"), 400021: ("CE", "FRANK WISBRUN SÃO PAULO GOLF CLUB"),
    999999: ("FPG", "GERAL"),
}

# Exceções: Conta Resumida 300045, 300040, 300041 -> JUV
CONTAS_JUV = {"300045", "300040", "300041"}

def mapear_cc(row):
    conta_res = str(row[COL_CONTA_RES]).strip()
    if conta_res in CONTAS_JUV:
        return "JUV"
    cr = row[COL_CR]
    if pd.notna(cr):
        return DEPARA.get(int(cr), ("FPG", ""))[0]
    return "FPG"

def mapear_projeto(row):
    cr = row[COL_CR]
    if pd.notna(cr):
        return DEPARA.get(int(cr), ("", "SEM PROJETO"))[1]
    return "SEM PROJETO"

df_razao["CentroCusto"] = df_razao.apply(mapear_cc, axis=1)
df_razao["Projeto"] = df_razao.apply(mapear_projeto, axis=1)

# Listas para filtros
CC_DISP = sorted(df_razao["CentroCusto"].unique().tolist())
PROJ_DISP = sorted(df_razao["Projeto"].unique().tolist())

# --- Orcamento Aprovado ---
ORC_PATH = os.path.join(BASE_DIR, "orcamento.xlsx")

df_orc_raw = pd.read_excel(ORC_PATH, sheet_name="ENTIDADE", header=1)
_orc_cols = list(df_orc_raw.columns)

# Construir mapeamento conta -> descricao e orcamento mensal
MESES_ORC = {"JAN": 1, "FEV": 2, "MAR": 3, "ABR": 4, "MAI": 5, "JUN": 6,
             "JUL": 7, "AGO": 8, "SET": 9, "OUT": 10, "NOV": 11, "DEZ": 12}

# Mapear grupos e contas
grupo_desc = {}  # grupo_code -> descricao
conta_grupo = {}  # cod_detalhe -> grupo_code
conta_desc = {}  # cod_detalhe -> descricao
orc_values = []  # lista de {CodDetalhe, MesNum, ValorOrc}

current_group = ""
current_group_desc = ""

for _, row in df_orc_raw.iterrows():
    consol = row[_orc_cols[1]]  # CONSOLIDADO
    detail = row[_orc_cols[2]]  # Unnamed:2 (subconta ou descricao do grupo)
    desc = row[_orc_cols[3]]    # DESCRICAO DA CONTA

    # Linha de grupo (CONSOLIDADO tem o codigo, Unnamed:2 tem a descricao)
    if pd.notna(consol) and isinstance(consol, str) and "." in consol:
        current_group = consol.strip()
        current_group_desc = str(detail).strip() if pd.notna(detail) else consol
        # Limpar encoding issues
        current_group_desc = current_group_desc.replace("\ufffd", "A").replace("\x00", "")
        grupo_desc[current_group] = current_group_desc
        continue
    # Grupo nivel superior (ex: "3" com "RECEITAS TOTAIS", "4" etc.)
    if pd.notna(consol) and isinstance(consol, str) and consol.strip().isdigit():
        top_desc = str(detail).strip() if pd.notna(detail) else consol
        grupo_desc[consol.strip()] = top_desc
        continue

    # Linha de detalhe
    if pd.notna(detail) and pd.notna(desc):
        cod = str(detail).strip()
        # Verificar se e um codigo numerico de conta
        cod_limpo = cod.lstrip("*").strip()
        if cod_limpo.isdigit():
            conta_desc[cod] = str(desc).strip()
            conta_grupo[cod] = current_group

            # Extrair valores mensais
            for mes_nome, mes_num in MESES_ORC.items():
                col_idx = list(MESES_ORC.keys()).index(mes_nome) + 5  # JAN começa na coluna 5
                val = row[_orc_cols[col_idx]] if col_idx < len(_orc_cols) else 0
                if pd.notna(val) and isinstance(val, (int, float)):
                    orc_values.append({
                        "CodDetalhe": cod,
                        "MesNum": mes_num,
                        "ValorOrc": abs(val),  # Orcamento em valores absolutos
                    })

df_orcamento = pd.DataFrame(orc_values)

# Adicionar grupos manuais que nao estao no orcamento
grupo_desc.setdefault("3.1.01.01.05", "RECEITAS DIVERSAS")

# Adicionar descricao do grupo ao razao
# Estrategia: primeiro tenta pelo codigo detalhe (via orcamento), depois pelo prefixo
def resolve_grupo_desc(row):
    """Encontra descricao do grupo pela conta detalhe ou prefixo."""
    cod_detalhe = row["CodDetalhe"]
    conta = str(row[COL_CONTA])

    # 1. Tentar via codigo detalhe -> grupo do orcamento -> descricao
    if cod_detalhe in conta_grupo:
        grp = conta_grupo[cod_detalhe]
        if grp in grupo_desc:
            return grupo_desc[grp]

    # 2. Tentar pelo prefixo da conta (varios niveis)
    parts = conta.split(".")
    for n in range(min(5, len(parts)), 1, -1):
        prefix = ".".join(parts[:n])
        if prefix in grupo_desc:
            return grupo_desc[prefix]

    return ".".join(parts[:5]) if len(parts) >= 5 else conta

df_razao["DescGrupo"] = df_razao.apply(resolve_grupo_desc, axis=1)
df_razao["DescConta"] = df_razao["CodDetalhe"].map(conta_desc).fillna("Outros")

# Separar receitas e despesas
df_receitas = df_razao[df_razao["Tipo"] == "Receita"].copy()
df_despesas = df_razao[df_razao["Tipo"] == "Despesa"].copy()

# Orcamento separado
if not df_orcamento.empty:
    df_orcamento["Tipo"] = df_orcamento["CodDetalhe"].apply(
        lambda c: "Receita" if c in conta_grupo and conta_grupo[c].startswith("3") else "Despesa"
    )
    df_orc_rec = df_orcamento[df_orcamento["Tipo"] == "Receita"].copy()
    df_orc_desp = df_orcamento[df_orcamento["Tipo"] == "Despesa"].copy()
else:
    df_orc_rec = pd.DataFrame(columns=["CodDetalhe", "MesNum", "ValorOrc"])
    df_orc_desp = pd.DataFrame(columns=["CodDetalhe", "MesNum", "ValorOrc"])

# Meses disponiveis no razao
MESES_DISP = sorted(df_razao["MesNum"].dropna().unique().astype(int).tolist())

# ============================================================
# PALETA DE CORES
# ============================================================
COLORS = {
    "receita": "#2E7D32",
    "receita_light": "#A5D6A7",
    "despesa": "#C62828",
    "despesa_light": "#EF9A9A",
    "resultado": "#1565C0",
    "resultado_light": "#64B5F6",
    "neutro": "#616161",
    "fundo": "#F5F5F5",
    "branco": "#FFFFFF",
    "titulo_bar": "#0D47A1",
    "alerta": "#F57F17",
    "texto": "#212121",
}

NOMES_MESES = {
    1: "Jan", 2: "Fev", 3: "Mar", 4: "Abr", 5: "Mai", 6: "Jun",
    7: "Jul", 8: "Ago", 9: "Set", 10: "Out", 11: "Nov", 12: "Dez",
}

NOMES_MESES_FULL = {
    1: "Janeiro", 2: "Fevereiro", 3: "Marco", 4: "Abril", 5: "Maio", 6: "Junho",
    7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro",
}


# ============================================================
# FUNCOES AUXILIARES
# ============================================================

def filtrar_dados(meses, grupos, centros_custo=None, projetos=None):
    """Filtra receitas e despesas conforme selecao."""
    rec = df_receitas.copy()
    desp = df_despesas.copy()

    if meses:
        rec = rec[rec["MesNum"].isin(meses)]
        desp = desp[desp["MesNum"].isin(meses)]
    if grupos:
        rec = rec[rec["GrupoConta"].isin(grupos)]
        desp = desp[desp["GrupoConta"].isin(grupos)]
    if centros_custo:
        rec = rec[rec["CentroCusto"].isin(centros_custo)]
        desp = desp[desp["CentroCusto"].isin(centros_custo)]
    if projetos:
        rec = rec[rec["Projeto"].isin(projetos)]
        desp = desp[desp["Projeto"].isin(projetos)]

    return rec, desp


def get_orcamento_filtrado(tipo, meses):
    """Retorna orcamento filtrado por tipo e meses."""
    orc = df_orc_rec if tipo == "Receita" else df_orc_desp
    if meses:
        orc = orc[orc["MesNum"].isin(meses)]
    return orc


def fmt_brl(valor):
    """Formata valor em Reais abreviado."""
    if pd.isna(valor) or valor == 0:
        return "R$ 0"
    neg = valor < 0
    v = abs(valor)
    if v >= 1_000_000:
        s = f"R$ {v/1_000_000:,.2f}M"
    elif v >= 1_000:
        s = f"R$ {v/1_000:,.1f}K"
    else:
        s = f"R$ {v:,.0f}"
    return f"-{s}" if neg else s


def fmt_brl_full(valor):
    """Formata valor completo em Reais."""
    if pd.isna(valor):
        return "R$ 0,00"
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_pct(valor):
    """Formata percentual."""
    if pd.isna(valor) or not np.isfinite(valor):
        return "0,0%"
    return f"{valor*100:,.1f}%".replace(".", ",")


# ============================================================
# COMPONENTES VISUAIS
# ============================================================

def make_kpi_card(titulo, valor, cor, formato="brl"):
    if formato == "brl":
        valor_fmt = fmt_brl(valor)
    elif formato == "pct":
        valor_fmt = fmt_pct(valor)
    else:
        valor_fmt = str(valor)

    return html.Div(
        style={
            "backgroundColor": COLORS["branco"], "borderRadius": "8px",
            "padding": "15px 20px", "boxShadow": "0 2px 4px rgba(0,0,0,0.08)",
            "border": "1px solid #E0E0E0", "flex": "1", "minWidth": "180px",
            "textAlign": "center",
        },
        children=[
            html.P(titulo, style={
                "margin": "0", "fontSize": "11px", "color": COLORS["neutro"],
                "fontFamily": "Segoe UI, sans-serif", "textTransform": "uppercase",
                "letterSpacing": "0.5px",
            }),
            html.H2(valor_fmt, style={
                "margin": "6px 0 0 0", "fontSize": "24px", "fontWeight": "bold",
                "color": cor, "fontFamily": "Segoe UI, sans-serif",
            }),
        ],
    )


def make_title_bar(texto):
    return html.Div(
        style={
            "backgroundColor": COLORS["titulo_bar"], "padding": "12px 24px",
            "borderRadius": "8px 8px 0 0", "marginBottom": "16px",
        },
        children=[
            html.H3(texto, style={
                "margin": "0", "color": COLORS["branco"], "fontSize": "18px",
                "fontFamily": "Segoe UI, sans-serif", "fontWeight": "600",
            }),
        ],
    )


def chart_layout(title="", height=350):
    return dict(
        title=dict(text=title, font=dict(size=14, color=COLORS["titulo_bar"], family="Segoe UI")),
        paper_bgcolor=COLORS["branco"], plot_bgcolor=COLORS["branco"],
        font=dict(family="Segoe UI", size=11, color=COLORS["texto"]),
        margin=dict(l=50, r=20, t=40, b=40), height=height,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, font=dict(size=10)),
        xaxis=dict(gridcolor="#EEEEEE", showgrid=True),
        yaxis=dict(gridcolor="#EEEEEE", showgrid=True),
    )


def make_card_container(children):
    return html.Div(children, style={
        "backgroundColor": COLORS["branco"], "borderRadius": "8px",
        "boxShadow": "0 2px 4px rgba(0,0,0,0.08)",
    })


def make_table_style(header_color=None):
    if header_color is None:
        header_color = COLORS["titulo_bar"]
    return dict(
        style_header={
            "backgroundColor": header_color, "color": COLORS["branco"],
            "fontWeight": "bold", "fontSize": "12px", "fontFamily": "Segoe UI",
            "textAlign": "center", "padding": "10px",
        },
        style_cell={
            "fontSize": "12px", "fontFamily": "Segoe UI", "padding": "8px 12px",
            "textAlign": "center", "border": "1px solid #E0E0E0",
        },
        style_data_conditional=[
            {"if": {"row_index": "odd"}, "backgroundColor": "#F8F9FA"},
        ],
        style_table={"borderRadius": "8px", "overflow": "hidden",
                     "boxShadow": "0 2px 4px rgba(0,0,0,0.08)"},
    )


# ============================================================
# APP DASH
# ============================================================

app = Dash(__name__, suppress_callback_exceptions=True, title="Dashboard Financeiro - FPGolfe")
server = app.server

# ============================================================
# ROTA PARA DOWNLOAD DO PPTX
# ============================================================

@server.route("/download-pptx")
def download_pptx():
    """Gera e retorna o PPTX executivo."""
    try:
        from gerar_pptx import gerar_apresentacao
        
        # Parâmetros (pode vir da query string)
        trimestre = flask_request.args.get("trimestre", "1T26")
        meses_str = flask_request.args.get("meses", "1,2,3")
        periodo_meses = [int(m) for m in meses_str.split(",") if m.strip()]
        
        # Gerar apresentação
        prs = gerar_apresentacao(
            df_razao=df_razao,
            df_orcamento=df_orcamento,
            grupo_desc=grupo_desc,
            conta_grupo=conta_grupo,
            conta_desc=conta_desc,
            trimestre=trimestre,
            periodo_meses=periodo_meses,
        )
        
        # Salvar em buffer
        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        
        filename = f"FPGolfe_Resultado_{trimestre}_{datetime.now().strftime('%Y%m%d_%H%M')}.pptx"
        return send_file(
            buf,
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            as_attachment=True,
            download_name=filename,
        )
    except Exception as e:
        import traceback
        return f"Erro ao gerar PPTX: {str(e)}<br><pre>{traceback.format_exc()}</pre>", 500

# Opcoes de filtros
meses_opts = [{"label": NOMES_MESES_FULL.get(m, str(m)), "value": m} for m in MESES_DISP]
grupos_rec = sorted(df_receitas["GrupoConta"].unique())
grupos_desp = sorted(df_despesas["GrupoConta"].unique())
todos_grupos = sorted(set(grupos_rec + grupos_desp))
grupos_opts = [{"label": f"{g} - {grupo_desc.get(g, '')[:30]}", "value": g} for g in todos_grupos]


# ============================================================
# LAYOUT
# ============================================================

app.layout = html.Div(
    style={"backgroundColor": COLORS["fundo"], "minHeight": "100vh", "fontFamily": "Segoe UI, sans-serif"},
    children=[
        # HEADER
        html.Div(
            style={
                "backgroundColor": COLORS["titulo_bar"], "padding": "16px 32px",
                "display": "flex", "alignItems": "center", "justifyContent": "space-between",
                "boxShadow": "0 2px 8px rgba(0,0,0,0.15)", "flexWrap": "wrap", "gap": "12px",
            },
            children=[
                html.Div(style={"display": "flex", "alignItems": "center", "gap": "16px"}, children=[
                    html.Img(src="data:image/png;base64,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",
                             style={"height": "48px", "borderRadius": "6px"}),
                    html.Div([
                        html.H1("Dashboard Financeiro", style={
                            "margin": "0", "color": COLORS["branco"], "fontSize": "22px", "fontWeight": "700",
                        }),
                        html.P("Federacao Paulista de Golfe - Razao x Orcamento 2026", style={
                            "margin": "2px 0 0 0", "color": COLORS["resultado_light"], "fontSize": "12px",
                        }),
                    ]),
                ]),
                html.Div(
                    style={"display": "flex", "gap": "12px", "alignItems": "center", "flexWrap": "wrap"},
                    children=[
                        html.Div([
                            html.Label("Mes", style={"color": "#90CAF5", "fontSize": "11px", "display": "block"}),
                            dcc.Dropdown(
                                id="filtro-mes",
                                options=meses_opts,
                                value=[],
                                multi=True,
                                placeholder="Todos",
                                style={"width": "180px", "fontSize": "12px"},
                            ),
                        ]),
                        html.Div([
                            html.Label("Centro de Custo", style={"color": "#90CAF5", "fontSize": "11px", "display": "block"}),
                            dcc.Dropdown(
                                id="filtro-cc",
                                options=[{"label": c, "value": c} for c in CC_DISP],
                                value=[],
                                multi=True,
                                placeholder="Todos",
                                style={"width": "180px", "fontSize": "12px"},
                            ),
                        ]),
                        html.Div([
                            html.Label("Projeto", style={"color": "#90CAF5", "fontSize": "11px", "display": "block"}),
                            dcc.Dropdown(
                                id="filtro-projeto",
                                options=[{"label": p, "value": p} for p in PROJ_DISP],
                                value=[],
                                multi=True,
                                placeholder="Todos",
                                style={"width": "250px", "fontSize": "12px"},
                            ),
                        ]),
                        html.Div([
                            html.Label("Grupo de Conta", style={"color": "#90CAF5", "fontSize": "11px", "display": "block"}),
                            dcc.Dropdown(
                                id="filtro-grupo",
                                options=grupos_opts,
                                value=[],
                                multi=True,
                                placeholder="Todos",
                                style={"width": "250px", "fontSize": "12px"},
                            ),
                        ]),
                        html.Div([
                            html.Label(" ", style={"color": "#90CAF5", "fontSize": "11px", "display": "block"}),
                            html.A(
                                "📊 Gerar Apresentação",
                                id="btn-pptx",
                                href="/download-pptx",
                                target="_blank",
                                style={
                                    "display": "inline-block",
                                    "padding": "7px 16px",
                                    "backgroundColor": "#FF6F00",
                                    "color": "white",
                                    "textDecoration": "none",
                                    "borderRadius": "4px",
                                    "fontSize": "12px",
                                    "fontWeight": "600",
                                    "cursor": "pointer",
                                    "boxShadow": "0 2px 4px rgba(255,111,0,0.3)",
                                    "whiteSpace": "nowrap",
                                },
                            ),
                        ]),
                    ],
                ),
            ],
        ),

        # TABS
        html.Div(
            style={"padding": "0 24px"},
            children=[
                dcc.Tabs(
                    id="tabs", value="tab-exec",
                    style={"marginTop": "16px"},
                    colors={"border": "#E0E0E0", "primary": COLORS["titulo_bar"], "background": "#E8EAF6"},
                    children=[
                        dcc.Tab(label="Visao Executiva", value="tab-exec",
                                style={"fontFamily": "Segoe UI", "fontWeight": "500", "padding": "10px 24px"},
                                selected_style={"fontFamily": "Segoe UI", "fontWeight": "700", "padding": "10px 24px",
                                                "borderTop": f"3px solid {COLORS['titulo_bar']}"}),
                        dcc.Tab(label="Receitas", value="tab-rec",
                                style={"fontFamily": "Segoe UI", "fontWeight": "500", "padding": "10px 24px"},
                                selected_style={"fontFamily": "Segoe UI", "fontWeight": "700", "padding": "10px 24px",
                                                "borderTop": f"3px solid {COLORS['receita']}"}),
                        dcc.Tab(label="Despesas", value="tab-desp",
                                style={"fontFamily": "Segoe UI", "fontWeight": "500", "padding": "10px 24px"},
                                selected_style={"fontFamily": "Segoe UI", "fontWeight": "700", "padding": "10px 24px",
                                                "borderTop": f"3px solid {COLORS['despesa']}"}),
                        dcc.Tab(label="Resultado / DRE", value="tab-dre",
                                style={"fontFamily": "Segoe UI", "fontWeight": "500", "padding": "10px 24px"},
                                selected_style={"fontFamily": "Segoe UI", "fontWeight": "700", "padding": "10px 24px",
                                                "borderTop": f"3px solid {COLORS['resultado']}"}),
                        dcc.Tab(label="Orcado x Realizado", value="tab-oxr",
                                style={"fontFamily": "Segoe UI", "fontWeight": "500", "padding": "10px 24px"},
                                selected_style={"fontFamily": "Segoe UI", "fontWeight": "700", "padding": "10px 24px",
                                                "borderTop": f"3px solid {COLORS['alerta']}"}),
                    ],
                ),
                html.Div(id="tab-content", style={"marginTop": "16px", "paddingBottom": "32px"}),
            ],
        ),
    ],
)


# ============================================================
# CALLBACK PRINCIPAL
# ============================================================

@callback(
    Output("tab-content", "children"),
    Input("tabs", "value"),
    Input("filtro-mes", "value"),
    Input("filtro-grupo", "value"),
    Input("filtro-cc", "value"),
    Input("filtro-projeto", "value"),
)
def render_tab(tab, meses, grupos, centros_custo, projetos):
    rec, desp = filtrar_dados(meses, grupos, centros_custo, projetos)

    if tab == "tab-exec":
        return build_page_executiva(rec, desp, meses)
    elif tab == "tab-rec":
        return build_page_receitas(rec, desp, meses)
    elif tab == "tab-desp":
        return build_page_despesas(rec, desp, meses)
    elif tab == "tab-dre":
        return build_page_dre(rec, desp, meses)
    elif tab == "tab-oxr":
        return build_page_orcado_realizado(rec, desp, meses)
    return html.Div("Selecione uma aba")


# ============================================================
# PAGINA 1 - VISAO EXECUTIVA
# ============================================================

def build_page_executiva(rec, desp, meses_filtro):
    receita_total = rec["Valor"].sum()
    despesa_total = desp["Valor"].sum()
    resultado = receita_total - despesa_total
    margem = resultado / receita_total if receita_total else 0

    # Orcamento para meses selecionados
    orc_rec = get_orcamento_filtrado("Receita", meses_filtro)
    orc_desp = get_orcamento_filtrado("Despesa", meses_filtro)
    receita_orcada = orc_rec["ValorOrc"].sum() if not orc_rec.empty else 0
    despesa_orcada = orc_desp["ValorOrc"].sum() if not orc_desp.empty else 0
    resultado_orcado = receita_orcada - despesa_orcada

    # KPIs
    cards = html.Div(
        style={"display": "flex", "gap": "12px", "flexWrap": "wrap"},
        children=[
            make_kpi_card("Receita Realizada", receita_total, COLORS["receita"]),
            make_kpi_card("Despesa Realizada", despesa_total, COLORS["despesa"]),
            make_kpi_card("Resultado", resultado,
                         COLORS["receita"] if resultado >= 0 else COLORS["despesa"]),
            make_kpi_card("Margem", margem, COLORS["resultado"], formato="pct"),
            make_kpi_card("Resultado Orcado", resultado_orcado, COLORS["neutro"]),
        ],
    )

    # --- Evolucao Mensal ---
    rec_mensal = rec.groupby("AnoMes")["Valor"].sum().reset_index().rename(columns={"Valor": "Receita"})
    desp_mensal = desp.groupby("AnoMes")["Valor"].sum().reset_index().rename(columns={"Valor": "Despesa"})
    evol = rec_mensal.merge(desp_mensal, on="AnoMes", how="outer").fillna(0).sort_values("AnoMes")
    evol["Resultado"] = evol["Receita"] - evol["Despesa"]

    fig_evol = go.Figure()
    fig_evol.add_trace(go.Scatter(x=evol["AnoMes"], y=evol["Receita"], name="Receita",
                                   line=dict(color=COLORS["receita"], width=2.5), mode="lines+markers",
                                   marker=dict(size=6)))
    fig_evol.add_trace(go.Scatter(x=evol["AnoMes"], y=evol["Despesa"], name="Despesa",
                                   line=dict(color=COLORS["despesa"], width=2.5), mode="lines+markers",
                                   marker=dict(size=6)))
    fig_evol.add_trace(go.Scatter(x=evol["AnoMes"], y=evol["Resultado"], name="Resultado",
                                   line=dict(color=COLORS["resultado"], width=2, dash="dash"), mode="lines+markers",
                                   marker=dict(size=4)))
    fig_evol.update_layout(**chart_layout("Evolucao Mensal - Receita x Despesa x Resultado", 340))

    # --- Real x Orcado por Mes ---
    fig_orcado = go.Figure()
    meses_chart = sorted(rec["MesNum"].dropna().unique().astype(int).tolist())
    labels_mes = [NOMES_MESES.get(m, str(m)) for m in meses_chart]

    rec_real_mes = rec.groupby("MesNum")["Valor"].sum()
    desp_real_mes = desp.groupby("MesNum")["Valor"].sum()
    orc_rec_mes = orc_rec.groupby("MesNum")["ValorOrc"].sum() if not orc_rec.empty else pd.Series(dtype=float)
    orc_desp_mes = orc_desp.groupby("MesNum")["ValorOrc"].sum() if not orc_desp.empty else pd.Series(dtype=float)

    fig_orcado.add_trace(go.Bar(name="Receita Real", x=labels_mes,
                                y=[rec_real_mes.get(m, 0) for m in meses_chart],
                                marker_color=COLORS["receita"]))
    fig_orcado.add_trace(go.Bar(name="Receita Orcada", x=labels_mes,
                                y=[orc_rec_mes.get(m, 0) for m in meses_chart],
                                marker_color=COLORS["receita_light"]))
    fig_orcado.add_trace(go.Bar(name="Despesa Real", x=labels_mes,
                                y=[desp_real_mes.get(m, 0) for m in meses_chart],
                                marker_color=COLORS["despesa"]))
    fig_orcado.add_trace(go.Bar(name="Despesa Orcada", x=labels_mes,
                                y=[orc_desp_mes.get(m, 0) for m in meses_chart],
                                marker_color=COLORS["despesa_light"]))
    fig_orcado.update_layout(**chart_layout("Realizado x Orcado por Mes", 340))
    fig_orcado.update_layout(barmode="group")

    # --- Tabela resumo por grupo ---
    rec_grp = rec.groupby("DescGrupo")["Valor"].sum().reset_index().rename(columns={"Valor": "Receita"})
    desp_grp = desp.groupby("DescGrupo")["Valor"].sum().reset_index().rename(columns={"Valor": "Despesa"})

    # Montar tabela: receitas primeiro, depois despesas
    rows_resumo = []
    for _, r in rec_grp.sort_values("Receita", ascending=False).iterrows():
        rows_resumo.append({"Grupo": r["DescGrupo"], "Tipo": "Receita", "Valor": r["Receita"]})
    for _, r in desp_grp.sort_values("Despesa", ascending=False).iterrows():
        rows_resumo.append({"Grupo": r["DescGrupo"], "Tipo": "Despesa", "Valor": r["Despesa"]})

    df_resumo = pd.DataFrame(rows_resumo)
    df_resumo["Valor_fmt"] = df_resumo["Valor"].apply(fmt_brl_full)

    tabela_resumo = dash_table.DataTable(
        id="tabela-resumo-exec",
        data=df_resumo[["Grupo", "Tipo", "Valor_fmt"]].rename(columns={"Valor_fmt": "Valor"}).to_dict("records"),
        columns=[
            {"name": "Grupo", "id": "Grupo"},
            {"name": "Tipo", "id": "Tipo"},
            {"name": "Valor", "id": "Valor"},
        ],
        **make_table_style(),
    )

    # --- Barras Horizontais: Composição das Receitas (visão executiva) ---
    rec_sorted = rec_grp.sort_values("Receita", ascending=True)
    total_rec = rec_sorted["Receita"].sum()
    percs = (rec_sorted["Receita"] / total_rec * 100) if total_rec > 0 else rec_sorted["Receita"] * 0
    
    # Texto à direita de cada barra: valor + percentual
    textos = [f"R$ {v/1e6:,.2f} Mi  ·  {p:.1f}%".replace(",", "X").replace(".", ",").replace("X", ".") 
              for v, p in zip(rec_sorted["Receita"], percs)]
    
    # Gradiente de verde (barras menores mais claras, maior mais escura)
    n = len(rec_sorted)
    cores = ["#A5D6A7", "#81C784", "#66BB6A", "#43A047", "#2E7D32", "#1B5E20"][-n:] if n <= 6 else ["#2E7D32"] * n
    
    fig_donut = go.Figure(go.Bar(
        y=rec_sorted["DescGrupo"],
        x=rec_sorted["Receita"],
        orientation="h",
        marker=dict(color=cores, line=dict(width=0)),
        text=textos,
        textposition="outside",
        textfont=dict(size=11, family="Segoe UI", color="#333"),
        hovertemplate="<b>%{y}</b><br>R$ %{x:,.2f}<extra></extra>",
        cliponaxis=False,
    ))
    
    fig_donut.update_layout(
        title=dict(text="<b>Composicao das Receitas</b>", font=dict(size=14, color=COLORS["titulo_bar"], family="Segoe UI"), x=0.02),
        height=320,
        margin=dict(l=10, r=140, t=50, b=20),
        paper_bgcolor="white",
        plot_bgcolor="white",
        showlegend=False,
        xaxis=dict(
            showgrid=False, showticklabels=False, zeroline=False, visible=False,
            range=[0, rec_sorted["Receita"].max() * 1.45],
        ),
        yaxis=dict(
            showgrid=False, zeroline=False,
            tickfont=dict(size=11, family="Segoe UI", color="#555"),
            automargin=True,
        ),
        bargap=0.35,
    )

    return html.Div(id="page-executiva", children=[
        make_title_bar("Visao Executiva - Dashboard Financeiro 2026"),
        cards,
        html.Div(
            style={"display": "grid", "gridTemplateColumns": "1fr 1fr", "gap": "16px", "marginTop": "16px"},
            children=[
                make_card_container(dcc.Graph(id="graph-exec-evol", figure=fig_evol, config={"displayModeBar": False})),
                make_card_container(dcc.Graph(id="graph-exec-orcado", figure=fig_orcado, config={"displayModeBar": False})),
            ],
        ),
        html.Div(
            style={"display": "grid", "gridTemplateColumns": "1.2fr 0.8fr", "gap": "16px", "marginTop": "16px"},
            children=[
                html.Div([
                    html.H4("Resumo por Grupo de Conta", style={
                        "margin": "0 0 12px 0", "color": COLORS["titulo_bar"], "fontSize": "14px",
                        "fontFamily": "Segoe UI",
                    }),
                    tabela_resumo,
                ], style={"backgroundColor": COLORS["branco"], "borderRadius": "8px", "padding": "16px",
                          "boxShadow": "0 2px 4px rgba(0,0,0,0.08)"}),
                make_card_container(dcc.Graph(id="graph-exec-donut", figure=fig_donut, config={"displayModeBar": False})),
            ],
        ),
    ])


# ============================================================
# PAGINA 2 - RECEITAS
# ============================================================

def build_page_receitas(rec, desp, meses_filtro):
    receita_total = rec["Valor"].sum()
    orc_rec = get_orcamento_filtrado("Receita", meses_filtro)
    receita_orcada = orc_rec["ValorOrc"].sum() if not orc_rec.empty else 0
    var_rec = (receita_total - receita_orcada) / receita_orcada if receita_orcada else 0

    cards = html.Div(
        style={"display": "flex", "gap": "12px", "flexWrap": "wrap"},
        children=[
            make_kpi_card("Receita Realizada", receita_total, COLORS["receita"]),
            make_kpi_card("Receita Orcada", receita_orcada, "#66BB6A"),
            make_kpi_card("Var. Real x Orcado", var_rec,
                         COLORS["receita"] if var_rec >= 0 else COLORS["despesa"], formato="pct"),
            make_kpi_card("Qtd Lancamentos", len(rec), COLORS["neutro"], formato="int"),
        ],
    )

    # --- Receita por Grupo (horizontal bar) ---
    rec_grp = rec.groupby("DescGrupo")["Valor"].sum().reset_index().sort_values("Valor", ascending=True)
    fig_grp = go.Figure(go.Bar(
        x=rec_grp["Valor"], y=rec_grp["DescGrupo"], orientation="h",
        marker_color=COLORS["receita"], text=rec_grp["Valor"].apply(fmt_brl),
        textposition="outside", textfont=dict(size=11),
    ))
    fig_grp.update_layout(**chart_layout("Receita por Grupo", 300))

    # --- Evolucao mensal por grupo ---
    rec_grp_mes = rec.groupby(["AnoMes", "DescGrupo"])["Valor"].sum().reset_index()
    grupos = rec_grp_mes["DescGrupo"].unique()
    cores = ["#1B5E20", "#2E7D32", "#43A047", "#66BB6A", "#A5D6A7"]

    fig_evol = go.Figure()
    for i, grp in enumerate(sorted(grupos)):
        dados = rec_grp_mes[rec_grp_mes["DescGrupo"] == grp].sort_values("AnoMes")
        fig_evol.add_trace(go.Bar(
            x=dados["AnoMes"], y=dados["Valor"], name=grp[:25],
            marker_color=cores[i % len(cores)],
        ))
    fig_evol.update_layout(**chart_layout("Receita Mensal por Grupo", 300))
    fig_evol.update_layout(barmode="stack")

    # --- Heatmap Conta x Mes ---
    rec_heat = rec.copy()
    rec_heat["NomeMes"] = rec_heat["MesNum"].map(NOMES_MESES)
    pivot = rec_heat.groupby(["DescGrupo", "NomeMes", "MesNum"])["Valor"].sum().reset_index()
    pivot = pivot.sort_values("MesNum")
    pivot_table = pivot.pivot_table(index="DescGrupo", columns="NomeMes", values="Valor",
                                     aggfunc="sum", sort=False).fillna(0)
    mes_order = [NOMES_MESES[i] for i in range(1, 13) if NOMES_MESES[i] in pivot_table.columns]
    if mes_order:
        pivot_table = pivot_table[mes_order]

    fig_heat = go.Figure(go.Heatmap(
        z=pivot_table.values, x=pivot_table.columns.tolist(), y=pivot_table.index.tolist(),
        colorscale=[[0, "#FFFFFF"], [1, COLORS["receita"]]],
        text=[[fmt_brl(v) for v in row] for row in pivot_table.values],
        texttemplate="%{text}", textfont=dict(size=10), showscale=False,
    ))
    fig_heat.update_layout(**chart_layout("Receita por Grupo x Mes", 260))
    fig_heat.update_layout(margin=dict(l=200))

    # --- Top contas ---
    top_contas = rec.groupby("DescConta")["Valor"].sum().reset_index().sort_values("Valor", ascending=False).head(10)
    top_contas["Valor_fmt"] = top_contas["Valor"].apply(fmt_brl_full)
    top_contas["Pct"] = (top_contas["Valor"] / receita_total).apply(fmt_pct) if receita_total else "0%"

    tabela_top = dash_table.DataTable(
        id="tabela-top-rec",
        data=top_contas[["DescConta", "Valor_fmt", "Pct"]].rename(
            columns={"Valor_fmt": "Valor", "DescConta": "Conta"}).to_dict("records"),
        columns=[
            {"name": "Conta", "id": "Conta"},
            {"name": "Valor", "id": "Valor"},
            {"name": "% do Total", "id": "Pct"},
        ],
        **make_table_style(COLORS["receita"]),
    )

    return html.Div(id="page-receitas", children=[
        make_title_bar("Analise de Receitas"),
        cards,
        html.Div(
            style={"display": "grid", "gridTemplateColumns": "1fr 1fr", "gap": "16px", "marginTop": "16px"},
            children=[
                make_card_container(dcc.Graph(id="graph-rec-grp", figure=fig_grp, config={"displayModeBar": False})),
                make_card_container(dcc.Graph(id="graph-rec-evol", figure=fig_evol, config={"displayModeBar": False})),
            ],
        ),
        html.Div(
            style={"marginTop": "16px"},
            children=[
                make_card_container(dcc.Graph(id="graph-rec-heat", figure=fig_heat, config={"displayModeBar": False})),
            ],
        ),
        html.Div(
            style={"marginTop": "16px", "backgroundColor": COLORS["branco"], "borderRadius": "8px",
                   "padding": "16px", "boxShadow": "0 2px 4px rgba(0,0,0,0.08)"},
            children=[
                html.H4("Top 10 Contas de Receita", style={
                    "margin": "0 0 12px 0", "color": COLORS["receita"], "fontSize": "14px"}),
                tabela_top,
            ],
        ),
    ])


# ============================================================
# PAGINA 3 - DESPESAS
# ============================================================

def build_page_despesas(rec, desp, meses_filtro):
    despesa_total = desp["Valor"].sum()
    orc_desp = get_orcamento_filtrado("Despesa", meses_filtro)
    despesa_orcada = orc_desp["ValorOrc"].sum() if not orc_desp.empty else 0
    var_desp = (despesa_total - despesa_orcada) / despesa_orcada if despesa_orcada else 0

    # Classificar Meio x Fim com base no prefixo da conta
    # 4.1.01.02 e 4.1.01.04/05 = Fim (esporte), resto = Meio (administrativo)
    desp_copy = desp.copy()
    desp_copy["Atividade"] = desp_copy["GrupoConta"].apply(
        lambda g: "Atividade Fim" if any(g.startswith(p) for p in ["4.1.01.02", "4.1.01.04", "4.1.01.05"]) else "Atividade Meio"
    )
    meio = desp_copy[desp_copy["Atividade"] == "Atividade Meio"]["Valor"].sum()
    fim = desp_copy[desp_copy["Atividade"] == "Atividade Fim"]["Valor"].sum()
    pct_meio = meio / despesa_total if despesa_total else 0

    cards = html.Div(
        style={"display": "flex", "gap": "12px", "flexWrap": "wrap"},
        children=[
            make_kpi_card("Despesa Realizada", despesa_total, COLORS["despesa"]),
            make_kpi_card("Despesa Orcada", despesa_orcada, "#EF5350"),
            make_kpi_card("Var. Real x Orcado", var_desp,
                         COLORS["despesa"] if var_desp > 0 else COLORS["receita"], formato="pct"),
            make_kpi_card("% Ativ. Meio", pct_meio, COLORS["neutro"], formato="pct"),
        ],
    )

    # --- Despesas por Grupo (Meio x Fim) ---
    desp_cat = desp_copy.groupby(["DescGrupo", "Atividade"])["Valor"].sum().reset_index()
    cats_ordem = desp_copy.groupby("DescGrupo")["Valor"].sum().sort_values(ascending=True).index.tolist()

    fig_cat = go.Figure()
    for ativ, cor in [("Atividade Meio", COLORS["despesa"]), ("Atividade Fim", COLORS["despesa_light"])]:
        dados = desp_cat[desp_cat["Atividade"] == ativ]
        dados = dados.set_index("DescGrupo").reindex(cats_ordem).fillna(0).reset_index()
        fig_cat.add_trace(go.Bar(
            y=dados["DescGrupo"], x=dados["Valor"], name=ativ,
            orientation="h", marker_color=cor,
        ))
    fig_cat.update_layout(**chart_layout("Despesas por Grupo (Meio x Fim)", 380))
    fig_cat.update_layout(barmode="stack", margin=dict(l=200))

    # --- Evolucao mensal ---
    desp_mensal = desp.groupby("AnoMes")["Valor"].sum().reset_index().sort_values("AnoMes")
    fig_evol = go.Figure()
    fig_evol.add_trace(go.Scatter(
        x=desp_mensal["AnoMes"], y=desp_mensal["Valor"], name="Despesa Total",
        line=dict(color=COLORS["despesa"], width=2.5), mode="lines+markers", marker=dict(size=6),
        fill="tozeroy", fillcolor="rgba(198,40,40,0.1)",
    ))
    fig_evol.update_layout(**chart_layout("Evolucao Mensal das Despesas", 340))

    # --- Treemap ---
    desp_tree = desp.groupby("DescGrupo")["Valor"].sum().reset_index().sort_values("Valor", ascending=False)
    fig_treemap = px.treemap(
        desp_tree, path=["DescGrupo"], values="Valor",
        color="Valor", color_continuous_scale=["#FFCDD2", "#B71C1C"],
    )
    fig_treemap.update_layout(**chart_layout("Distribuicao das Despesas", 340))
    fig_treemap.update_layout(coloraxis_showscale=False, margin=dict(l=10, r=10, t=40, b=10))
    fig_treemap.update_traces(textinfo="label+value+percent root", textfont=dict(size=11))

    # --- Heatmap Despesas Grupo x Mes ---
    desp_heat = desp.copy()
    desp_heat["NomeMes"] = desp_heat["MesNum"].map(NOMES_MESES)
    pivot = desp_heat.groupby(["DescGrupo", "NomeMes", "MesNum"])["Valor"].sum().reset_index()
    pivot = pivot.sort_values("MesNum")
    pivot_table = pivot.pivot_table(index="DescGrupo", columns="NomeMes", values="Valor",
                                     aggfunc="sum", sort=False).fillna(0)
    mes_order = [NOMES_MESES[i] for i in range(1, 13) if NOMES_MESES[i] in pivot_table.columns]
    if mes_order:
        pivot_table = pivot_table[mes_order]

    fig_heat = go.Figure(go.Heatmap(
        z=pivot_table.values, x=pivot_table.columns.tolist(), y=pivot_table.index.tolist(),
        colorscale=[[0, "#FFFFFF"], [1, COLORS["despesa"]]],
        text=[[fmt_brl(v) for v in row] for row in pivot_table.values],
        texttemplate="%{text}", textfont=dict(size=10), showscale=False,
    ))
    fig_heat.update_layout(**chart_layout("Despesas por Grupo x Mes", 320))
    fig_heat.update_layout(margin=dict(l=200))

    # --- Top contas ---
    top_contas = desp.groupby("DescConta")["Valor"].sum().reset_index().sort_values("Valor", ascending=False).head(10)
    top_contas["Valor_fmt"] = top_contas["Valor"].apply(fmt_brl_full)
    top_contas["Pct"] = (top_contas["Valor"] / despesa_total).apply(fmt_pct) if despesa_total else "0%"

    tabela_top = dash_table.DataTable(
        id="tabela-top-desp",
        data=top_contas[["DescConta", "Valor_fmt", "Pct"]].rename(
            columns={"Valor_fmt": "Valor", "DescConta": "Conta"}).to_dict("records"),
        columns=[
            {"name": "Conta", "id": "Conta"},
            {"name": "Valor", "id": "Valor"},
            {"name": "% do Total", "id": "Pct"},
        ],
        **make_table_style(COLORS["despesa"]),
    )

    return html.Div(id="page-despesas", children=[
        make_title_bar("Analise de Despesas"),
        cards,
        html.Div(
            style={"display": "grid", "gridTemplateColumns": "1fr 1fr", "gap": "16px", "marginTop": "16px"},
            children=[
                make_card_container(dcc.Graph(id="graph-desp-cat", figure=fig_cat, config={"displayModeBar": False})),
                make_card_container(dcc.Graph(id="graph-desp-evol", figure=fig_evol, config={"displayModeBar": False})),
            ],
        ),
        html.Div(
            style={"display": "grid", "gridTemplateColumns": "1fr 1fr", "gap": "16px", "marginTop": "16px"},
            children=[
                make_card_container(dcc.Graph(id="graph-desp-tree", figure=fig_treemap, config={"displayModeBar": False})),
                make_card_container(dcc.Graph(id="graph-desp-heat", figure=fig_heat, config={"displayModeBar": False})),
            ],
        ),
        html.Div(
            style={"marginTop": "16px", "backgroundColor": COLORS["branco"], "borderRadius": "8px",
                   "padding": "16px", "boxShadow": "0 2px 4px rgba(0,0,0,0.08)"},
            children=[
                html.H4("Top 10 Contas de Despesa", style={
                    "margin": "0 0 12px 0", "color": COLORS["despesa"], "fontSize": "14px"}),
                tabela_top,
            ],
        ),
    ])


# ============================================================
# PAGINA 4 - RESULTADO / DRE
# ============================================================

def build_page_dre(rec, desp, meses_filtro):
    receita_total = rec["Valor"].sum()
    despesa_total = desp["Valor"].sum()
    resultado = receita_total - despesa_total
    margem = resultado / receita_total if receita_total else 0

    # Orcamento
    orc_rec = get_orcamento_filtrado("Receita", meses_filtro)
    orc_desp = get_orcamento_filtrado("Despesa", meses_filtro)
    receita_orcada = orc_rec["ValorOrc"].sum() if not orc_rec.empty else 0
    despesa_orcada = orc_desp["ValorOrc"].sum() if not orc_desp.empty else 0
    resultado_orcado = receita_orcada - despesa_orcada

    cards = html.Div(
        style={"display": "flex", "gap": "12px", "flexWrap": "wrap"},
        children=[
            make_kpi_card("Resultado Real", resultado,
                         COLORS["receita"] if resultado >= 0 else COLORS["despesa"]),
            make_kpi_card("Margem", margem, COLORS["resultado"], formato="pct"),
            make_kpi_card("Resultado Orcado", resultado_orcado, COLORS["neutro"]),
        ],
    )

    # --- DRE Estruturada ---
    # Receitas por grupo
    rec_by_grp = rec.groupby("DescGrupo")["Valor"].sum().sort_values(ascending=False)
    # Despesas por grupo
    desp_copy = desp.copy()
    desp_copy["Atividade"] = desp_copy["GrupoConta"].apply(
        lambda g: "Fim" if any(g.startswith(p) for p in ["4.1.01.02", "4.1.01.04", "4.1.01.05"]) else "Meio"
    )
    desp_meio = desp_copy[desp_copy["Atividade"] == "Meio"]
    desp_fim = desp_copy[desp_copy["Atividade"] == "Fim"]
    meio_total = desp_meio["Valor"].sum()
    fim_total = desp_fim["Valor"].sum()
    desp_meio_grp = desp_meio.groupby("DescGrupo")["Valor"].sum().sort_values(ascending=False)
    desp_fim_grp = desp_fim.groupby("DescGrupo")["Valor"].sum().sort_values(ascending=False)

    dre_rows = []

    # RECEITAS
    dre_rows.append({"Conta": "RECEITA BRUTA", "Real": receita_total, "Orcado": receita_orcada, "nivel": 1})
    for grp, val in rec_by_grp.items():
        dre_rows.append({"Conta": f"    {grp}", "Real": val, "Orcado": 0, "nivel": 2})

    # DESPESAS ATIVIDADE FIM
    dre_rows.append({"Conta": "(-) DESPESAS ATIVIDADE FIM", "Real": -fim_total, "Orcado": 0, "nivel": 1})
    for grp, val in desp_fim_grp.items():
        dre_rows.append({"Conta": f"    {grp}", "Real": -val, "Orcado": 0, "nivel": 2})

    # MARGEM BRUTA
    margem_bruta = receita_total - fim_total
    dre_rows.append({"Conta": "= MARGEM BRUTA", "Real": margem_bruta, "Orcado": 0, "nivel": 1})

    # DESPESAS ATIVIDADE MEIO
    dre_rows.append({"Conta": "(-) DESPESAS ATIVIDADE MEIO", "Real": -meio_total, "Orcado": 0, "nivel": 1})
    for grp, val in desp_meio_grp.items():
        dre_rows.append({"Conta": f"    {grp}", "Real": -val, "Orcado": 0, "nivel": 2})

    # RESULTADO
    dre_rows.append({"Conta": "= RESULTADO OPERACIONAL", "Real": resultado, "Orcado": resultado_orcado, "nivel": 1})
    dre_rows.append({"Conta": "    Margem Operacional %", "Real": margem, "Orcado": 0, "nivel": 2, "is_pct": True})

    dre_df = pd.DataFrame(dre_rows)
    dre_display = dre_df.copy()
    dre_display["Real_fmt"] = dre_display.apply(
        lambda r: fmt_pct(r["Real"]) if r.get("is_pct") else fmt_brl_full(r["Real"]), axis=1
    )
    dre_display["Orcado_fmt"] = dre_display.apply(
        lambda r: fmt_pct(r["Orcado"]) if r.get("is_pct") else (fmt_brl_full(r["Orcado"]) if r["Orcado"] != 0 else "-"), axis=1
    )

    nivel1_indices = [i for i, r in enumerate(dre_rows) if r["nivel"] == 1]

    tabela_dre = dash_table.DataTable(
        id="tabela-dre-resultado",
        data=dre_display[["Conta", "Real_fmt", "Orcado_fmt"]].rename(
            columns={"Real_fmt": "Realizado", "Orcado_fmt": "Orcado"}).to_dict("records"),
        columns=[
            {"name": "Conta DRE", "id": "Conta"},
            {"name": "Realizado", "id": "Realizado"},
            {"name": "Orcado", "id": "Orcado"},
        ],
        style_header={
            "backgroundColor": COLORS["titulo_bar"], "color": COLORS["branco"],
            "fontWeight": "bold", "fontSize": "13px", "fontFamily": "Segoe UI",
            "textAlign": "center", "padding": "12px",
        },
        style_cell={
            "fontSize": "12px", "fontFamily": "Segoe UI Semibold, Segoe UI", "padding": "8px 16px",
            "border": "1px solid #E0E0E0",
        },
        style_cell_conditional=[
            {"if": {"column_id": "Conta"}, "textAlign": "left", "width": "50%"},
            {"if": {"column_id": "Realizado"}, "textAlign": "right", "width": "25%"},
            {"if": {"column_id": "Orcado"}, "textAlign": "right", "width": "25%"},
        ],
        style_data_conditional=[
            {"if": {"row_index": nivel1_indices},
             "fontWeight": "bold", "backgroundColor": "#E3F2FD", "fontSize": "13px"},
            {"if": {"row_index": len(dre_rows) - 2, "column_id": "Realizado"},
             "color": COLORS["receita"] if resultado >= 0 else COLORS["despesa"],
             "fontWeight": "bold", "fontSize": "14px"},
        ],
        style_table={"borderRadius": "8px", "overflow": "hidden", "boxShadow": "0 2px 4px rgba(0,0,0,0.08)"},
    )

    # --- Margem mensal ---
    rec_mensal = rec.groupby("AnoMes")["Valor"].sum().reset_index().sort_values("AnoMes")
    desp_mensal = desp.groupby("AnoMes")["Valor"].sum().reset_index().sort_values("AnoMes")
    res_mensal = rec_mensal.merge(desp_mensal, on="AnoMes", how="outer", suffixes=("_rec", "_desp")).fillna(0)
    res_mensal["Resultado"] = res_mensal["Valor_rec"] - res_mensal["Valor_desp"]
    res_mensal["Margem"] = (res_mensal["Resultado"] / res_mensal["Valor_rec"]).fillna(0)

    fig_margem = go.Figure()
    fig_margem.add_trace(go.Scatter(
        x=res_mensal["AnoMes"], y=res_mensal["Margem"], name="Margem %",
        line=dict(color=COLORS["resultado"], width=2.5),
        fill="tozeroy", fillcolor="rgba(21,101,192,0.1)", mode="lines+markers", marker=dict(size=6),
    ))
    fig_margem.add_hline(y=0, line_dash="dash", line_color=COLORS["neutro"], line_width=1)
    fig_margem.update_layout(**chart_layout("Margem Operacional Mensal", 300))
    fig_margem.update_yaxes(tickformat=".0%")

    # --- Waterfall ---
    fig_waterfall = go.Figure(go.Waterfall(
        name="Resultado", orientation="v",
        measure=["absolute", "relative", "relative", "total"],
        x=["Receita Bruta", "(-) Desp. Fim", "(-) Desp. Meio", "Resultado"],
        y=[receita_total, -fim_total, -meio_total, 0],
        text=[fmt_brl(receita_total), fmt_brl(-fim_total), fmt_brl(-meio_total), fmt_brl(resultado)],
        textposition="outside", textfont=dict(size=11),
        connector=dict(line=dict(color="#BDBDBD", width=1)),
        increasing=dict(marker=dict(color=COLORS["receita"])),
        decreasing=dict(marker=dict(color=COLORS["despesa"])),
        totals=dict(marker=dict(color=COLORS["resultado"])),
    ))
    fig_waterfall.update_layout(**chart_layout("Composicao do Resultado", 300))

    # --- Real vs Orcado barras ---
    fig_comp = go.Figure()
    fig_comp.add_trace(go.Bar(
        name="Receita Real", x=["Receita", "Despesa", "Resultado"],
        y=[receita_total, despesa_total, resultado],
        marker_color=[COLORS["receita"], COLORS["despesa"], COLORS["resultado"]],
        text=[fmt_brl(receita_total), fmt_brl(despesa_total), fmt_brl(resultado)],
        textposition="outside", textfont=dict(size=11),
    ))
    fig_comp.add_trace(go.Bar(
        name="Orcado", x=["Receita", "Despesa", "Resultado"],
        y=[receita_orcada, despesa_orcada, resultado_orcado],
        marker_color=[COLORS["receita_light"], COLORS["despesa_light"], COLORS["resultado_light"]],
        text=[fmt_brl(receita_orcada), fmt_brl(despesa_orcada), fmt_brl(resultado_orcado)],
        textposition="outside", textfont=dict(size=11),
    ))
    fig_comp.update_layout(**chart_layout("Real x Orcado - Visao Geral", 300))
    fig_comp.update_layout(barmode="group")

    return html.Div(id="page-dre", children=[
        make_title_bar("Demonstrativo de Resultado (DRE) - 1o Trimestre 2026"),
        cards,
        html.Div(
            style={"display": "grid", "gridTemplateColumns": "1.2fr 0.8fr", "gap": "16px", "marginTop": "16px"},
            children=[
                html.Div([
                    html.H4("DRE - Realizado x Orcado", style={
                        "margin": "0 0 12px 0", "color": COLORS["titulo_bar"], "fontSize": "14px"}),
                    tabela_dre,
                ], style={"backgroundColor": COLORS["branco"], "borderRadius": "8px", "padding": "16px",
                          "boxShadow": "0 2px 4px rgba(0,0,0,0.08)"}),
                html.Div([
                    dcc.Graph(id="graph-dre-margem", figure=fig_margem, config={"displayModeBar": False}),
                ], style={"backgroundColor": COLORS["branco"], "borderRadius": "8px",
                          "boxShadow": "0 2px 4px rgba(0,0,0,0.08)"}),
            ],
        ),
        html.Div(
            style={"display": "grid", "gridTemplateColumns": "1fr 1fr", "gap": "16px", "marginTop": "16px"},
            children=[
                make_card_container(dcc.Graph(id="graph-dre-waterfall", figure=fig_waterfall, config={"displayModeBar": False})),
                make_card_container(dcc.Graph(id="graph-dre-comp", figure=fig_comp, config={"displayModeBar": False})),
            ],
        ),
    ])


# ============================================================
# PAGINA 5 - ORCADO x REALIZADO (RELATORIO ANALITICO)
# ============================================================

def build_page_orcado_realizado(rec, desp, meses_filtro):
    """Relatorio analitico Orcado x Realizado com opcao de impressao PDF."""

    meses_ativos = sorted(meses_filtro) if meses_filtro else MESES_DISP

    # ---- MONTAR DADOS POR CONTA DETALHE E MES ----
    # Realizado por conta detalhe e mes
    all_data = pd.concat([rec.assign(Natureza="Receita"), desp.assign(Natureza="Despesa")])

    real_pivot = all_data.groupby(["Natureza", "DescGrupo", "DescConta", "CodDetalhe", "MesNum"])["Valor"].sum().reset_index()

    # Orcamento por conta detalhe e mes
    orc_all = pd.concat([
        df_orc_rec.assign(Natureza="Receita") if not df_orc_rec.empty else pd.DataFrame(),
        df_orc_desp.assign(Natureza="Despesa") if not df_orc_desp.empty else pd.DataFrame(),
    ])
    if not orc_all.empty and meses_filtro:
        orc_all = orc_all[orc_all["MesNum"].isin(meses_filtro)]
    elif not orc_all.empty:
        orc_all = orc_all[orc_all["MesNum"].isin(MESES_DISP)]

    # Mapear descricoes no orcamento
    if not orc_all.empty:
        orc_all["DescConta"] = orc_all["CodDetalhe"].map(conta_desc).fillna("Outros")
        orc_all["DescGrupo"] = orc_all["CodDetalhe"].apply(
            lambda c: grupo_desc.get(conta_grupo.get(c, ""), c) if c in conta_grupo else "Outros"
        )

    # ---- TABELA RESUMO POR GRUPO ----
    rows_grupo = []

    for natureza in ["Receita", "Despesa"]:
        nat_real = real_pivot[real_pivot["Natureza"] == natureza]
        nat_orc = orc_all[orc_all["Natureza"] == natureza] if not orc_all.empty else pd.DataFrame()

        real_grp = nat_real.groupby("DescGrupo")["Valor"].sum()
        orc_grp = nat_orc.groupby("DescGrupo")["ValorOrc"].sum() if not nat_orc.empty else pd.Series(dtype=float)

        all_groups = sorted(set(list(real_grp.index) + list(orc_grp.index)))
        for grp in all_groups:
            r_val = real_grp.get(grp, 0)
            o_val = orc_grp.get(grp, 0)
            var_abs = r_val - o_val
            var_pct = var_abs / o_val if o_val != 0 else 0
            rows_grupo.append({
                "Natureza": natureza,
                "Grupo": grp,
                "Realizado": r_val,
                "Orcado": o_val,
                "Var_Abs": var_abs,
                "Var_Pct": var_pct,
            })

    df_grupo = pd.DataFrame(rows_grupo)

    # Totais
    rec_real_total = df_grupo[df_grupo["Natureza"] == "Receita"]["Realizado"].sum()
    rec_orc_total = df_grupo[df_grupo["Natureza"] == "Receita"]["Orcado"].sum()
    desp_real_total = df_grupo[df_grupo["Natureza"] == "Despesa"]["Realizado"].sum()
    desp_orc_total = df_grupo[df_grupo["Natureza"] == "Despesa"]["Orcado"].sum()
    res_real = rec_real_total - desp_real_total
    res_orc = rec_orc_total - desp_orc_total

    # KPIs
    var_resultado = res_real - res_orc
    var_pct_resultado = var_resultado / abs(res_orc) if res_orc != 0 else 0

    cards = html.Div(
        style={"display": "flex", "gap": "12px", "flexWrap": "wrap"},
        children=[
            make_kpi_card("Resultado Real", res_real,
                         COLORS["receita"] if res_real >= 0 else COLORS["despesa"]),
            make_kpi_card("Resultado Orcado", res_orc,
                         COLORS["neutro"]),
            make_kpi_card("Variacao", var_resultado,
                         COLORS["receita"] if var_resultado >= 0 else COLORS["despesa"]),
            make_kpi_card("Var. %", var_pct_resultado,
                         COLORS["receita"] if var_pct_resultado >= 0 else COLORS["despesa"], formato="pct"),
        ],
    )

    # ---- TABELA ANALITICA COMPLETA (estilo relatorio) ----
    report_rows = []

    # RECEITAS
    report_rows.append({"Conta": "RECEITAS", "Realizado": rec_real_total,
                        "Orcado": rec_orc_total, "Var_Abs": rec_real_total - rec_orc_total,
                        "Var_Pct": (rec_real_total - rec_orc_total) / rec_orc_total if rec_orc_total else 0,
                        "nivel": 0})

    rec_grupos = df_grupo[df_grupo["Natureza"] == "Receita"].sort_values("Realizado", ascending=False)
    for _, rg in rec_grupos.iterrows():
        report_rows.append({"Conta": f"  {rg['Grupo']}", "Realizado": rg["Realizado"],
                            "Orcado": rg["Orcado"], "Var_Abs": rg["Var_Abs"],
                            "Var_Pct": rg["Var_Pct"], "nivel": 1})

        # Detalhe por conta dentro do grupo
        grp_real = real_pivot[(real_pivot["Natureza"] == "Receita") & (real_pivot["DescGrupo"] == rg["Grupo"])]
        grp_orc = orc_all[(orc_all["Natureza"] == "Receita") & (orc_all["DescGrupo"] == rg["Grupo"])] if not orc_all.empty else pd.DataFrame()

        det_real = grp_real.groupby("DescConta")["Valor"].sum()
        det_orc = grp_orc.groupby("DescConta")["ValorOrc"].sum() if not grp_orc.empty else pd.Series(dtype=float)
        all_contas = sorted(set(list(det_real.index) + list(det_orc.index)))

        for conta in all_contas:
            rv = det_real.get(conta, 0)
            ov = det_orc.get(conta, 0)
            if rv == 0 and ov == 0:
                continue
            va = rv - ov
            vp = va / ov if ov != 0 else 0
            report_rows.append({"Conta": f"    {conta}", "Realizado": rv,
                                "Orcado": ov, "Var_Abs": va, "Var_Pct": vp, "nivel": 2})

    # DESPESAS
    report_rows.append({"Conta": "DESPESAS", "Realizado": desp_real_total,
                        "Orcado": desp_orc_total, "Var_Abs": desp_real_total - desp_orc_total,
                        "Var_Pct": (desp_real_total - desp_orc_total) / desp_orc_total if desp_orc_total else 0,
                        "nivel": 0})

    desp_grupos = df_grupo[df_grupo["Natureza"] == "Despesa"].sort_values("Realizado", ascending=False)
    for _, rg in desp_grupos.iterrows():
        report_rows.append({"Conta": f"  {rg['Grupo']}", "Realizado": rg["Realizado"],
                            "Orcado": rg["Orcado"], "Var_Abs": rg["Var_Abs"],
                            "Var_Pct": rg["Var_Pct"], "nivel": 1})

        grp_real = real_pivot[(real_pivot["Natureza"] == "Despesa") & (real_pivot["DescGrupo"] == rg["Grupo"])]
        grp_orc = orc_all[(orc_all["Natureza"] == "Despesa") & (orc_all["DescGrupo"] == rg["Grupo"])] if not orc_all.empty else pd.DataFrame()

        det_real = grp_real.groupby("DescConta")["Valor"].sum()
        det_orc = grp_orc.groupby("DescConta")["ValorOrc"].sum() if not grp_orc.empty else pd.Series(dtype=float)
        all_contas = sorted(set(list(det_real.index) + list(det_orc.index)))

        for conta in all_contas:
            rv = det_real.get(conta, 0)
            ov = det_orc.get(conta, 0)
            if rv == 0 and ov == 0:
                continue
            va = rv - ov
            vp = va / ov if ov != 0 else 0
            report_rows.append({"Conta": f"    {conta}", "Realizado": rv,
                                "Orcado": ov, "Var_Abs": va, "Var_Pct": vp, "nivel": 2})

    # RESULTADO
    report_rows.append({"Conta": "RESULTADO (SUPERAVIT/DEFICIT)", "Realizado": res_real,
                        "Orcado": res_orc, "Var_Abs": var_resultado,
                        "Var_Pct": var_pct_resultado, "nivel": 0})

    # Formatar para exibicao
    df_report = pd.DataFrame(report_rows)
    df_display = df_report.copy()
    df_display["Realizado_fmt"] = df_display["Realizado"].apply(fmt_brl_full)
    df_display["Orcado_fmt"] = df_display["Orcado"].apply(fmt_brl_full)
    df_display["Var_Abs_fmt"] = df_display["Var_Abs"].apply(fmt_brl_full)
    df_display["Var_Pct_fmt"] = df_display["Var_Pct"].apply(fmt_pct)

    nivel0_idx = [i for i, r in enumerate(report_rows) if r["nivel"] == 0]
    nivel1_idx = [i for i, r in enumerate(report_rows) if r["nivel"] == 1]
    # Indices com variacao negativa
    var_neg_idx = [i for i, r in enumerate(report_rows) if r["Var_Abs"] < 0 and r["nivel"] <= 1]
    var_pos_idx = [i for i, r in enumerate(report_rows) if r["Var_Abs"] > 0 and r["nivel"] <= 1]

    tabela_report = dash_table.DataTable(
        id="tabela-oxr-analitico",
        data=df_display[["Conta", "Realizado_fmt", "Orcado_fmt", "Var_Abs_fmt", "Var_Pct_fmt"]].rename(
            columns={"Realizado_fmt": "Realizado", "Orcado_fmt": "Orcado",
                     "Var_Abs_fmt": "Variacao R$", "Var_Pct_fmt": "Var. %"}).to_dict("records"),
        columns=[
            {"name": "Conta", "id": "Conta"},
            {"name": "Realizado", "id": "Realizado"},
            {"name": "Orcado", "id": "Orcado"},
            {"name": "Variacao R$", "id": "Variacao R$"},
            {"name": "Var. %", "id": "Var. %"},
        ],
        style_header={
            "backgroundColor": COLORS["titulo_bar"], "color": COLORS["branco"],
            "fontWeight": "bold", "fontSize": "12px", "fontFamily": "Segoe UI",
            "textAlign": "center", "padding": "10px", "whiteSpace": "normal",
        },
        style_cell={
            "fontSize": "11px", "fontFamily": "Segoe UI", "padding": "6px 10px",
            "border": "1px solid #E0E0E0", "whiteSpace": "nowrap",
        },
        style_cell_conditional=[
            {"if": {"column_id": "Conta"}, "textAlign": "left", "width": "40%", "whiteSpace": "normal"},
            {"if": {"column_id": "Realizado"}, "textAlign": "right", "width": "15%"},
            {"if": {"column_id": "Orcado"}, "textAlign": "right", "width": "15%"},
            {"if": {"column_id": "Variacao R$"}, "textAlign": "right", "width": "15%"},
            {"if": {"column_id": "Var. %"}, "textAlign": "right", "width": "15%"},
        ],
        style_data_conditional=[
            # Nivel 0 = headers principais (RECEITAS, DESPESAS, RESULTADO)
            {"if": {"row_index": nivel0_idx},
             "fontWeight": "bold", "backgroundColor": "#0D47A1", "color": "#FFFFFF",
             "fontSize": "13px"},
            # Nivel 1 = subgrupos
            {"if": {"row_index": nivel1_idx},
             "fontWeight": "600", "backgroundColor": "#E3F2FD", "fontSize": "12px"},
            # Variacao negativa em vermelho (nos grupos)
            {"if": {"row_index": var_neg_idx, "column_id": "Variacao R$"},
             "color": COLORS["despesa"]},
            {"if": {"row_index": var_neg_idx, "column_id": "Var. %"},
             "color": COLORS["despesa"]},
            # Variacao positiva em verde (nos grupos)
            {"if": {"row_index": var_pos_idx, "column_id": "Variacao R$"},
             "color": COLORS["receita"]},
            {"if": {"row_index": var_pos_idx, "column_id": "Var. %"},
             "color": COLORS["receita"]},
            # Zebra para linhas de detalhe
            {"if": {"row_index": "odd"}, "backgroundColor": "#FAFAFA"},
        ],
        style_table={"borderRadius": "8px", "overflow": "hidden",
                     "boxShadow": "0 2px 4px rgba(0,0,0,0.08)"},
        page_size=200,
    )

    # ---- TABELA MENSAL (colunas por mes) ----
    meses_labels = [NOMES_MESES[m] for m in meses_ativos]
    monthly_rows = []

    for natureza, sinal in [("Receita", 1), ("Despesa", 1)]:
        nat_real = real_pivot[real_pivot["Natureza"] == natureza]
        nat_orc = orc_all[orc_all["Natureza"] == natureza] if not orc_all.empty else pd.DataFrame()

        monthly_rows.append({"Conta": natureza.upper() + "S", "tipo": "header"})

        grupos = nat_real.groupby("DescGrupo")["Valor"].sum().sort_values(ascending=False).index
        for grp in grupos:
            row_data = {"Conta": f"  {grp}", "tipo": "grupo"}
            for m in meses_ativos:
                real_m = nat_real[(nat_real["DescGrupo"] == grp) & (nat_real["MesNum"] == m)]["Valor"].sum()
                orc_m = nat_orc[(nat_orc["DescGrupo"] == grp) & (nat_orc["MesNum"] == m)]["ValorOrc"].sum() if not nat_orc.empty else 0
                row_data[f"Real_{NOMES_MESES[m]}"] = real_m
                row_data[f"Orc_{NOMES_MESES[m]}"] = orc_m
                row_data[f"Var_{NOMES_MESES[m]}"] = real_m - orc_m
            monthly_rows.append(row_data)

    # Adicionar resultado mensal
    monthly_rows.append({"Conta": "RESULTADO", "tipo": "header"})
    row_res = {"Conta": "  Superavit/Deficit", "tipo": "grupo"}
    for m in meses_ativos:
        rec_m = real_pivot[(real_pivot["Natureza"] == "Receita") & (real_pivot["MesNum"] == m)]["Valor"].sum()
        desp_m = real_pivot[(real_pivot["Natureza"] == "Despesa") & (real_pivot["MesNum"] == m)]["Valor"].sum()
        orc_rec_m = orc_all[(orc_all["Natureza"] == "Receita") & (orc_all["MesNum"] == m)]["ValorOrc"].sum() if not orc_all.empty else 0
        orc_desp_m = orc_all[(orc_all["Natureza"] == "Despesa") & (orc_all["MesNum"] == m)]["ValorOrc"].sum() if not orc_all.empty else 0
        row_res[f"Real_{NOMES_MESES[m]}"] = rec_m - desp_m
        row_res[f"Orc_{NOMES_MESES[m]}"] = orc_rec_m - orc_desp_m
        row_res[f"Var_{NOMES_MESES[m]}"] = (rec_m - desp_m) - (orc_rec_m - orc_desp_m)
    monthly_rows.append(row_res)

    # Montar colunas da tabela mensal
    month_columns = [{"name": "Conta", "id": "Conta"}]
    for m in meses_ativos:
        ml = NOMES_MESES[m]
        month_columns.append({"name": f"Real {ml}", "id": f"Real_{ml}"})
        month_columns.append({"name": f"Orc. {ml}", "id": f"Orc_{ml}"})
        month_columns.append({"name": f"Var. {ml}", "id": f"Var_{ml}"})

    # Formatar valores
    df_monthly = pd.DataFrame(monthly_rows)
    for col in df_monthly.columns:
        if col.startswith(("Real_", "Orc_", "Var_")):
            df_monthly[col] = df_monthly[col].apply(lambda v: fmt_brl_full(v) if pd.notna(v) and isinstance(v, (int, float)) else "-")

    header_idx = [i for i, r in enumerate(monthly_rows) if r.get("tipo") == "header"]
    grupo_idx = [i for i, r in enumerate(monthly_rows) if r.get("tipo") == "grupo"]

    # Identificar colunas de variacao para colorir
    var_cols = [f"Var_{NOMES_MESES[m]}" for m in meses_ativos]

    tabela_mensal = dash_table.DataTable(
        id="tabela-oxr-mensal",
        data=df_monthly.drop(columns=["tipo"], errors="ignore").to_dict("records"),
        columns=month_columns,
        style_header={
            "backgroundColor": COLORS["alerta"], "color": COLORS["branco"],
            "fontWeight": "bold", "fontSize": "11px", "fontFamily": "Segoe UI",
            "textAlign": "center", "padding": "8px 6px", "whiteSpace": "normal",
        },
        style_cell={
            "fontSize": "10px", "fontFamily": "Segoe UI", "padding": "5px 6px",
            "border": "1px solid #E0E0E0", "whiteSpace": "nowrap", "textAlign": "right",
            "minWidth": "90px",
        },
        style_cell_conditional=[
            {"if": {"column_id": "Conta"}, "textAlign": "left", "width": "22%",
             "whiteSpace": "normal", "minWidth": "180px"},
        ],
        style_data_conditional=[
            {"if": {"row_index": header_idx},
             "fontWeight": "bold", "backgroundColor": COLORS["alerta"], "color": "#FFFFFF",
             "fontSize": "12px"},
            {"if": {"row_index": grupo_idx},
             "fontWeight": "600", "backgroundColor": "#FFF8E1", "fontSize": "11px"},
        ],
        style_table={"borderRadius": "8px", "overflow": "auto",
                     "boxShadow": "0 2px 4px rgba(0,0,0,0.08)", "maxWidth": "100%"},
        page_size=200,
    )

    # ---- GRAFICO VARIACAO POR GRUPO ----
    # Mostrar maiores variacoes (positivas e negativas)
    df_var = df_grupo.copy()
    df_var = df_var[df_var["Orcado"] > 0]  # So grupos com orcamento
    df_var = df_var.sort_values("Var_Abs")

    fig_var = go.Figure()
    fig_var.add_trace(go.Bar(
        y=df_var["Grupo"], x=df_var["Var_Abs"], orientation="h",
        marker_color=[COLORS["receita"] if v >= 0 else COLORS["despesa"] for v in df_var["Var_Abs"]],
        text=df_var["Var_Abs"].apply(fmt_brl), textposition="outside", textfont=dict(size=10),
    ))
    fig_var.update_layout(**chart_layout("Variacao Real x Orcado por Grupo (R$)", 450))
    fig_var.update_layout(margin=dict(l=220))
    fig_var.add_vline(x=0, line_color=COLORS["neutro"], line_width=1)

    # ---- GRAFICO % EXECUCAO ORCAMENTARIA ----
    df_exec = df_grupo[df_grupo["Orcado"] > 0].copy()
    df_exec["PctExec"] = df_exec["Realizado"] / df_exec["Orcado"]
    df_exec = df_exec.sort_values("PctExec")

    fig_exec = go.Figure()
    fig_exec.add_trace(go.Bar(
        y=df_exec["Grupo"], x=df_exec["PctExec"], orientation="h",
        marker_color=[
            COLORS["receita"] if 0.9 <= p <= 1.1 else (COLORS["alerta"] if p < 0.9 else COLORS["despesa"])
            for p in df_exec["PctExec"]
        ],
        text=df_exec["PctExec"].apply(lambda v: f"{v*100:.0f}%"), textposition="outside",
        textfont=dict(size=10),
    ))
    fig_exec.add_vline(x=1.0, line_dash="dash", line_color=COLORS["neutro"], line_width=1.5,
                       annotation_text="100%", annotation_position="top right")
    fig_exec.update_layout(**chart_layout("Execucao Orcamentaria (%)", 450))
    fig_exec.update_layout(margin=dict(l=220))
    fig_exec.update_xaxes(tickformat=".0%")

    # ---- PERIODO DO RELATORIO ----
    periodo_texto = f"Periodo: {', '.join([NOMES_MESES_FULL[m] for m in meses_ativos])} / 2026"

    # ---- BOTAO IMPRIMIR PDF ----
    btn_print = html.Button(
        "Imprimir / Salvar PDF",
        id="btn-print-oxr",
        n_clicks=0,
        style={
            "backgroundColor": COLORS["titulo_bar"], "color": COLORS["branco"],
            "border": "none", "borderRadius": "6px", "padding": "10px 24px",
            "fontSize": "14px", "fontWeight": "600", "cursor": "pointer",
            "fontFamily": "Segoe UI", "boxShadow": "0 2px 4px rgba(0,0,0,0.15)",
            "marginRight": "12px",
        },
    )

    # CSS de impressao esta em assets/print.css (carregado automaticamente pelo Dash)

    return html.Div(id="page-oxr", children=[
        make_title_bar("Relatorio Analitico - Orcado x Realizado"),

        # Barra com periodo e botao
        html.Div(
            style={"display": "flex", "justifyContent": "space-between", "alignItems": "center",
                   "marginBottom": "16px", "flexWrap": "wrap", "gap": "12px"},
            children=[
                html.Div([
                    html.H4(periodo_texto, style={
                        "margin": "0", "color": COLORS["titulo_bar"], "fontSize": "16px",
                        "fontFamily": "Segoe UI",
                    }),
                    html.P(f"Receitas: {len(rec)} lanc. | Despesas: {len(desp)} lanc.", style={
                        "margin": "4px 0 0 0", "color": COLORS["neutro"], "fontSize": "12px",
                    }),
                ]),
                html.Div([btn_print], className="print-hide"),
            ],
        ),

        cards,

        # TABELA ANALITICA PRINCIPAL
        html.Div(
            style={"marginTop": "20px", "backgroundColor": COLORS["branco"], "borderRadius": "8px",
                   "padding": "16px", "boxShadow": "0 2px 4px rgba(0,0,0,0.08)"},
            children=[
                html.H4("Comparativo Analitico por Grupo e Conta", style={
                    "margin": "0 0 12px 0", "color": COLORS["titulo_bar"], "fontSize": "14px",
                    "fontFamily": "Segoe UI",
                }),
                tabela_report,
            ],
        ),

        # TABELA MENSAL
        html.Div(
            id="section-mensal",
            style={"marginTop": "20px", "backgroundColor": COLORS["branco"], "borderRadius": "8px",
                   "padding": "16px", "boxShadow": "0 2px 4px rgba(0,0,0,0.08)"},
            children=[
                html.H4("Detalhamento Mensal - Real x Orcado x Variacao", style={
                    "margin": "0 0 12px 0", "color": COLORS["alerta"], "fontSize": "14px",
                    "fontFamily": "Segoe UI",
                }),
                tabela_mensal,
            ],
        ),

        # GRAFICOS
        html.Div(
            id="section-graficos",
            style={"display": "grid", "gridTemplateColumns": "1fr 1fr", "gap": "16px", "marginTop": "20px"},
            children=[
                make_card_container(dcc.Graph(id="graph-oxr-var", figure=fig_var, config={"displayModeBar": False})),
                make_card_container(dcc.Graph(id="graph-oxr-exec", figure=fig_exec, config={"displayModeBar": False})),
            ],
        ),

        # Rodape do relatorio
        html.Div(
            style={"marginTop": "24px", "padding": "12px 16px", "backgroundColor": "#ECEFF1",
                   "borderRadius": "8px", "textAlign": "center"},
            children=[
                html.P(
                    f"Federacao Paulista de Golfe - Relatorio Orcado x Realizado - {periodo_texto}",
                    style={"margin": "0", "fontSize": "11px", "color": COLORS["neutro"],
                           "fontFamily": "Segoe UI"},
                ),
                html.P(
                    "Gerado automaticamente pelo Dashboard Financeiro",
                    style={"margin": "2px 0 0 0", "fontSize": "10px", "color": "#9E9E9E",
                           "fontFamily": "Segoe UI"},
                ),
            ],
        ),
    ])


# ============================================================
# INICIAR SERVIDOR
# ============================================================

if __name__ == "__main__":
    import socket

    hostname = socket.gethostname()
    try:
        local_ip = socket.gethostbyname(hostname)
    except Exception:
        local_ip = "127.0.0.1"

    print("=" * 60)
    print("  DASHBOARD FINANCEIRO - FEDERACAO PAULISTA DE GOLFE")
    print("=" * 60)
    print(f"  Acesso local:    http://localhost:8050")
    print(f"  Acesso na rede:  http://{local_ip}:8050")
    print(f"  Razao:           {RAZAO_PATH}")
    print(f"  Orcamento:       {ORC_PATH}")
    print(f"  Meses:           {[NOMES_MESES[m] for m in MESES_DISP]}")
    print(f"  Receitas:        {len(df_receitas)} lancamentos")
    print(f"  Despesas:        {len(df_despesas)} lancamentos")
    print("=" * 60)

    app.run(debug=False, host="0.0.0.0", port=8050)
