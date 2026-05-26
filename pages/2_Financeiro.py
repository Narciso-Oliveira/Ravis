#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Pagina Financeiro - Ravi's Pet Shop
Fonte principal: tabela geral.xlsx
Fonte de apoio opcional: export (12).xlsx e apoio_financeiro.xlsx
"""

import os
import re
import unicodedata
from datetime import datetime

import pandas as pd
import plotly.graph_objects as go
import streamlit as st


# ============================================================
# CONFIGURACAO DA PAGINA
# ============================================================
st.set_page_config(
    page_title="Financeiro | Ravi's Pet Shop",
    page_icon="💸",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# CSS CUSTOMIZADO
# ============================================================
st.markdown(
    """
<style>
    .stApp { background-color: #0E1117; }

    .finance-card {
        background: linear-gradient(135deg, #1B1F2B 0%, #1E2235 100%);
        border-radius: 14px;
        padding: 18px 20px;
        border-left: 5px solid #00D4AA;
        margin-bottom: 12px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.30);
        min-height: 142px;
    }
    .finance-card h4 {
        color: #9CA3AF;
        font-size: 0.76rem;
        font-weight: 800;
        text-transform: uppercase;
        letter-spacing: 0.9px;
        margin: 0 0 8px 0;
        font-family: 'Segoe UI', sans-serif;
    }
    .finance-card .value {
        color: #00D4AA;
        font-size: 1.55rem;
        font-weight: 900;
        margin: 0;
        line-height: 1.1;
        font-family: 'Segoe UI', sans-serif;
    }
    .finance-card .sub {
        color: #7B8194;
        font-size: 0.78rem;
        margin-top: 8px;
        font-weight: 600;
        line-height: 1.35;
    }

    .definition-card {
        background: linear-gradient(135deg, #141824 0%, #1B1F2B 100%);
        border-radius: 14px;
        padding: 17px 19px;
        border-top: 4px solid #00D4AA;
        box-shadow: 0 4px 15px rgba(0,0,0,0.25);
        min-height: 118px;
    }
    .definition-card h4 {
        color: #00D4AA;
        font-size: 0.86rem;
        font-weight: 800;
        text-transform: uppercase;
        letter-spacing: 0.9px;
        margin: 0 0 8px 0;
    }
    .definition-card p {
        color: #D1D5DB;
        font-size: 0.90rem;
        margin: 0;
        font-weight: 500;
        line-height: 1.45;
    }

    .config-panel {
        background: linear-gradient(135deg, #1B1F2B 0%, #1E2235 100%);
        border-radius: 16px;
        padding: 18px 22px;
        margin: 10px 0 22px 0;
        border-left: 5px solid #00D4AA;
        box-shadow: 0 4px 15px rgba(0,0,0,0.28);
    }
    .config-panel .title {
        color:#00D4AA;
        font-size:1.1rem;
        font-weight:900;
        margin-bottom:4px;
    }
    .config-panel .subtitle {
        color:#9CA3AF;
        font-size:0.85rem;
        font-weight:600;
    }

    .insight-box {
        background: linear-gradient(135deg, #1B1F2B 0%, #1E2235 100%);
        border-radius: 12px;
        padding: 16px 20px;
        margin-bottom: 10px;
        border-left: 4px solid #FFA726;
        box-shadow: 0 2px 10px rgba(0,0,0,0.22);
    }
    .insight-box p {
        color: #D1D5DB;
        font-size: 0.95rem;
        margin: 0;
        font-weight: 500;
        line-height: 1.5;
    }

    .header-title {
        color: #FFFFFF;
        font-size: 2.2rem;
        font-weight: 900;
        margin: 0;
        letter-spacing: -0.5px;
        font-family: 'Segoe UI', sans-serif;
    }
    .header-sub {
        color: #7B8194;
        font-size: 1rem;
        font-weight: 600;
        margin-top: 0;
    }

    .small-muted {
        color:#7B8194;
        font-size:0.82rem;
        font-weight:600;
        line-height:1.35;
    }

    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    section[data-testid="stSidebar"] {
        background-color: #151822;
    }
    section[data-testid="stSidebar"] .stMarkdown p,
    section[data-testid="stSidebar"] .stMarkdown h2,
    section[data-testid="stSidebar"] .stMarkdown h3 {
        color: #D1D5DB !important;
    }
</style>
""",
    unsafe_allow_html=True,
)


# ============================================================
# CONSTANTES
# ============================================================
PRIMARY = "#00D4AA"
GREEN = "#00D4AA"
RED = "#FF5252"
ORANGE = "#FFA726"
YELLOW = "#FFD166"
WHITE = "#FFFFFF"
MUTED = "#7B8194"
CARD_BG = "#1B1F2B"

MESES_PT = {
    "janeiro": 1,
    "fevereiro": 2,
    "março": 3,
    "marco": 3,
    "abril": 4,
    "maio": 5,
    "junho": 6,
    "julho": 7,
    "agosto": 8,
    "setembro": 9,
    "outubro": 10,
    "novembro": 11,
    "dezembro": 12,
}

MESES_NOME = {
    "01": "janeiro",
    "02": "fevereiro",
    "03": "março",
    "04": "abril",
    "05": "maio",
    "06": "junho",
    "07": "julho",
    "08": "agosto",
    "09": "setembro",
    "10": "outubro",
    "11": "novembro",
    "12": "dezembro",
}

SETOR_ORDER = ["Banho", "Serviços extras", "Loja", "Clínica", "Outros"]

BASE_FILE_CANDIDATES = [
    "tabela geral.xlsx",
    "data/tabela geral.xlsx",
    "/mnt/data/tabela geral.xlsx",
]

EXPORT_FILE_CANDIDATES = [
    "export (12).xlsx",
    "data/export (12).xlsx",
    "/mnt/data/export (12).xlsx",
]

APOIO_FILE_CANDIDATES = [
    "apoio_financeiro.xlsx",
    "data/apoio_financeiro.xlsx",
    "/mnt/data/apoio_financeiro.xlsx",
]

APOIO_COLUNAS = [
    "Mes_Ano",
    "Meta_Faturamento",
    "Custos_Fixos",
    "Custos_Variaveis",
    "Impostos",
    "Pro_Labore",
    "Outros_Custos",
    "Faturamento_Ano_Anterior",
    "Observacoes",
]


# ============================================================
# FUNCOES AUXILIARES
# ============================================================
def find_existing_file(candidates):
    """Retorna o primeiro arquivo existente da lista de candidatos."""
    for path in candidates:
        if os.path.exists(path):
            return path
    return None


def normalize_text(value):
    """Normaliza texto para comparacoes sem acento e sem caixa."""
    if pd.isna(value):
        return ""
    text = str(value).strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = text.encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"\s+", " ", text)
    return text


def parse_date_pt(text):
    """Converte datas em portugues, datas ISO, timestamps e seriais Excel."""
    if pd.isna(text):
        return pd.NaT

    if isinstance(text, (pd.Timestamp, datetime)):
        return pd.to_datetime(text, errors="coerce")

    if isinstance(text, (int, float)) and not pd.isna(text):
        try:
            value = float(text)
            if 20000 <= value <= 80000:
                return pd.to_datetime(value, unit="D", origin="1899-12-30", errors="coerce")
        except Exception:
            pass

    raw = str(text).strip()
    if not raw:
        return pd.NaT

    match = re.search(r"(\d{1,2})\s+de\s+([a-zA-ZçÇãÃáÁéÉíÍóÓúÚôÔõÕ]+)\s+de\s+(\d{4})", raw)
    if match:
        day = int(match.group(1))
        month_name = normalize_text(match.group(2))
        year = int(match.group(3))
        month = MESES_PT.get(month_name)
        if month:
            return pd.Timestamp(year, month, day)

    return pd.to_datetime(raw, dayfirst=True, errors="coerce")


def parse_mes_ano(value):
    """Converte diferentes formatos de mes para YYYY-MM."""
    if pd.isna(value):
        return ""

    if isinstance(value, (pd.Timestamp, datetime)):
        dt = pd.to_datetime(value, errors="coerce")
        return "" if pd.isna(dt) else dt.strftime("%Y-%m")

    raw = str(value).strip()
    if not raw:
        return ""

    match_iso = re.search(r"(\d{4})[-/](\d{1,2})", raw)
    if match_iso:
        return f"{int(match_iso.group(1)):04d}-{int(match_iso.group(2)):02d}"

    match_br = re.search(r"(\d{1,2})[-/](\d{4})", raw)
    if match_br:
        return f"{int(match_br.group(2)):04d}-{int(match_br.group(1)):02d}"

    dt = parse_date_pt(raw)
    if pd.notna(dt):
        return dt.strftime("%Y-%m")

    return raw


def nome_mes(mes_ano):
    """Converte '2026-04' para 'abril/2026'."""
    try:
        year, month = str(mes_ano).split("-")
        return f"{MESES_NOME.get(month.zfill(2), month)}/{year}"
    except Exception:
        return str(mes_ano)


def money_to_float(value):
    """Converte valores monetarios em numero, aceitando R$, virgula e ponto."""
    if pd.isna(value):
        return 0.0

    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text:
        return 0.0

    text = re.sub(r"[^\d,.\-]", "", text)
    if not text or text in {"-", ",", "."}:
        return 0.0

    try:
        if "," in text and text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        elif text.count(".") > 1 and "," not in text:
            text = text.replace(".", "")
        return float(text)
    except Exception:
        return 0.0


def series_money_to_float(series):
    """Converte uma Series para numerico monetario."""
    if series is None:
        return pd.Series(dtype=float)
    return series.apply(money_to_float).astype(float)


def format_currency(value):
    """Formata numero em moeda brasileira."""
    try:
        value = float(value)
    except Exception:
        value = 0.0

    sign = "-" if value < 0 else ""
    value = abs(value)
    formatted = f"{value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{sign}R$ {formatted}"


def format_percent(value):
    """Formata percentual em padrao brasileiro."""
    try:
        value = float(value)
    except Exception:
        value = 0.0
    return f"{value:.1f}%".replace(".", ",")


def safe_divide(numerator, denominator, multiplier=1.0):
    """Evita divisao por zero."""
    try:
        numerator = float(numerator)
        denominator = float(denominator)
    except Exception:
        return 0.0
    if denominator == 0:
        return 0.0
    return (numerator / denominator) * multiplier


def previous_month(mes_ano):
    try:
        return (pd.Period(mes_ano, freq="M") - 1).strftime("%Y-%m")
    except Exception:
        return ""


def same_month_previous_year(mes_ano):
    try:
        period = pd.Period(mes_ano, freq="M")
        return (period - 12).strftime("%Y-%m")
    except Exception:
        return ""


def classify_setor(row_or_value):
    """Classifica faturamento por setor usando Classificacao como base."""
    if isinstance(row_or_value, pd.Series):
        classificacao = normalize_text(row_or_value.get("Classificação", ""))
        servico = normalize_text(row_or_value.get("Servico/Produto", ""))
        categoria_produto = normalize_text(row_or_value.get("Categoria do Produto", ""))
        categoria_vet = normalize_text(row_or_value.get("Categoria do Servico Veterinario", ""))
        text = " ".join([classificacao, servico, categoria_produto, categoria_vet])
    else:
        classificacao = normalize_text(row_or_value)
        text = classificacao

    if "banho" in classificacao:
        return "Banho"

    if "servicos extras" in classificacao or "servico extra" in classificacao or "extra" in classificacao:
        return "Serviços extras"

    if "loja" in classificacao:
        return "Loja"

    if "clinica" in classificacao or "veterin" in classificacao:
        return "Clínica"

    if "loja" in text or "revenda" in text or "produto" in text or "venda" in text:
        return "Loja"

    if "clinica" in text or "veterin" in text or "vacina" in text or "consulta" in text:
        return "Clínica"

    if "tosa" in text or "hidratacao" in text or "escovacao" in text or "unha" in text:
        return "Serviços extras"

    if "banho" in text:
        return "Banho"

    return "Outros"


def classify_tipo_receita(row):
    """Cria uma categoria legivel para o filtro de tipo de receita."""
    setor = row.get("Setor", "Outros")
    tipo_banho = str(row.get("Tipo Banho", "")).strip()

    if setor == "Banho":
        if tipo_banho and tipo_banho.lower() not in {"nan", "none", "outros"}:
            return f"Banho - {tipo_banho}"
        return "Banho"

    if setor == "Serviços extras":
        return "Serviços extras"

    if setor == "Loja":
        return "Loja"

    if setor == "Clínica":
        return "Clínica"

    return "Outros"


def classify_export_setor(row):
    """Classifica receitas/despesas da planilha export por centro de custo e categoria."""
    text = normalize_text(
        " ".join(
            [
                str(row.get("Centro de Custo", "")),
                str(row.get("Categoria", "")),
                str(row.get("Tipo Despesa", "")),
                str(row.get("Descricao", "")),
            ]
        )
    )

    if "banho" in text or "tosa" in text:
        return "Banho"

    if "loja" in text or "revenda" in text or "produto" in text or "venda" in text:
        return "Loja"

    if "clinica" in text or "veterin" in text or "consulta" in text or "vacina" in text:
        return "Clínica"

    if "extra" in text:
        return "Serviços extras"

    return "Outros"


def classify_expense_bucket(row):
    """Classifica despesas em grupos financeiros usados nos calculos."""
    text = normalize_text(
        " ".join(
            [
                str(row.get("Categoria", "")),
                str(row.get("Tipo Despesa", "")),
                str(row.get("Centro de Custo", "")),
                str(row.get("Descricao", "")),
                str(row.get("Conta", "")),
            ]
        )
    )

    if any(term in text for term in ["imposto", "icms", "simples", "tribut", "fiscal"]):
        return "Impostos"

    if any(term in text for term in ["pro labore", "pro-labore", "retirada", "socio", "socios"]):
        return "Pró-labore"

    if any(
        term in text
        for term in [
            "compra",
            "produto",
            "revenda",
            "itens de consumo",
            "insumo",
            "maquininha",
            "taxa maquina",
            "tx maquininha",
            "comissao",
        ]
    ):
        return "Custos variáveis"

    if any(
        term in text
        for term in [
            "folha fixa",
            "salario",
            "funcionario",
            "aluguel",
            "contabilidade",
            "contador",
            "sistema",
            "administrativa",
            "operacional",
            "energia",
            "eletricidade",
            "internet",
            "seguranca",
            "manutencao",
            "limpeza",
            "reserva trabalh",
            "transporte",
            "alimentacao",
        ]
    ):
        return "Custos fixos"

    return "Outros custos"


def origem_fluxo(natureza):
    """Classifica Natureza em Receita, Despesa ou Indefinido."""
    text = normalize_text(natureza)
    if any(term in text for term in ["receita", "recebimento", "entrada", "venda"]):
        return "Receita"
    if any(term in text for term in ["despesa", "custo", "saida", "pagamento"]):
        return "Despesa"
    return "Indefinido"


# ============================================================
# LEITURA E LIMPEZA DOS DADOS
# ============================================================
@st.cache_data(show_spinner=False)
def load_data():
    """Carrega tabela geral.xlsx de forma tolerante a colunas ausentes."""
    path = find_existing_file(BASE_FILE_CANDIDATES)
    if not path:
        return pd.DataFrame(), None, ["Arquivo 'tabela geral.xlsx' não encontrado."]

    try:
        df = pd.read_excel(path, engine="openpyxl")
    except Exception as exc:
        return pd.DataFrame(), path, [f"Erro ao ler 'tabela geral.xlsx': {exc}"]

    original_columns = set(df.columns)
    missing = []

    expected_defaults = {
        "Data Realizacao": None,
        "Valor Faturado": 0.0,
        "Classificação": "Outros",
        "Cliente": "",
        "Tipo Banho": "",
        "Profissional": "",
        "Servico/Produto": "",
        "Chave Atendimento": "",
        "Categoria do Produto": "",
        "Categoria do Servico Veterinario": "",
    }

    for col, default in expected_defaults.items():
        if col not in df.columns:
            df[col] = default
            missing.append(col)

    if "Valor Faturado" not in original_columns and "Valor Total do Pedido" in df.columns:
        df["Valor Faturado"] = df["Valor Total do Pedido"]

    df["Data"] = df["Data Realizacao"].apply(parse_date_pt)
    if df["Data"].isna().all() and "Data Realizacao.old" in df.columns:
        df["Data"] = df["Data Realizacao.old"].apply(parse_date_pt)

    df["Mes_Ano"] = df["Data"].dt.strftime("%Y-%m").fillna("")
    df["Valor Faturado"] = series_money_to_float(df["Valor Faturado"]).fillna(0.0)
    df["Classificação"] = df["Classificação"].fillna("Outros").astype(str)
    df["Cliente"] = df["Cliente"].fillna("").astype(str)
    df["Profissional"] = df["Profissional"].fillna("").astype(str)
    df["Tipo Banho"] = df["Tipo Banho"].fillna("").astype(str)
    df["Servico/Produto"] = df["Servico/Produto"].fillna("").astype(str)
    df["Chave Atendimento"] = df["Chave Atendimento"].fillna("").astype(str)

    df["Setor"] = df.apply(classify_setor, axis=1)
    df["Tipo_Receita"] = df.apply(classify_tipo_receita, axis=1)

    messages = []
    if missing:
        messages.append("Colunas ausentes criadas com valores padrão: " + ", ".join(missing))

    if df["Data"].isna().any():
        n_bad = int(df["Data"].isna().sum())
        messages.append(f"{n_bad} linhas sem data válida foram mantidas, mas não entram nos filtros mensais.")

    return df, path, messages


@st.cache_data(show_spinner=False)
def load_financial_support():
    """Carrega apoio_financeiro.xlsx e export (12).xlsx, ambos opcionais."""
    messages = []

    # Apoio financeiro estruturado
    apoio_path = find_existing_file(APOIO_FILE_CANDIDATES)
    if apoio_path:
        try:
            apoio_df = pd.read_excel(apoio_path, engine="openpyxl")
        except Exception as exc:
            apoio_df = pd.DataFrame(columns=APOIO_COLUNAS)
            messages.append(f"Erro ao ler apoio_financeiro.xlsx: {exc}")
    else:
        apoio_df = pd.DataFrame(columns=APOIO_COLUNAS)

    for col in APOIO_COLUNAS:
        if col not in apoio_df.columns:
            apoio_df[col] = "" if col in {"Mes_Ano", "Observacoes"} else 0.0

    if not apoio_df.empty:
        apoio_df["Mes_Ano"] = apoio_df["Mes_Ano"].apply(parse_mes_ano)
        for col in [
            "Meta_Faturamento",
            "Custos_Fixos",
            "Custos_Variaveis",
            "Impostos",
            "Pro_Labore",
            "Outros_Custos",
            "Faturamento_Ano_Anterior",
        ]:
            apoio_df[col] = series_money_to_float(apoio_df[col]).fillna(0.0)

    # Export financeiro geral
    export_path = find_existing_file(EXPORT_FILE_CANDIDATES)
    if export_path:
        try:
            export_df = pd.read_excel(export_path, engine="openpyxl")
        except Exception as exc:
            export_df = pd.DataFrame()
            messages.append(f"Erro ao ler export (12).xlsx: {exc}")
    else:
        export_df = pd.DataFrame()

    expected_export_columns = [
        "Natureza",
        "Categoria",
        "Centro de Custo",
        "Tipo Despesa",
        "Conta",
        "Caixa",
        "Fornecedor",
        "Cliente",
        "Funcionario",
        "Descricao",
        "Valor",
        "Valor Bruto",
        "Data do pagamento/recebimento",
        "Data do lancamento",
        "Data da competencia",
        "Data do vencimento",
        "Fatura Cancelada?",
    ]

    for col in expected_export_columns:
        if col not in export_df.columns:
            export_df[col] = None

    if not export_df.empty:
        cancel_col = "Fatura Cancelada?"
        cancel_mask = export_df[cancel_col].apply(lambda value: normalize_text(value) in {"sim", "s", "true", "1"})
        export_df = export_df[~cancel_mask].copy()

        export_df["Valor"] = series_money_to_float(export_df["Valor"]).fillna(0.0)
        export_df["Valor Bruto"] = series_money_to_float(export_df["Valor Bruto"]).fillna(0.0)

        date_cols = [
            "Data do pagamento/recebimento",
            "Data do lancamento",
            "Data da competencia",
            "Data do vencimento",
        ]
        export_df["Data"] = pd.NaT
        for col in date_cols:
            parsed = export_df[col].apply(parse_date_pt)
            export_df["Data"] = export_df["Data"].fillna(parsed)

        export_df["Mes_Ano"] = export_df["Data"].dt.strftime("%Y-%m").fillna("")
        export_df["Fluxo"] = export_df["Natureza"].apply(origem_fluxo)
        export_df["Setor"] = export_df.apply(classify_export_setor, axis=1)
        export_df["Bucket_Custo"] = export_df.apply(classify_expense_bucket, axis=1)
        export_df["Tipo_Receita"] = export_df["Categoria"].fillna("Receita apoio").astype(str)
        export_df["Cliente"] = export_df["Cliente"].fillna("").astype(str)

    files_info = {
        "apoio_path": apoio_path,
        "export_path": export_path,
    }

    return apoio_df, export_df, files_info, messages


def calcular_defaults_financeiros(apoio_df, export_df, df_raw, meses_sel):
    """Calcula valores iniciais dos inputs a partir das planilhas de apoio."""
    meses_sel = [m for m in meses_sel if m]
    n_meses = max(len(meses_sel), 1)

    defaults = {
        "meta_mensal": 0.0,
        "custos_fixos_mensais": 0.0,
        "custos_variaveis_mensais": 0.0,
        "impostos_mensais": 0.0,
        "pro_labore_mensal": 0.0,
        "outros_custos_mensais": 0.0,
        "faturamento_ano_anterior": 0.0,
        "observacoes": "",
        "origem_defaults": "manual",
    }

    apoio_periodo = pd.DataFrame()
    if not apoio_df.empty and meses_sel:
        apoio_periodo = apoio_df[apoio_df["Mes_Ano"].isin(meses_sel)].copy()

    if not apoio_periodo.empty:
        defaults["origem_defaults"] = "apoio_financeiro.xlsx"
        defaults["meta_mensal"] = float(apoio_periodo["Meta_Faturamento"].mean())
        defaults["custos_fixos_mensais"] = float(apoio_periodo["Custos_Fixos"].mean())
        defaults["custos_variaveis_mensais"] = float(apoio_periodo["Custos_Variaveis"].mean())
        defaults["impostos_mensais"] = float(apoio_periodo["Impostos"].mean())
        defaults["pro_labore_mensal"] = float(apoio_periodo["Pro_Labore"].mean())
        defaults["outros_custos_mensais"] = float(apoio_periodo["Outros_Custos"].mean())

        latest_month = max(meses_sel)
        latest_row = apoio_periodo[apoio_periodo["Mes_Ano"] == latest_month]
        if not latest_row.empty:
            defaults["faturamento_ano_anterior"] = float(latest_row["Faturamento_Ano_Anterior"].iloc[0])
            obs = latest_row["Observacoes"].dropna().astype(str)
            defaults["observacoes"] = obs.iloc[0] if not obs.empty else ""

    elif not export_df.empty and meses_sel:
        desp_periodo = export_df[
            (export_df["Fluxo"] == "Despesa")
            & (export_df["Mes_Ano"].isin(meses_sel))
        ].copy()

        if not desp_periodo.empty:
            defaults["origem_defaults"] = "export (12).xlsx"
            bucket_sum = desp_periodo.groupby("Bucket_Custo")["Valor"].sum().to_dict()
            defaults["custos_fixos_mensais"] = float(bucket_sum.get("Custos fixos", 0.0) / n_meses)
            defaults["custos_variaveis_mensais"] = float(bucket_sum.get("Custos variáveis", 0.0) / n_meses)
            defaults["impostos_mensais"] = float(bucket_sum.get("Impostos", 0.0) / n_meses)
            defaults["pro_labore_mensal"] = float(bucket_sum.get("Pró-labore", 0.0) / n_meses)
            defaults["outros_custos_mensais"] = float(bucket_sum.get("Outros custos", 0.0) / n_meses)

    mes_ref = max(meses_sel) if meses_sel else ""
    mes_aa = same_month_previous_year(mes_ref)
    if mes_aa and defaults["faturamento_ano_anterior"] == 0 and not df_raw.empty:
        hist = df_raw[df_raw["Mes_Ano"] == mes_aa]["Valor Faturado"].sum()
        defaults["faturamento_ano_anterior"] = float(hist)

    return defaults


def aplicar_filtros_sistema(df, meses_sel=None, setores_sel=None, tipos_sel=None, prof_sel=None, cli_sel=None):
    """Aplica filtros na base do sistema sem quebrar quando filtros estao vazios."""
    if df.empty:
        return df.copy()

    filtered = df.copy()

    if meses_sel:
        filtered = filtered[filtered["Mes_Ano"].isin(meses_sel)]

    if setores_sel:
        filtered = filtered[filtered["Setor"].isin(setores_sel)]

    if tipos_sel:
        filtered = filtered[filtered["Tipo_Receita"].isin(tipos_sel)]

    if prof_sel:
        filtered = filtered[filtered["Profissional"].isin(prof_sel)]

    if cli_sel:
        filtered = filtered[filtered["Cliente"].isin(cli_sel)]

    return filtered


def aplicar_filtros_export_receitas(export_df, meses_sel=None, setores_sel=None, tipos_sel=None, cli_sel=None):
    """Aplica filtros em receitas da planilha export."""
    if export_df.empty:
        return export_df.copy()

    filtered = export_df[export_df["Fluxo"] == "Receita"].copy()

    if meses_sel:
        filtered = filtered[filtered["Mes_Ano"].isin(meses_sel)]

    if setores_sel:
        filtered = filtered[filtered["Setor"].isin(setores_sel)]

    if tipos_sel:
        filtered = filtered[filtered["Tipo_Receita"].isin(tipos_sel)]

    if cli_sel and "Cliente" in filtered.columns:
        filtered = filtered[filtered["Cliente"].isin(cli_sel)]

    return filtered


def receita_por_mes(df_sistema_sem_mes, export_receitas_sem_mes):
    """Retorna serie mensal de faturamento, juntando sistema e receitas de apoio."""
    receita_sistema = pd.Series(dtype=float)
    receita_export = pd.Series(dtype=float)

    if not df_sistema_sem_mes.empty:
        receita_sistema = df_sistema_sem_mes.groupby("Mes_Ano")["Valor Faturado"].sum()

    if not export_receitas_sem_mes.empty:
        receita_export = export_receitas_sem_mes.groupby("Mes_Ano")["Valor"].sum()

    mensal = receita_sistema.add(receita_export, fill_value=0.0)
    mensal = mensal[mensal.index.astype(str) != ""]
    return mensal.sort_index()


def calcular_indicadores_financeiros(
    df_raw,
    df_f,
    df_sem_filtro_mes,
    export_receitas_f,
    export_receitas_sem_mes,
    meses_sel,
    inputs,
):
    """Calcula os KPIs, series e tabela mensal financeira."""
    meses_sel = sorted([m for m in meses_sel if m])
    n_meses = max(len(meses_sel), 1)
    mes_ref = max(meses_sel) if meses_sel else ""

    receita_sistema = float(df_f["Valor Faturado"].sum()) if not df_f.empty else 0.0
    receita_export = float(export_receitas_f["Valor"].sum()) if not export_receitas_f.empty else 0.0
    faturamento_total = receita_sistema + receita_export

    clientes_unicos = df_f["Cliente"].replace("", pd.NA).dropna().nunique() if not df_f.empty else 0
    ticket_medio = safe_divide(faturamento_total, clientes_unicos)

    meta_total = float(inputs["meta_mensal"]) * n_meses
    custos_fixos = float(inputs["custos_fixos_mensais"]) * n_meses
    custos_variaveis = float(inputs["custos_variaveis_mensais"]) * n_meses
    impostos = float(inputs["impostos_mensais"]) * n_meses
    pro_labore = float(inputs["pro_labore_mensal"]) * n_meses
    outros_custos = float(inputs["outros_custos_mensais"]) * n_meses

    custos_operacionais = custos_fixos + custos_variaveis + outros_custos
    despesas_totais = custos_operacionais + impostos + pro_labore

    lucro_operacional = faturamento_total - custos_operacionais
    lucro_liquido = faturamento_total - despesas_totais

    margem_operacional = safe_divide(lucro_operacional, faturamento_total, 100)
    margem_liquida = safe_divide(lucro_liquido, faturamento_total, 100)
    meta_pct = safe_divide(faturamento_total, meta_total, 100)
    diff_equilibrio = faturamento_total - despesas_totais

    receita_setor = pd.Series(0.0, index=SETOR_ORDER)
    if not df_f.empty:
        receita_setor = receita_setor.add(df_f.groupby("Setor")["Valor Faturado"].sum(), fill_value=0.0)
    if not export_receitas_f.empty:
        receita_setor = receita_setor.add(export_receitas_f.groupby("Setor")["Valor"].sum(), fill_value=0.0)
    receita_setor = receita_setor.reindex(SETOR_ORDER, fill_value=0.0)

    participacao_setor = receita_setor.apply(lambda value: safe_divide(value, faturamento_total, 100))

    mensal_receita = receita_por_mes(df_sem_filtro_mes, export_receitas_sem_mes)
    receita_mes_ref = float(mensal_receita.get(mes_ref, 0.0)) if mes_ref else faturamento_total
    mes_anterior = previous_month(mes_ref)
    receita_mes_anterior = float(mensal_receita.get(mes_anterior, 0.0)) if mes_anterior else 0.0
    crescimento_mensal = safe_divide(receita_mes_ref - receita_mes_anterior, receita_mes_anterior, 100)

    mes_aa = same_month_previous_year(mes_ref)
    receita_aa_hist = float(mensal_receita.get(mes_aa, 0.0)) if mes_aa else 0.0
    receita_aa = receita_aa_hist if receita_aa_hist > 0 else float(inputs["faturamento_ano_anterior"])
    crescimento_anual = safe_divide(receita_mes_ref - receita_aa, receita_aa, 100)

    # Linha mensal para tendencia: usa todos os meses da base filtrada, nao apenas o periodo selecionado
    mensal_tendencia = mensal_receita.reset_index()
    mensal_tendencia.columns = ["Mes_Ano", "Faturamento"]
    if not mensal_tendencia.empty:
        mensal_tendencia["Mes"] = mensal_tendencia["Mes_Ano"].apply(nome_mes)

    # Tabela mensal do periodo selecionado
    meses_tabela = meses_sel if meses_sel else list(mensal_receita.index)
    rows = []

    for mes in meses_tabela:
        fat_mes = float(mensal_receita.get(mes, 0.0))
        meta_mes = float(inputs["meta_mensal"])
        fixos_mes = float(inputs["custos_fixos_mensais"])
        variaveis_mes = float(inputs["custos_variaveis_mensais"])
        impostos_mes = float(inputs["impostos_mensais"])
        pro_labore_mes = float(inputs["pro_labore_mensal"])
        outros_mes = float(inputs["outros_custos_mensais"])

        custos_op_mes = fixos_mes + variaveis_mes + outros_mes
        despesas_mes = custos_op_mes + impostos_mes + pro_labore_mes
        lucro_op_mes = fat_mes - custos_op_mes
        lucro_liq_mes = fat_mes - despesas_mes

        df_mes = df_sem_filtro_mes[df_sem_filtro_mes["Mes_Ano"] == mes] if not df_sem_filtro_mes.empty else pd.DataFrame()
        clientes_mes = df_mes["Cliente"].replace("", pd.NA).dropna().nunique() if not df_mes.empty else 0
        ticket_mes = safe_divide(fat_mes, clientes_mes)

        mes_ant = previous_month(mes)
        fat_ant = float(mensal_receita.get(mes_ant, 0.0)) if mes_ant else 0.0
        crescimento_mes = safe_divide(fat_mes - fat_ant, fat_ant, 100)

        mes_ano_ant = same_month_previous_year(mes)
        fat_ano_ant = float(mensal_receita.get(mes_ano_ant, 0.0)) if mes_ano_ant else 0.0
        if fat_ano_ant == 0.0 and mes == mes_ref:
            fat_ano_ant = float(inputs["faturamento_ano_anterior"])
        crescimento_ano = safe_divide(fat_mes - fat_ano_ant, fat_ano_ant, 100)

        rows.append(
            {
                "Mês": nome_mes(mes),
                "Mes_Ano": mes,
                "Faturamento": fat_mes,
                "Meta": meta_mes,
                "% Meta": safe_divide(fat_mes, meta_mes, 100),
                "Custos fixos": fixos_mes,
                "Custos variáveis": variaveis_mes,
                "Impostos": impostos_mes,
                "Pró-labore": pro_labore_mes,
                "Outros custos": outros_mes,
                "Lucro operacional": lucro_op_mes,
                "Lucro líquido": lucro_liq_mes,
                "Margem operacional": safe_divide(lucro_op_mes, fat_mes, 100),
                "Margem líquida": safe_divide(lucro_liq_mes, fat_mes, 100),
                "Ticket médio": ticket_mes,
                "Crescimento vs mês anterior": crescimento_mes,
                "Crescimento vs ano anterior": crescimento_ano,
            }
        )

    mensal_df = pd.DataFrame(rows)

    return {
        "faturamento_total": faturamento_total,
        "receita_sistema": receita_sistema,
        "receita_export": receita_export,
        "meta_total": meta_total,
        "meta_mensal": float(inputs["meta_mensal"]),
        "meta_pct": meta_pct,
        "custos_fixos": custos_fixos,
        "custos_variaveis": custos_variaveis,
        "impostos": impostos,
        "pro_labore": pro_labore,
        "outros_custos": outros_custos,
        "custos_operacionais": custos_operacionais,
        "despesas_totais": despesas_totais,
        "lucro_operacional": lucro_operacional,
        "lucro_liquido": lucro_liquido,
        "margem_operacional": margem_operacional,
        "margem_liquida": margem_liquida,
        "ticket_medio": ticket_medio,
        "clientes_unicos": clientes_unicos,
        "crescimento_mensal": crescimento_mensal,
        "crescimento_anual": crescimento_anual,
        "ponto_equilibrio": despesas_totais,
        "diff_equilibrio": diff_equilibrio,
        "receita_setor": receita_setor,
        "participacao_setor": participacao_setor,
        "mensal_tendencia": mensal_tendencia,
        "mensal_df": mensal_df,
        "mes_ref": mes_ref,
        "mes_ref_nome": nome_mes(mes_ref) if mes_ref else "período",
        "mes_anterior": mes_anterior,
        "receita_mes_ref": receita_mes_ref,
        "receita_mes_anterior": receita_mes_anterior,
        "receita_ano_anterior": receita_aa,
        "observacoes": inputs.get("observacoes", ""),
        "n_meses": n_meses,
    }


# ============================================================
# COMPONENTES VISUAIS
# ============================================================
def metric_color(value, kind="positive"):
    """Define cor inteligente para cards."""
    try:
        value = float(value)
    except Exception:
        value = 0.0

    if kind == "meta":
        if value >= 100:
            return GREEN
        if value >= 80:
            return ORANGE
        return RED

    if kind == "margin":
        if value >= 15:
            return GREEN
        if value >= 0:
            return ORANGE
        return RED

    if kind == "growth":
        if value > 0:
            return GREEN
        if value < 0:
            return RED
        return ORANGE

    if kind == "break_even":
        return GREEN if value >= 0 else RED

    return GREEN if value >= 0 else RED


def render_finance_card(title, value, subtitle, color=PRIMARY):
    st.markdown(
        f"""
        <div class="finance-card" style="border-left-color:{color};">
            <h4>{title}</h4>
            <p class="value" style="color:{color};">{value}</p>
            <p class="sub">{subtitle}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def criar_cards_financeiros(ind):
    """Renderiza os cards principais."""
    c1, c2, c3, c4 = st.columns(4)

    with c1:
        render_finance_card(
            "💰 Faturamento Total",
            format_currency(ind["faturamento_total"]),
            f"Sistema: {format_currency(ind['receita_sistema'])}<br>Apoio: {format_currency(ind['receita_export'])}",
            GREEN if ind["faturamento_total"] > 0 else ORANGE,
        )

    with c2:
        render_finance_card(
            "🎯 Meta de Faturamento",
            format_currency(ind["meta_total"]),
            f"Meta mensal: {format_currency(ind['meta_mensal'])}<br>Período: {ind['n_meses']} mês(es)",
            PRIMARY,
        )

    with c3:
        cor_meta = metric_color(ind["meta_pct"], "meta")
        render_finance_card(
            "📊 % Meta Atingida",
            format_percent(ind["meta_pct"]),
            "Faturamento realizado sobre a meta definida",
            cor_meta,
        )
        st.progress(min(max(ind["meta_pct"] / 100, 0), 1))

    with c4:
        render_finance_card(
            "⚙️ Lucro Operacional",
            format_currency(ind["lucro_operacional"]),
            "Receita menos custos fixos, variáveis e outros custos",
            metric_color(ind["lucro_operacional"]),
        )

    c5, c6, c7, c8 = st.columns(4)

    with c5:
        render_finance_card(
            "🏁 Lucro Líquido",
            format_currency(ind["lucro_liquido"]),
            "O que sobra depois de impostos e pró-labore",
            metric_color(ind["lucro_liquido"]),
        )

    with c6:
        render_finance_card(
            "📈 Margem Operacional",
            format_percent(ind["margem_operacional"]),
            "Lucro operacional dividido pelo faturamento",
            metric_color(ind["margem_operacional"], "margin"),
        )

    with c7:
        render_finance_card(
            "💎 Margem Líquida",
            format_percent(ind["margem_liquida"]),
            "Lucro líquido dividido pelo faturamento",
            metric_color(ind["margem_liquida"], "margin"),
        )

    with c8:
        render_finance_card(
            "🧾 Ticket Médio",
            format_currency(ind["ticket_medio"]),
            f"Faturamento por cliente único<br>Clientes: {ind['clientes_unicos']}",
            GREEN if ind["ticket_medio"] > 0 else ORANGE,
        )

    c9, c10, c11 = st.columns(3)

    with c9:
        render_finance_card(
            "🔄 Crescimento Mensal",
            format_percent(ind["crescimento_mensal"]),
            f"{ind['mes_ref_nome']} vs mês anterior",
            metric_color(ind["crescimento_mensal"], "growth"),
        )

    with c10:
        render_finance_card(
            "🌱 Crescimento Anual",
            format_percent(ind["crescimento_anual"]),
            "Mês de referência vs mesmo mês do ano anterior",
            metric_color(ind["crescimento_anual"], "growth"),
        )

    with c11:
        status = "Acima do equilíbrio" if ind["diff_equilibrio"] >= 0 else "Abaixo do equilíbrio"
        render_finance_card(
            "⚖️ Ponto de Equilíbrio",
            format_currency(ind["ponto_equilibrio"]),
            f"{status}: {format_currency(abs(ind['diff_equilibrio']))}",
            metric_color(ind["diff_equilibrio"], "break_even"),
        )


def apply_plotly_layout(fig, title, height=360):
    """Aplica layout Plotly padrao do dashboard."""
    fig.update_layout(
        title=dict(text=title, font=dict(color="#E5E7EB", size=15, family="Segoe UI")),
        plot_bgcolor=CARD_BG,
        paper_bgcolor=CARD_BG,
        font=dict(color="#9CA3AF", size=12),
        height=height,
        margin=dict(l=45, r=30, t=55, b=45),
        legend=dict(font=dict(color="#9CA3AF", size=12), orientation="h", yanchor="bottom", y=-0.28),
    )
    fig.update_xaxes(showgrid=False, tickfont=dict(color="#9CA3AF"))
    fig.update_yaxes(showgrid=True, gridcolor="#2A2E3E", tickfont=dict(color="#9CA3AF"))
    return fig


def criar_graficos_financeiros(ind):
    """Renderiza os graficos financeiros obrigatorios."""
    g1, g2 = st.columns(2)

    with g1:
        fig_meta = go.Figure()
        realizado = ind["faturamento_total"]
        meta = ind["meta_total"]
        fig_meta.add_trace(
            go.Bar(
                x=["Faturamento realizado", "Meta"],
                y=[realizado, meta],
                marker_color=[GREEN if realizado >= meta and meta > 0 else ORANGE, "#2A2E3E"],
                text=[format_currency(realizado), format_currency(meta)],
                textposition="outside",
                textfont=dict(color=WHITE, size=12),
                cliponaxis=False,
            )
        )
        if meta > 0:
            fig_meta.add_hline(
                y=meta,
                line_color=RED,
                line_width=2,
                line_dash="dash",
                annotation_text="Meta",
                annotation_font_color=WHITE,
                annotation_bgcolor=RED,
            )
        fig_meta = apply_plotly_layout(fig_meta, "FATURAMENTO VS META", 350)
        st.plotly_chart(fig_meta, use_container_width=True)

    with g2:
        setor_df = ind["receita_setor"].reset_index()
        setor_df.columns = ["Setor", "Faturamento"]

        fig_setor = go.Figure()
        fig_setor.add_trace(
            go.Bar(
                x=setor_df["Setor"],
                y=setor_df["Faturamento"],
                marker_color=[GREEN if value > 0 else "#2A2E3E" for value in setor_df["Faturamento"]],
                text=[format_currency(value) for value in setor_df["Faturamento"]],
                textposition="outside",
                textfont=dict(color=WHITE, size=11),
                cliponaxis=False,
            )
        )
        fig_setor = apply_plotly_layout(fig_setor, "FATURAMENTO POR SETOR", 350)
        st.plotly_chart(fig_setor, use_container_width=True)

    g3, g4 = st.columns(2)

    with g3:
        part_df = ind["participacao_setor"].reset_index()
        part_df.columns = ["Setor", "Participacao"]
        values_for_pie = ind["receita_setor"].values

        if float(sum(values_for_pie)) <= 0:
            labels = SETOR_ORDER
            values = [1, 1, 1, 1, 1]
            textinfo = "label"
        else:
            labels = part_df["Setor"]
            values = values_for_pie
            textinfo = "label+percent"

        fig_part = go.Figure(
            data=[
                go.Pie(
                    labels=labels,
                    values=values,
                    hole=0.58,
                    marker_colors=[GREEN, ORANGE, "#36A2EB", "#A78BFA", "#6B7280"],
                    textinfo=textinfo,
                    textfont=dict(color=WHITE, size=12),
                )
            ]
        )
        fig_part.update_layout(
            title=dict(text="PARTICIPAÇÃO POR SETOR", font=dict(color="#E5E7EB", size=15)),
            plot_bgcolor=CARD_BG,
            paper_bgcolor=CARD_BG,
            font=dict(color="#9CA3AF"),
            height=360,
            margin=dict(l=20, r=20, t=55, b=35),
            legend=dict(font=dict(color="#9CA3AF", size=12), orientation="h", yanchor="bottom", y=-0.15),
        )
        st.plotly_chart(fig_part, use_container_width=True)

    with g4:
        labels = [
            "Receita",
            "Custos operacionais",
            "Impostos",
            "Pró-labore",
            "Lucro operacional",
            "Lucro líquido",
        ]
        values = [
            ind["faturamento_total"],
            ind["custos_operacionais"],
            ind["impostos"],
            ind["pro_labore"],
            ind["lucro_operacional"],
            ind["lucro_liquido"],
        ]
        colors = [
            GREEN,
            ORANGE,
            YELLOW,
            "#A78BFA",
            metric_color(ind["lucro_operacional"]),
            metric_color(ind["lucro_liquido"]),
        ]

        fig_comp = go.Figure()
        fig_comp.add_trace(
            go.Bar(
                x=labels,
                y=values,
                marker_color=colors,
                text=[format_currency(value) for value in values],
                textposition="outside",
                textfont=dict(color=WHITE, size=11),
                cliponaxis=False,
            )
        )
        fig_comp = apply_plotly_layout(fig_comp, "RECEITA, CUSTOS E LUCRO", 360)
        fig_comp.update_xaxes(tickangle=-18)
        st.plotly_chart(fig_comp, use_container_width=True)

    g5, g6 = st.columns(2)

    with g5:
        mensal_tendencia = ind["mensal_tendencia"]
        fig_cres = go.Figure()
        if not mensal_tendencia.empty:
            fig_cres.add_trace(
                go.Scatter(
                    x=mensal_tendencia["Mes"],
                    y=mensal_tendencia["Faturamento"],
                    mode="lines+markers+text",
                    line=dict(color=GREEN, width=3),
                    marker=dict(size=9, color=GREEN),
                    text=[format_currency(value) for value in mensal_tendencia["Faturamento"]],
                    textposition="top center",
                    textfont=dict(color=WHITE, size=10),
                )
            )
        else:
            fig_cres.add_annotation(text="Sem dados de faturamento mensal", showarrow=False, font=dict(color=WHITE))
        fig_cres = apply_plotly_layout(fig_cres, "CRESCIMENTO MENSAL", 350)
        st.plotly_chart(fig_cres, use_container_width=True)

    with g6:
        fig_eq = go.Figure()
        fig_eq.add_trace(
            go.Bar(
                x=["Faturamento", "Despesas totais"],
                y=[ind["faturamento_total"], ind["ponto_equilibrio"]],
                marker_color=[GREEN if ind["diff_equilibrio"] >= 0 else RED, ORANGE],
                text=[format_currency(ind["faturamento_total"]), format_currency(ind["ponto_equilibrio"])],
                textposition="outside",
                textfont=dict(color=WHITE, size=12),
                cliponaxis=False,
            )
        )
        fig_eq.add_hline(
            y=ind["ponto_equilibrio"],
            line_color=ORANGE,
            line_width=2,
            line_dash="dash",
            annotation_text="Equilíbrio",
            annotation_font_color=WHITE,
            annotation_bgcolor=ORANGE,
        )
        fig_eq = apply_plotly_layout(fig_eq, "PONTO DE EQUILÍBRIO", 350)
        st.plotly_chart(fig_eq, use_container_width=True)


def criar_tabela_mensal(ind):
    """Renderiza tabela mensal formatada."""
    mensal_df = ind["mensal_df"].copy()
    if mensal_df.empty:
        st.info("Sem dados mensais para a seleção atual.")
        return

    display = mensal_df.drop(columns=["Mes_Ano"], errors="ignore").copy()

    money_cols = [
        "Faturamento",
        "Meta",
        "Custos fixos",
        "Custos variáveis",
        "Impostos",
        "Pró-labore",
        "Outros custos",
        "Lucro operacional",
        "Lucro líquido",
        "Ticket médio",
    ]
    pct_cols = [
        "% Meta",
        "Margem operacional",
        "Margem líquida",
        "Crescimento vs mês anterior",
        "Crescimento vs ano anterior",
    ]

    for col in money_cols:
        if col in display.columns:
            display[col] = display[col].apply(format_currency)

    for col in pct_cols:
        if col in display.columns:
            display[col] = display[col].apply(format_percent)

    st.dataframe(display, use_container_width=True, hide_index=True)


def insight_html(text, color=ORANGE):
    st.markdown(
        f"""
        <div class="insight-box" style="border-left-color:{color};">
            <p>{text}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def criar_insights_financeiros(ind):
    """Cria insights automaticos financeiros."""
    st.markdown("### 💡 Insights Financeiros")

    setor_top = ind["receita_setor"].idxmax() if not ind["receita_setor"].empty else "Sem dados"
    setor_top_val = float(ind["receita_setor"].max()) if not ind["receita_setor"].empty else 0.0
    setor_top_part = safe_divide(setor_top_val, ind["faturamento_total"], 100)

    lucro_op_status = "positivo" if ind["lucro_operacional"] >= 0 else "negativo"
    lucro_liq_status = "sobrou dinheiro depois de tudo" if ind["lucro_liquido"] >= 0 else "não sobrou dinheiro depois de impostos e pró-labore"
    cresceu_caiu = "cresceu" if ind["crescimento_mensal"] >= 0 else "caiu"
    equilibrio_status = "acima" if ind["diff_equilibrio"] >= 0 else "abaixo"

    col_a, col_b = st.columns(2)

    with col_a:
        insight_html(
            f"🎯 A empresa atingiu <b>{format_percent(ind['meta_pct'])}</b> da meta de faturamento no período.",
            metric_color(ind["meta_pct"], "meta"),
        )
        insight_html(
            f"⚙️ O lucro operacional foi <b>{lucro_op_status}</b>: <b>{format_currency(ind['lucro_operacional'])}</b>.",
            metric_color(ind["lucro_operacional"]),
        )
        insight_html(
            f"🏁 O lucro líquido mostra que <b>{lucro_liq_status}</b>: <b>{format_currency(ind['lucro_liquido'])}</b>.",
            metric_color(ind["lucro_liquido"]),
        )

    with col_b:
        insight_html(
            f"🏷️ O setor com maior participação foi <b>{setor_top}</b>, com <b>{format_percent(setor_top_part)}</b> do faturamento.",
            GREEN if setor_top_val > 0 else ORANGE,
        )
        insight_html(
            f"💎 A margem líquida foi <b>{format_percent(ind['margem_liquida'])}</b>.",
            metric_color(ind["margem_liquida"], "margin"),
        )
        insight_html(
            f"🔄 O faturamento <b>{cresceu_caiu}</b> <b>{format_percent(abs(ind['crescimento_mensal']))}</b> vs mês anterior.",
            metric_color(ind["crescimento_mensal"], "growth"),
        )

    insight_html(
        f"⚖️ O faturamento está <b>{equilibrio_status}</b> do ponto de equilíbrio em <b>{format_currency(abs(ind['diff_equilibrio']))}</b>.",
        metric_color(ind["diff_equilibrio"], "break_even"),
    )

    if ind.get("observacoes"):
        insight_html(f"📝 Observações financeiras: <b>{ind['observacoes']}</b>", PRIMARY)


def render_template_apoio():
    """Mostra a estrutura esperada da planilha apoio_financeiro.xlsx."""
    exemplo = pd.DataFrame(
        [
            {
                "Mes_Ano": "2026-04",
                "Meta_Faturamento": 50000,
                "Custos_Fixos": 18000,
                "Custos_Variaveis": 7000,
                "Impostos": 3000,
                "Pro_Labore": 5000,
                "Outros_Custos": 1200,
                "Faturamento_Ano_Anterior": 42000,
                "Observacoes": "Exemplo de linha mensal",
            }
        ]
    )

    st.dataframe(exemplo, use_container_width=True, hide_index=True)

    csv_template = exemplo.head(0).to_csv(index=False).encode("utf-8")
    st.download_button(
        "Baixar CSV modelo da planilha de apoio",
        data=csv_template,
        file_name="apoio_financeiro_modelo.csv",
        mime="text/csv",
        use_container_width=True,
    )


# ============================================================
# EXECUCAO DA PAGINA
# ============================================================
df_raw, base_path, base_messages = load_data()
apoio_df, export_df, files_info, support_messages = load_financial_support()

if df_raw.empty:
    st.warning("Arquivo 'tabela geral.xlsx' não encontrado ou sem dados válidos. Coloque o arquivo na raiz do app ou na pasta data/.")
    st.stop()

# Opcoes de filtros
meses_disponiveis = sorted([m for m in df_raw["Mes_Ano"].dropna().unique() if str(m).strip()])

# Prioriza o mês mais recente com apoio financeiro. Assim a tela abre com custos preenchidos
# quando export (12).xlsx ou apoio_financeiro.xlsx trouxerem um mês específico.
meses_apoio_disponiveis = []
if not apoio_df.empty:
    meses_apoio_disponiveis = sorted([m for m in apoio_df["Mes_Ano"].dropna().unique() if str(m).strip()])

meses_export_despesas = []
if not export_df.empty:
    meses_export_despesas = sorted(
        [
            m
            for m in export_df.loc[export_df["Fluxo"] == "Despesa", "Mes_Ano"].dropna().unique()
            if str(m).strip()
        ]
    )

default_mes = None
for source_months in [meses_apoio_disponiveis, meses_export_despesas, meses_disponiveis]:
    candidates = [m for m in source_months if m in meses_disponiveis]
    if candidates:
        default_mes = candidates[-1]
        break
default_meses = [default_mes] if default_mes else []

setores_disponiveis = SETOR_ORDER
tipo_receita_sistema = sorted([x for x in df_raw["Tipo_Receita"].dropna().unique() if str(x).strip()])
tipo_receita_export = []
if not export_df.empty:
    tipo_receita_export = sorted(
        [
            x
            for x in export_df.loc[export_df["Fluxo"] == "Receita", "Tipo_Receita"].dropna().unique()
            if str(x).strip()
        ]
    )
tipos_receita_disponiveis = sorted(set(tipo_receita_sistema + tipo_receita_export))

profissionais_disponiveis = sorted([x for x in df_raw["Profissional"].dropna().unique() if str(x).strip()])
clientes_disponiveis = sorted([x for x in df_raw["Cliente"].dropna().unique() if str(x).strip()])

# Header
st.markdown(
    """
<div style="
    width:100%;
    background: linear-gradient(135deg, #00D4AA 0%, #00A884 100%);
    padding: 22px 26px;
    border-radius: 18px;
    margin-bottom: 18px;
    box-shadow: 0 6px 22px rgba(0,212,170,0.22);
    text-align: center;
">
    <div style="color:#0E1117; font-size:2.4rem; font-weight:900; letter-spacing:-0.5px; font-family:'Segoe UI', sans-serif;">
        🐾 Ravi's Pet Shop
    </div>
    <div style="color:#10231F; font-size:1rem; font-weight:800; margin-top:4px; font-family:'Segoe UI', sans-serif;">
        Dashboard Financeiro Executivo
    </div>
</div>
""",
    unsafe_allow_html=True,
)

h1, h2 = st.columns([3, 1])
with h1:
    st.markdown('<p class="header-title">FINANCEIRO</p>', unsafe_allow_html=True)
    st.markdown(
        '<p class="header-sub">Faturamento, crescimento, lucro operacional e lucro líquido em uma só tela</p>',
        unsafe_allow_html=True,
    )
with h2:
    min_d = df_raw.loc[df_raw["Mes_Ano"].isin(default_meses), "Data"].min() if default_meses else df_raw["Data"].min()
    max_d = df_raw.loc[df_raw["Mes_Ano"].isin(default_meses), "Data"].max() if default_meses else df_raw["Data"].max()
    if pd.notna(min_d) and pd.notna(max_d):
        st.markdown(
            f"""
            <div style="text-align:right; background:#1B1F2B; padding:12px 16px; border-radius:10px; margin-top:5px; box-shadow:0 2px 8px rgba(0,0,0,0.2);">
                <span style="color:#7B8194; font-size:0.75rem; font-weight:700;">BASE DE DADOS</span><br>
                <span style="color:#FFFFFF; font-weight:800; font-size:1rem;">{min_d.strftime('%d/%m/%Y')} - {max_d.strftime('%d/%m/%Y')}</span>
            </div>
            """,
            unsafe_allow_html=True,
        )

# Avisos amigaveis
if base_messages:
    with st.expander("Avisos da base principal", expanded=False):
        for msg in base_messages:
            st.warning(msg)

if support_messages:
    with st.expander("Avisos das planilhas de apoio", expanded=False):
        for msg in support_messages:
            st.warning(msg)

# Explicacao conceitual
d1, d2, d3 = st.columns(3)
with d1:
    st.markdown(
        """
        <div class="definition-card">
            <h4>Faturamento</h4>
            <p>Dinheiro que entrou no período selecionado. É receita, não lucro.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
with d2:
    st.markdown(
        """
        <div class="definition-card">
            <h4>Lucro Operacional</h4>
            <p>Resultado da operação antes de impostos e pró-labore.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
with d3:
    st.markdown(
        """
        <div class="definition-card">
            <h4>Lucro Líquido</h4>
            <p>O que realmente sobra depois de custos, impostos e pró-labore.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.markdown("<br>", unsafe_allow_html=True)

# Painel de configuracoes
if "mostrar_config_financeiro" not in st.session_state:
    st.session_state["mostrar_config_financeiro"] = True

def alternar_config_financeiro():
    st.session_state["mostrar_config_financeiro"] = not st.session_state["mostrar_config_financeiro"]

col_btn_seta, col_btn_config, col_status_config, col_espaco_config = st.columns([0.45, 1.8, 2.4, 4])

with col_btn_seta:
    st.button(
        "«" if st.session_state["mostrar_config_financeiro"] else "»",
        key="botao_seta_config_financeiro",
        help="Abrir ou fechar painel financeiro",
        use_container_width=True,
        on_click=alternar_config_financeiro,
    )

with col_btn_config:
    st.button(
        "Fechar configurações" if st.session_state["mostrar_config_financeiro"] else "Abrir configurações",
        key="botao_toggle_config_financeiro",
        use_container_width=True,
        on_click=alternar_config_financeiro,
    )

with col_status_config:
    origem_export = "com export" if files_info.get("export_path") else "sem export"
    origem_apoio = "com apoio" if files_info.get("apoio_path") else "sem apoio"
    st.caption(f"Configurações {'abertas' if st.session_state['mostrar_config_financeiro'] else 'fechadas'} | {origem_export} | {origem_apoio}")

# Valores iniciais seguros, mesmo com painel fechado
meses_sel = default_meses
setores_sel = setores_disponiveis
tipos_receita_sel = tipos_receita_disponiveis
prof_sel = []
cli_sel = []

defaults = calcular_defaults_financeiros(apoio_df, export_df, df_raw, meses_sel)

meta_mensal = defaults["meta_mensal"]
custos_fixos_mensais = defaults["custos_fixos_mensais"]
custos_variaveis_mensais = defaults["custos_variaveis_mensais"]
impostos_mensais = defaults["impostos_mensais"]
pro_labore_mensal = defaults["pro_labore_mensal"]
outros_custos_mensais = defaults["outros_custos_mensais"]
faturamento_ano_anterior = defaults["faturamento_ano_anterior"]
observacoes = defaults["observacoes"]

if st.session_state["mostrar_config_financeiro"]:
    st.markdown(
        """
        <div class="config-panel">
            <div class="title">💸 Painel Financeiro</div>
            <div class="subtitle">Escolha o período e revise as premissas financeiras usadas nos cálculos. Quando houver dados no export ou no apoio financeiro, o dashboard sugere valores iniciais, mas você pode editar tudo antes da análise.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    filtro1, filtro2, filtro3 = st.columns([1.15, 1, 1])
    with filtro1:
        st.markdown("#### 📅 Período")
        meses_sel = st.multiselect(
            "Mês/período",
            meses_disponiveis,
            default=default_meses,
            format_func=nome_mes,
            key="fin_meses_sel",
        )

    with filtro2:
        st.markdown("#### 🧭 Setores")
        setores_sel = st.multiselect(
            "Setor",
            setores_disponiveis,
            default=setores_disponiveis,
            key="fin_setores_sel",
        )

    with filtro3:
        st.markdown("#### 🏷️ Tipo de receita")
        tipos_receita_sel = st.multiselect(
            "Tipo de receita",
            tipos_receita_disponiveis,
            default=tipos_receita_disponiveis,
            key="fin_tipos_receita_sel",
        )

    filtro4, filtro5 = st.columns(2)
    with filtro4:
        prof_sel = st.multiselect(
            "Profissional",
            profissionais_disponiveis,
            default=[],
            key="fin_profissionais_sel",
        )

    with filtro5:
        cli_sel = st.multiselect(
            "Cliente",
            clientes_disponiveis,
            default=[],
            key="fin_clientes_sel",
        )

    defaults = calcular_defaults_financeiros(apoio_df, export_df, df_raw, meses_sel)

    st.markdown("#### ⚙️ Premissas financeiras do mês")
    st.markdown(
        f"""
        <div style="
            background: linear-gradient(135deg, #141824 0%, #1B1F2B 100%);
            border-radius: 12px;
            padding: 14px 18px;
            margin: 6px 0 14px 0;
            border-left: 4px solid #00D4AA;
            color: #D1D5DB;
            font-size: 0.92rem;
            line-height: 1.45;
            font-weight: 600;
        ">
            Informe aqui os valores de referência para <b>1 mês</b>.<br>
            Se selecionar vários meses, o dashboard replica esses valores para cada mês do período.
            Exemplo: custo fixo mensal de R$ 10.000 em 3 meses entra como R$ 30.000 no cálculo.<br>
            <span style="color:#9CA3AF;">Valores sugeridos inicialmente por: <b>{defaults['origem_defaults']}</b>. Todos os campos continuam editáveis.</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    inp1, inp2, inp3, inp4 = st.columns(4)
    with inp1:
        meta_mensal = st.number_input(
            "Meta de faturamento mensal",
            min_value=0.0,
            value=float(defaults["meta_mensal"]),
            step=500.0,
            format="%.2f",
            key="fin_meta_mensal",
        )
        custos_fixos_mensais = st.number_input(
            "Custos fixos mensais",
            min_value=0.0,
            value=float(defaults["custos_fixos_mensais"]),
            step=100.0,
            format="%.2f",
            key="fin_custos_fixos",
        )

    with inp2:
        custos_variaveis_mensais = st.number_input(
            "Custos variáveis mensais",
            min_value=0.0,
            value=float(defaults["custos_variaveis_mensais"]),
            step=100.0,
            format="%.2f",
            key="fin_custos_variaveis",
        )
        impostos_mensais = st.number_input(
            "Impostos mensais",
            min_value=0.0,
            value=float(defaults["impostos_mensais"]),
            step=100.0,
            format="%.2f",
            key="fin_impostos",
        )

    with inp3:
        pro_labore_mensal = st.number_input(
            "Pró-labore mensal",
            min_value=0.0,
            value=float(defaults["pro_labore_mensal"]),
            step=100.0,
            format="%.2f",
            key="fin_pro_labore",
        )
        outros_custos_mensais = st.number_input(
            "Outros custos mensais",
            min_value=0.0,
            value=float(defaults["outros_custos_mensais"]),
            step=100.0,
            format="%.2f",
            key="fin_outros_custos",
        )

    with inp4:
        faturamento_ano_anterior = st.number_input(
            "Faturamento do mesmo mês no ano anterior",
            min_value=0.0,
            value=float(defaults["faturamento_ano_anterior"]),
            step=500.0,
            format="%.2f",
            key="fin_faturamento_ano_anterior",
        )
        observacoes = st.text_area(
            "Observações financeiras opcionais",
            value=str(defaults.get("observacoes", "")),
            height=92,
            key="fin_observacoes",
        )

    with st.expander("Estrutura opcional de apoio_financeiro.xlsx", expanded=False):
        st.markdown(
            '<p class="small-muted">Crie este arquivo na raiz do app ou em data/ para alimentar metas, custos e histórico anual mês a mês.</p>',
            unsafe_allow_html=True,
        )
        render_template_apoio()

# Fallback quando filtros ficarem vazios
if not meses_sel:
    meses_sel = meses_disponiveis
if not setores_sel:
    setores_sel = setores_disponiveis
if not tipos_receita_sel:
    tipos_receita_sel = tipos_receita_disponiveis

# Aplicar filtros
df_f = aplicar_filtros_sistema(df_raw, meses_sel, setores_sel, tipos_receita_sel, prof_sel, cli_sel)
df_sem_filtro_mes = aplicar_filtros_sistema(df_raw, None, setores_sel, tipos_receita_sel, prof_sel, cli_sel)

export_receitas_f = aplicar_filtros_export_receitas(export_df, meses_sel, setores_sel, tipos_receita_sel, cli_sel)
export_receitas_sem_mes = aplicar_filtros_export_receitas(export_df, None, setores_sel, tipos_receita_sel, cli_sel)

inputs = {
    "meta_mensal": meta_mensal,
    "custos_fixos_mensais": custos_fixos_mensais,
    "custos_variaveis_mensais": custos_variaveis_mensais,
    "impostos_mensais": impostos_mensais,
    "pro_labore_mensal": pro_labore_mensal,
    "outros_custos_mensais": outros_custos_mensais,
    "faturamento_ano_anterior": faturamento_ano_anterior,
    "observacoes": observacoes,
}

indicadores = calcular_indicadores_financeiros(
    df_raw=df_raw,
    df_f=df_f,
    df_sem_filtro_mes=df_sem_filtro_mes,
    export_receitas_f=export_receitas_f,
    export_receitas_sem_mes=export_receitas_sem_mes,
    meses_sel=meses_sel,
    inputs=inputs,
)

# Periodo selecionado no topo
min_periodo = df_f["Data"].min() if not df_f.empty else pd.NaT
max_periodo = df_f["Data"].max() if not df_f.empty else pd.NaT
if pd.notna(min_periodo) and pd.notna(max_periodo):
    st.markdown(
        f"""
        <div style="text-align:right; background:#1B1F2B; padding:10px 14px; border-radius:10px; margin:4px 0 16px 0; box-shadow:0 2px 8px rgba(0,0,0,0.20);">
            <span style="color:#7B8194; font-size:0.75rem; font-weight:700;">PERÍODO SELECIONADO</span><br>
            <span style="color:#FFFFFF; font-weight:800; font-size:0.95rem;">{min_periodo.strftime('%d/%m/%Y')} - {max_periodo.strftime('%d/%m/%Y')}</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

# Cards
criar_cards_financeiros(indicadores)

st.markdown("<br>", unsafe_allow_html=True)

# Graficos
st.markdown("### 📊 Painel financeiro")
criar_graficos_financeiros(indicadores)

st.markdown("<br>", unsafe_allow_html=True)

# Tabela mensal
st.markdown("### 📋 Desempenho financeiro mensal")
criar_tabela_mensal(indicadores)

st.markdown("<br>", unsafe_allow_html=True)

# Detalhe de despesas vindas do export, se existir
if not export_df.empty:
    desp_filtradas = export_df[
        (export_df["Fluxo"] == "Despesa")
        & (export_df["Mes_Ano"].isin(meses_sel))
    ].copy()
    if not desp_filtradas.empty:
        with st.expander("Despesas detectadas no export (12).xlsx", expanded=False):
            resumo_desp = (
                desp_filtradas.groupby(["Bucket_Custo", "Categoria"], dropna=False)["Valor"]
                .sum()
                .reset_index()
                .sort_values("Valor", ascending=False)
            )
            resumo_desp["Valor"] = resumo_desp["Valor"].apply(format_currency)
            st.dataframe(resumo_desp, use_container_width=True, hide_index=True)

# Insights
criar_insights_financeiros(indicadores)

# Rodape
st.markdown("---")
fonte_export = files_info.get("export_path") if files_info.get("export_path") else "export opcional não encontrado"
fonte_apoio = files_info.get("apoio_path") if files_info.get("apoio_path") else "apoio_financeiro.xlsx opcional não encontrado"
st.markdown(
    f"""
<div style="text-align:center; color:#7B8194; font-size:0.8rem; font-weight:600;">
    Fonte principal: {base_path} | Apoio: {fonte_export} | {fonte_apoio} | Atualizado em {datetime.now().strftime('%d/%m/%Y %H:%M')}
</div>
""",
    unsafe_allow_html=True,
)
