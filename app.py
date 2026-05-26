#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Dashboard: Servicos / Banho e Tosa
Fonte de dados: tabela geral.xlsx
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import os, re
from datetime import datetime

# ============================================================
# CONFIGURACAO DA PAGINA
# ============================================================
st.set_page_config(
    page_title="Servicos | Banho e Tosa",
    page_icon="🐾",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# CSS CUSTOMIZADO
# ============================================================
st.markdown("""
<style>
    .stApp { background-color: #0E1117; }
    
    .kpi-card {
        background: linear-gradient(135deg, #1B1F2B 0%, #1E2235 100%);
        border-radius: 14px;
        padding: 20px 22px;
        border-left: 5px solid #00D4AA;
        margin-bottom: 12px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.3);
    }
    .kpi-card h4 {
        color: #9CA3AF;
        font-size: 0.85rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 1px;
        margin: 0 0 8px 0;
        font-family: 'Segoe UI', sans-serif;
    }
    .kpi-card .value {
        color: #00D4AA;
        font-size: 1.9rem;
        font-weight: 800;
        margin: 0;
        font-family: 'Segoe UI', sans-serif;
    }
    .kpi-card .sub {
        color: #7B8194;
        font-size: 0.8rem;
        margin-top: 6px;
        font-weight: 500;
    }
    
    .conv-card {
        background: linear-gradient(135deg, #1B1F2B 0%, #1E2235 100%);
        border-radius: 14px;
        padding: 22px;
        border-top: 4px solid #00D4AA;
        box-shadow: 0 4px 15px rgba(0,0,0,0.3);
    }
    .conv-card h4 {
        color: #00D4AA;
        font-size: 0.85rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 1px;
        margin: 0 0 10px 0;
    }
    .conv-card .pct {
        color: #00D4AA;
        font-size: 2.4rem;
        font-weight: 800;
    }
    .conv-card .sub {
        color: #7B8194;
        font-size: 0.8rem;
        font-weight: 500;
    }
    
    .ticket-card {
        background: linear-gradient(135deg, #1B1F2B 0%, #1E2235 100%);
        border-radius: 14px;
        padding: 18px 22px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.3);
    }
    .ticket-card h4 {
        color: #9CA3AF;
        font-size: 0.85rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 1px;
        margin: 0 0 8px 0;
    }
    .ticket-card .value {
        color: #FFFFFF;
        font-size: 1.7rem;
        font-weight: 800;
    }
    .ticket-card .sub {
        color: #7B8194;
        font-size: 0.78rem;
        font-weight: 500;
    }
    
    .insight-box {
        background: linear-gradient(135deg, #1B1F2B 0%, #1E2235 100%);
        border-radius: 12px;
        padding: 16px 20px;
        margin-bottom: 10px;
        border-left: 4px solid #FFA726;
        box-shadow: 0 2px 10px rgba(0,0,0,0.2);
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
        font-weight: 800;
        margin: 0;
        letter-spacing: -0.5px;
    }
    .header-sub {
        color: #7B8194;
        font-size: 1rem;
        font-weight: 500;
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
""", unsafe_allow_html=True)

# ============================================================
# FUNCOES
# ============================================================
MESES_PT = {
    'janeiro': 1, 'fevereiro': 2, 'março': 3, 'marco': 3,
    'abril': 4, 'maio': 5, 'junho': 6, 'julho': 7,
    'agosto': 8, 'setembro': 9, 'outubro': 10,
    'novembro': 11, 'dezembro': 12
}

MESES_NOME = {
    '01': 'janeiro',
    '02': 'fevereiro',
    '03': 'março',
    '04': 'abril',
    '05': 'maio',
    '06': 'junho',
    '07': 'julho',
    '08': 'agosto',
    '09': 'setembro',
    '10': 'outubro',
    '11': 'novembro',
    '12': 'dezembro'
}

def nome_mes(mes_ano):
    """
    Converte '2026-04' para 'abril'.
    """
    try:
        mes = str(mes_ano).split('-')[1]
        return MESES_NOME.get(mes, str(mes_ano))
    except:
        return str(mes_ano)

def parse_date_pt(text):
    """Converte 'quarta-feira, 1 de abril de 2026' para datetime."""
    if pd.isna(text):
        return pd.NaT
    try:
        match = re.search(r'(\d{1,2})\s+de\s+(\w+)\s+de\s+(\d{4})', str(text))
        if match:
            day = int(match.group(1))
            month_name = match.group(2).lower()
            year = int(match.group(3))
            month = MESES_PT.get(month_name)
            if month:
                return pd.Timestamp(year, month, day)
    except:
        pass
    return pd.NaT

def get_extra_type(service_name):
    if pd.isna(service_name):
        return 'Outro'
    s = str(service_name)
    if 'Hidratacao' in s or 'hidratacao' in s: return 'Hidratacao'
    if 'Escovacao' in s or 'CREME DENTAL' in s: return 'Escovacao de dentes'
    if 'Tosa' in s: return 'Tosa'
    if 'Desembolo' in s: return 'Desembolamento'
    if 'Clorexidina' in s: return 'Clorexidina'
    if 'Spa das patas' in s: return 'Spa das patas'
    if 'Arredondamento' in s: return 'Arredondamento'
    if 'Matizacao' in s: return 'Matizacao'
    if 'Remocao' in s: return 'Remocao de subpelo'
    if 'Corte de Unhas' in s: return 'Corte de unhas'
    return 'Outro'

# ============================================================
# CARREGAR DADOS
# ============================================================
@st.cache_data
def load_data():
    # Tenta encontrar o arquivo
    for path in ['tabela geral.xlsx', 'data/tabela geral.xlsx']:
        if os.path.exists(path):
            df = pd.read_excel(path, engine='openpyxl')
            break
    else:
        st.error("Arquivo 'tabela geral.xlsx' nao encontrado.")
        return pd.DataFrame()
    
    # Converter data
    df['Data'] = df['Data Realizacao'].apply(parse_date_pt)
    df['Mes_Ano'] = df['Data'].dt.to_period('M').astype(str)
    
    # Garantir numerico
    df['Valor Faturado'] = pd.to_numeric(df['Valor Faturado'], errors='coerce').fillna(0)
    
    return df

df_raw = load_data()
if df_raw.empty:
    st.stop()

# Filtrar apenas Banho e Servicos extras (excluir Outros)
df_serv = df_raw[df_raw['Classificação'].isin(['Banho', 'Serviços extras'])].copy()

# ============================================================
# CONFIGURACOES NO TOPO DO DASHBOARD
# ============================================================
# Nao usa mais st.sidebar, porque a sidebar nativa do Streamlit pode ser recolhida
# pelo navegador e desaparecer. Agora as configuracoes ficam dentro do proprio dash.

meses = sorted(df_serv['Mes_Ano'].dropna().unique())
profissionais = sorted(df_serv['Profissional'].dropna().unique())
clientes = sorted(df_serv['Cliente'].dropna().unique())

if 'mostrar_config' not in st.session_state:
    st.session_state['mostrar_config'] = False

def alternar_configuracoes():
    st.session_state['mostrar_config'] = not st.session_state['mostrar_config']

# Valores padrao quando as configuracoes estiverem fechadas
cap_diaria = 15
meta_banhos = 340
meta_extras = 300
meses_sel = meses
tipos = ['Avulso', 'Pacote']
prof_sel = []
cli_sel = []

# Botao sempre visivel, dentro do corpo principal do dashboard
col_btn_seta, col_btn_config, col_status_config, col_espaco_config = st.columns([0.45, 1.8, 2.4, 4])

with col_btn_seta:
    st.button(
        "«" if st.session_state['mostrar_config'] else "»",
        key="botao_seta_configuracoes",
        help="Abrir ou fechar painel de configuracoes",
        use_container_width=True,
        on_click=alternar_configuracoes
    )

with col_btn_config:
    st.button(
        "Fechar configurações" if st.session_state['mostrar_config'] else "Abrir configurações",
        key="botao_toggle_configuracoes",
        use_container_width=True,
        on_click=alternar_configuracoes
    )

with col_status_config:
    st.caption("Configurações abertas" if st.session_state['mostrar_config'] else "Configurações fechadas")

# Painel de configuracoes dentro do dashboard, sem depender da sidebar
if st.session_state['mostrar_config']:
    st.markdown("""
    <div style="
        background: linear-gradient(135deg, #1B1F2B 0%, #1E2235 100%);
        border-radius: 16px;
        padding: 18px 22px;
        margin: 10px 0 22px 0;
        border-left: 5px solid #00D4AA;
        box-shadow: 0 4px 15px rgba(0,0,0,0.28);
    ">
        <div style="color:#00D4AA; font-size:1.1rem; font-weight:800; margin-bottom:4px;">
            🐾 Painel de Controle
        </div>
        <div style="color:#9CA3AF; font-size:0.85rem; font-weight:500;">
            Ajuste metas, capacidade, período e filtros do dashboard.
        </div>
    </div>
    """, unsafe_allow_html=True)

    cfg1, cfg2, cfg3 = st.columns([1, 1, 1])

    with cfg1:
        st.markdown("#### ⚙️ Capacidade e Metas")
        cap_diaria = st.number_input("Capacidade diaria de banhos", min_value=1, value=15, step=1)
        meta_banhos = st.number_input("Meta mensal de banhos", min_value=0, value=340, step=10)
        meta_extras = st.number_input("Meta mensal de serviços extras", min_value=0, value=300, step=10)

    with cfg2:
        st.markdown("#### 📅 Período")
        meses_sel = st.multiselect("Meses", meses, default=meses)

    with cfg3:
        st.markdown("#### 🔍 Detalhamento")
        tipos = st.multiselect("Tipo de Banho", ['Avulso', 'Pacote'], default=['Avulso', 'Pacote'])
        prof_sel = st.multiselect("Profissional", profissionais, default=[])
        cli_sel = st.multiselect("Cliente", clientes, default=[])

    st.markdown("<br>", unsafe_allow_html=True)

# ============================================================
# APLICAR FILTROS
# ============================================================
df_f = df_serv.copy()
if meses_sel:
    df_f = df_f[df_f['Mes_Ano'].isin(meses_sel)]
if tipos:
    df_f = df_f[df_f['Tipo Banho'].isin(tipos)]
if prof_sel:
    df_f = df_f[df_f['Profissional'].isin(prof_sel)]
if cli_sel:
    df_f = df_f[df_f['Cliente'].isin(cli_sel)]

# ============================================================
# METRICAS
# ============================================================
banhos = df_f[df_f['Classificação'] == 'Banho']
extras = df_f[df_f['Classificação'] == 'Serviços extras']

total_banhos = len(banhos)
banhos_avulsos = len(banhos[banhos['Tipo Banho'] == 'Avulso'])
banhos_pacote = len(banhos[banhos['Tipo Banho'] == 'Pacote'])
total_extras = len(extras)
total_servicos = total_banhos + total_extras

fat_banhos = float(banhos['Valor Faturado'].sum())
fat_extras = float(extras['Valor Faturado'].sum())
fat_total = fat_banhos + fat_extras

ticket_geral = fat_total / total_servicos if total_servicos > 0 else 0

# Conversao por atendimento (Chave Atendimento)
atend = df_f.groupby(['Chave Atendimento', 'Tipo Banho']).agg(
    tem_banho=('Classificação', lambda x: (x == 'Banho').any()),
    tem_extra=('Classificação', lambda x: (x == 'Serviços extras').any()),
    valor_banho=('Valor Faturado', lambda x: x[df_f.loc[x.index, 'Classificação'] == 'Banho'].sum()),
    valor_extra=('Valor Faturado', lambda x: x[df_f.loc[x.index, 'Classificação'] == 'Serviços extras'].sum()),
).reset_index()

atend_b = atend[atend['tem_banho'] == True]
atend_av = atend_b[atend_b['Tipo Banho'] == 'Avulso']
atend_pk = atend_b[atend_b['Tipo Banho'] == 'Pacote']

n_av = len(atend_av)
n_av_ex = len(atend_av[atend_av['tem_extra'] == True])
conv_av = (n_av_ex / n_av * 100) if n_av > 0 else 0

n_pk = len(atend_pk)
n_pk_ex = len(atend_pk[atend_pk['tem_extra'] == True])
conv_pk = (n_pk_ex / n_pk * 100) if n_pk > 0 else 0

# Ticket com/sem extra
at_sem = atend_b[atend_b['tem_extra'] == False]
at_com = atend_b[atend_b['tem_extra'] == True]

tk_sem = float(at_sem['valor_banho'].sum() / len(at_sem)) if len(at_sem) > 0 else 0
tk_com = float((at_com['valor_banho'].sum() + at_com['valor_extra'].sum()) / len(at_com)) if len(at_com) > 0 else 0
impacto = ((tk_com - tk_sem) / tk_sem * 100) if tk_sem > 0 else 0

# Clientes em pacote
cli_total = df_f['Cliente'].nunique()
cli_pacote = df_f[df_f['Tipo Banho'] == 'Pacote']['Cliente'].nunique()
pct_pacote = (cli_pacote / cli_total * 100) if cli_total > 0 else 0

pct_extras_fat = (fat_extras / fat_total * 100) if fat_total > 0 else 0

# ============================================================
# CABECALHO
# ============================================================
st.markdown("""
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
    <div style="color:#10231F; font-size:1rem; font-weight:700; margin-top:4px; font-family:'Segoe UI', sans-serif;">
        Dashboard de Banho, Tosa e Serviços Extras
    </div>
</div>
""", unsafe_allow_html=True)

c_h1, c_h2 = st.columns([3, 1])
with c_h1:
    st.markdown('<p class="header-title">SERVIÇOS / BANHO E TOSA</p>', unsafe_allow_html=True)
    st.markdown('<p class="header-sub">Acompanhe o desempenho de banhos, tosas e servicos extras</p>', unsafe_allow_html=True)
with c_h2:
    min_d = df_f['Data'].min()
    max_d = df_f['Data'].max()
    if pd.notna(min_d) and pd.notna(max_d):
        st.markdown(f"""
        <div style="text-align:right; background:#1B1F2B; padding:12px 16px; border-radius:10px; margin-top:5px; box-shadow:0 2px 8px rgba(0,0,0,0.2);">
            <span style="color:#7B8194; font-size:0.75rem; font-weight:600;">PERIODO</span><br>
            <span style="color:#FFFFFF; font-weight:700; font-size:1rem;">{min_d.strftime('%d/%m/%Y')} - {max_d.strftime('%d/%m/%Y')}</span>
        </div>
        """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ============================================================
# ROW 1: KPIs
# ============================================================
c1, c2, c3, c4, c5 = st.columns(5)

with c1:
    st.markdown(f"""
    <div class="kpi-card">
        <h4>💰 Faturamento Total</h4>
        <p class="value">R$ {fat_total:,.2f}</p>
        <p class="sub">Banhos + Extras</p>
    </div>
    """, unsafe_allow_html=True)

with c2:
    st.markdown(f"""
    <div class="kpi-card">
        <h4>📋 Servicos Realizados</h4>
        <p class="value">{total_servicos:,}</p>
        <p class="sub">Banhos + Extras</p>
    </div>
    """, unsafe_allow_html=True)

with c3:
    st.markdown(f"""
    <div class="kpi-card">
        <h4>🛁 Total de Banhos</h4>
        <p class="value">{total_banhos:,}</p>
        <p class="sub">Avulsos: {banhos_avulsos} | Pacote: {banhos_pacote}</p>
    </div>
    """, unsafe_allow_html=True)

with c4:
    if total_extras > 0:
        tk_ex = fat_extras / total_extras
        st.markdown(f"""
        <div class="kpi-card">
            <h4>⭐ Servicos Extras</h4>
            <p class="value">{total_extras:,}</p>
            <p class="sub">Ticket medio: R$ {tk_ex:,.2f}</p>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <div class="kpi-card">
            <h4>⭐ Servicos Extras</h4>
            <p class="value">0</p>
            <p class="sub">Sem extras no periodo</p>
        </div>
        """, unsafe_allow_html=True)

with c5:
    st.markdown(f"""
    <div class="kpi-card">
        <h4>🎯 Ticket Medio Geral</h4>
        <p class="value">R$ {ticket_geral:,.2f}</p>
        <p class="sub">Por servico realizado</p>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ============================================================
# ROW 2: CONVERSAO + GRAFICO
# ============================================================
cc1, cc2, cc3 = st.columns([1, 1, 2])

with cc1:
    st.markdown(f"""
    <div class="conv-card">
        <h4>CONVERSAO AVULSO → EXTRA</h4>
        <p class="pct">{conv_av:.1f}%</p>
        <p class="sub">{n_av_ex} de {n_av} atendimentos avulsos<br>com pelo menos 1 extra</p>
    </div>
    """, unsafe_allow_html=True)

with cc2:
    st.markdown(f"""
    <div class="conv-card">
        <h4>CONVERSAO PACOTE → EXTRA</h4>
        <p class="pct">{conv_pk:.1f}%</p>
        <p class="sub">{n_pk_ex} de {n_pk} atendimentos pacote<br>com pelo menos 1 extra</p>
    </div>
    """, unsafe_allow_html=True)

with cc3:
    ext_mes = extras.groupby('Mes_Ano').size().reset_index(name='Qtd').sort_values('Mes_Ano')
    ext_mes['Mes_Nome'] = ext_mes['Mes_Ano'].apply(nome_mes)

    max_ext = ext_mes['Qtd'].max() if not ext_mes.empty else 0
    y_max_ext = max(max_ext, meta_extras) * 1.25 if max(max_ext, meta_extras) > 0 else 10

    fig_ext = go.Figure()

    fig_ext.add_trace(go.Bar(
        x=ext_mes['Mes_Nome'],
        y=ext_mes['Qtd'],
        marker_color=[
            '#00D4AA' if qtd >= meta_extras else '#B00000'
            for qtd in ext_mes['Qtd']
        ],
        text=ext_mes['Qtd'],
        textposition='outside',
        textfont=dict(
            color='#FFFFFF',
            size=14,
            family='Arial Black'
        ),
        cliponaxis=False
    ))

    fig_ext.add_hline(
        y=meta_extras,
        line_color='#FF0000',
        line_width=3,
        annotation_text=f"Meta: {meta_extras}",
        annotation_position="top right",
        annotation_font_color="#FFFFFF",
        annotation_bgcolor="#B00000"
    )

    fig_ext.update_layout(
        title=dict(
            text="SERVIÇOS EXTRAS / MÊS",
            font=dict(color='#E5E7EB', size=15, family='Segoe UI')
        ),
        plot_bgcolor='#1B1F2B',
        paper_bgcolor='#1B1F2B',
        font=dict(color='#9CA3AF', size=12),
        height=300,
        xaxis=dict(
            showgrid=False,
            tickfont=dict(size=12),
            type='category'
        ),
        yaxis=dict(
            showgrid=True,
            gridcolor='#2A2E3E',
            range=[0, y_max_ext]
        ),
        margin=dict(l=40, r=30, t=55, b=35),
        uniformtext_minsize=12,
        uniformtext_mode='show'
    )

    st.plotly_chart(fig_ext, use_container_width=True)

# ============================================================
# ROW 3: TICKET MEDIO
# ============================================================
ct1, ct2, ct3 = st.columns([1, 1, 1])

with ct1:
    st.markdown(f"""
    <div class="ticket-card">
        <h4>Ticket Medio SEM Extra</h4>
        <p class="value">R$ {tk_sem:,.2f}</p>
        <p class="sub">Atendimentos sem venda adicional</p>
    </div>
    """, unsafe_allow_html=True)

with ct2:
    cor_imp = '#00D4AA' if impacto > 0 else '#FF5252'
    st.markdown(f"""
    <div class="ticket-card" style="text-align:center; border-top:3px solid {cor_imp};">
        <h4>Impacto dos Extras</h4>
        <p class="value" style="color:{cor_imp};">+{impacto:.1f}%</p>
        <p class="sub">Ganho com venda de extras</p>
    </div>
    """, unsafe_allow_html=True)

with ct3:
    st.markdown(f"""
    <div class="ticket-card">
        <h4>Ticket Medio COM Extra</h4>
        <p class="value">R$ {tk_com:,.2f}</p>
        <p class="sub">Atendimentos com extras</p>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ============================================================
# ROW 4: BANHOS/MES + DONUT + RESUMO
# ============================================================
cg1, cg2, cg3 = st.columns([1, 1, 1])

with cg1:
    b_mes = banhos.groupby('Mes_Ano').size().reset_index(name='Qtd').sort_values('Mes_Ano')
    b_mes['Mes_Nome'] = b_mes['Mes_Ano'].apply(nome_mes)

    max_banhos = b_mes['Qtd'].max() if not b_mes.empty else 0
    y_max_banhos = max(max_banhos, meta_banhos) * 1.25 if max(max_banhos, meta_banhos) > 0 else 10

    fig_b = go.Figure()

    fig_b.add_trace(go.Bar(
        x=b_mes['Mes_Nome'],
        y=b_mes['Qtd'],
        marker_color=[
            '#00D4AA' if qtd >= meta_banhos else '#B00000'
            for qtd in b_mes['Qtd']
        ],
        text=b_mes['Qtd'],
        textposition='outside',
        textfont=dict(
            color='#FFFFFF',
            size=14,
            family='Arial Black'
        ),
        cliponaxis=False
    ))

    fig_b.add_hline(
        y=meta_banhos,
        line_color='#FF0000',
        line_width=3,
        annotation_text=f"Meta: {meta_banhos}",
        annotation_position="top right",
        annotation_font_color="#FFFFFF",
        annotation_bgcolor="#B00000"
    )

    fig_b.update_layout(
        title=dict(
            text="BANHOS / MÊS",
            font=dict(color='#E5E7EB', size=15)
        ),
        plot_bgcolor='#1B1F2B',
        paper_bgcolor='#1B1F2B',
        font=dict(color='#9CA3AF', size=12),
        height=300,
        xaxis=dict(
            showgrid=False,
            tickfont=dict(size=12),
            type='category'
        ),
        yaxis=dict(
            showgrid=True,
            gridcolor='#2A2E3E',
            range=[0, y_max_banhos]
        ),
        margin=dict(l=40, r=30, t=55, b=35),
        uniformtext_minsize=12,
        uniformtext_mode='show'
    )

    st.plotly_chart(fig_b, use_container_width=True)

with cg2:
    fig_dn = go.Figure(data=[go.Pie(
        labels=['Banhos', 'Servicos extras'],
        values=[total_banhos, total_extras],
        hole=0.55,
        marker_colors=['#FF6B35', '#00D4AA'],
        textinfo='label+percent',
        textfont=dict(size=13, color='white')
    )])
    fig_dn.update_layout(
        title=dict(text="BANHO X EXTRAS", font=dict(color='#E5E7EB', size=15)),
        plot_bgcolor='#1B1F2B', paper_bgcolor='#1B1F2B',
        font=dict(color='#9CA3AF'), height=300,
        showlegend=True,
        legend=dict(font=dict(color='#9CA3AF', size=12)),
        margin=dict(l=20, r=20, t=45, b=20)
    )
    st.plotly_chart(fig_dn, use_container_width=True)

with cg3:
    st.markdown("#### 📊 RESUMO")
    resumo = pd.DataFrame({
        'Indicador': [
            'Faturamento', 'Servicos', 'Banhos',
            'Extras', 'Ticket medio', 'Conv. avulso', 'Conv. pacote'
        ],
        'Valor': [
            f"R$ {fat_total:,.2f}", f"{total_servicos:,}", f"{total_banhos:,}",
            f"{total_extras:,}", f"R$ {ticket_geral:,.2f}",
            f"{conv_av:.1f}%", f"{conv_pk:.1f}%"
        ]
    })
    st.dataframe(resumo, use_container_width=True, hide_index=True)

st.markdown("<br>", unsafe_allow_html=True)

# ============================================================
# ROW 5: RANKING + OCUPACAO
# ============================================================
cr1, cr2 = st.columns(2)

with cr1:
    ext_df = extras.copy()
    ext_df['Tipo_Extra'] = ext_df['Servico/Produto'].apply(get_extra_type)
    ranking = ext_df.groupby('Tipo_Extra').agg(
        Quantidade=('Tipo_Extra', 'count'),
        Faturamento=('Valor Faturado', 'sum')
    ).sort_values('Quantidade', ascending=True).reset_index()
    
    fig_rk = go.Figure()
    fig_rk.add_trace(go.Bar(
        y=ranking['Tipo_Extra'], x=ranking['Quantidade'],
        orientation='h', marker_color='#FFA726',
        text=ranking['Quantidade'], textposition='outside',
        textfont=dict(color='#E5E7EB', size=12)
    ))
    fig_rk.update_layout(
        title=dict(text="RANKING DE EXTRAS", font=dict(color='#E5E7EB', size=15)),
        plot_bgcolor='#1B1F2B', paper_bgcolor='#1B1F2B',
        font=dict(color='#9CA3AF', size=12), height=380,
        xaxis=dict(showgrid=True, gridcolor='#2A2E3E'),
        yaxis=dict(showgrid=False, tickfont=dict(size=12)),
        margin=dict(l=150, r=50, t=45, b=30)
    )
    st.plotly_chart(fig_rk, use_container_width=True)

with cr2:
    banhos_dia = banhos.copy()
    banhos_dia['Dia'] = banhos_dia['Data'].dt.date
    ocup_list = []
    for mes in sorted(df_f['Mes_Ano'].dropna().unique()):
        dm = banhos_dia[banhos_dia['Mes_Ano'] == mes]
        dias = dm['Dia'].nunique()
        total = len(dm)
        cap = dias * cap_diaria
        taxa = (total / cap * 100) if cap > 0 else 0
        ocup_list.append({'Mes': nome_mes(mes), 'Taxa': taxa, 'Real': total, 'Cap': cap})
    ocup_df = pd.DataFrame(ocup_list)
    
    if not ocup_df.empty:
        fig_oc = go.Figure()
        fig_oc.add_trace(go.Bar(
            x=ocup_df['Mes'], y=ocup_df['Taxa'],
            marker_color=['#00D4AA' if t < 80 else '#FFA726' if t < 95 else '#FF5252' for t in ocup_df['Taxa']],
            text=[f"{t:.0f}%" for t in ocup_df['Taxa']], textposition='outside',
            textfont=dict(color='#E5E7EB', size=13),
            cliponaxis=False
        ))
        fig_oc.update_layout(
            title=dict(text=f"TAXA DE OCUPACAO (Cap: {cap_diaria}/dia)", font=dict(color='#E5E7EB', size=15)),
            plot_bgcolor='#1B1F2B', paper_bgcolor='#1B1F2B',
            font=dict(color='#9CA3AF', size=12), height=380,
            xaxis=dict(showgrid=False, tickfont=dict(size=12), type='category'),
            yaxis=dict(showgrid=True, gridcolor='#2A2E3E', range=[0, 130]),
            margin=dict(l=40, r=20, t=45, b=30)
        )
        fig_oc.add_hline(y=100, line_dash="dash", line_color="#FF5252", annotation_text="100%",
                         annotation_font_color="#FF5252")
        st.plotly_chart(fig_oc, use_container_width=True)

st.markdown("<br>", unsafe_allow_html=True)

# ============================================================
# ROW 6: TABELA MENSAL
# ============================================================
st.markdown("### 📊 DESEMPENHO MENSAL")
rows = []
for mes in sorted(df_f['Mes_Ano'].dropna().unique()):
    dm = df_f[df_f['Mes_Ano'] == mes]
    ba = len(dm[(dm['Classificação'] == 'Banho') & (dm['Tipo Banho'] == 'Avulso')])
    bp = len(dm[(dm['Classificação'] == 'Banho') & (dm['Tipo Banho'] == 'Pacote')])
    tb = ba + bp
    ne = len(dm[dm['Classificação'] == 'Serviços extras'])
    ts = tb + ne
    ft = float(dm['Valor Faturado'].sum())
    tk = ft / ts if ts > 0 else 0
    pe = (ne / tb * 100) if tb > 0 else 0
    rows.append({
        'Mes': nome_mes(mes), 'B. Avulsos': ba, 'B. Pacote': bp,
        'Total Banhos': tb, 'Extras': ne, 'Total Srv': ts,
        'Faturamento': f"R$ {ft:,.2f}", 'Ticket': f"R$ {tk:,.2f}",
        '% Extras/Banhos': f"{pe:.1f}%"
    })
if rows:
    tba = sum(r['B. Avulsos'] for r in rows)
    tbp = sum(r['B. Pacote'] for r in rows)
    ttb = sum(r['Total Banhos'] for r in rows)
    tex = sum(r['Extras'] for r in rows)
    tts = sum(r['Total Srv'] for r in rows)
    ttk = fat_total / tts if tts > 0 else 0
    tpe = (tex / ttb * 100) if ttb > 0 else 0
    rows.append({
        'Mes': '📌 TOTAL', 'B. Avulsos': tba, 'B. Pacote': tbp,
        'Total Banhos': ttb, 'Extras': tex, 'Total Srv': tts,
        'Faturamento': f"R$ {fat_total:,.2f}", 'Ticket': f"R$ {ttk:,.2f}",
        '% Extras/Banhos': f"{tpe:.1f}%"
    })
st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

st.markdown("<br>", unsafe_allow_html=True)

# ============================================================
# ROW 7: INSIGHTS
# ============================================================
st.markdown("### 💡 INSIGHTS AUTOMATICOS")

ci1, ci2 = st.columns(2)
with ci1:
    st.markdown(f"""
    <div class="insight-box">
        <p>💰 Extras representam <b>{pct_extras_fat:.1f}%</b> do faturamento total.</p>
    </div>
    """, unsafe_allow_html=True)
    
    diff_tk = tk_com - tk_sem
    st.markdown(f"""
    <div class="insight-box">
        <p>📈 O ticket medio COM extra e <b>R$ {diff_tk:,.2f}</b> maior que SEM extra.</p>
    </div>
    """, unsafe_allow_html=True)

with ci2:
    st.markdown(f"""
    <div class="insight-box">
        <p>🔄 Conversao de extras em banho avulso: <b>{conv_av:.1f}%</b>.</p>
    </div>
    """, unsafe_allow_html=True)
    
    if not ocup_df.empty:
        ocup_med = ocup_df['Taxa'].mean()
        st.markdown(f"""
        <div class="insight-box">
            <p>📊 Taxa de ocupacao media do periodo: <b>{ocup_med:.1f}%</b>.</p>
        </div>
        """, unsafe_allow_html=True)

st.markdown(f"""
<div class="insight-box">
    <p>🏷️ <b>{pct_pacote:.1f}%</b> dos clientes possuem atendimentos vinculados a pacote.</p>
</div>
""", unsafe_allow_html=True)

# ============================================================
# RODAPE
# ============================================================
st.markdown("---")
st.markdown(f"""
<div style="text-align:center; color:#7B8194; font-size:0.8rem; font-weight:500;">
    Fonte: tabela geral.xlsx | Atualizado em {datetime.now().strftime('%d/%m/%Y %H:%M')}
</div>
""", unsafe_allow_html=True)
