
import os
import re
import sqlite3
import hashlib
import base64
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

try:
    import plotly.express as px
except Exception:
    px = None

try:
    from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
except Exception:
    AgGrid = None
    GridOptionsBuilder = None
    GridUpdateMode = None

APP_TITLE = "Controle de Medidas Preventivas - ANTT Itapema"
BASE_DIR = Path(__file__).resolve().parent
LOGO_ANTT_PATH = BASE_DIR / "assets" / "antt_logo.svg"
DB_PATH = "medidas_antt.db"
EXCEL_PATH = "Controle_Medidas_Preventivas_Itapema_BI_ID_Automatico_v2.xlsx"
UPLOAD_DIR = Path("uploads")
UPLOAD_DIR.mkdir(exist_ok=True)

CONCESSIONARIAS = ["Litoral Sul", "Planalto Sul", "ViaCosteira"]
DISCIPLINAS = [
    "Pavimento", "Sinalização Horizontal", "Sinalização Vertical",
    "Elementos de Proteção e Segurança", "Drenagem e OAC", "OAEs",
    "Sistema Elétrico e Iluminação", "Terraplenos e Estruturas de Contenção",
    "Faixa de Domínio", "Túneis", "Edificações", "Obras Obrigatórias", "Operação",
]
STATUS_API = ["Rascunho", "Emitida", "Comunicada"]
STATUS_ANC = ["Não cumprido", "Cumprido", "Convertido em medida sancionatória", "Arquivada"]
STATUS_AE = ["Aguardando acompanhamento", "Em acompanhamento", "Finalizada", "Vencida"]
SIM_NAO = ["Não", "Sim"]
STATUS_VALIDACAO_ANC = ["Aceita", "Rejeitada", "Complementação solicitada"]
STATUS_VALIDACAO_AE = ["Aceito", "Rejeitado", "Complementação solicitada"]
EXTENSOES_UPLOAD = ["pdf", "html", "htm"]
REGEX_DOCUMENTO_SEI = r"^\d+$"

NOME_API = "Alerta de Potencial Inconformidade (API)"
NOME_ANC = "Aviso de Não Conformidade (ANC)"
NOME_AE = "Ação Educativa (AE)"
MAPA_TIPO_PAGINA = {NOME_API: NOME_API, NOME_ANC: NOME_ANC, NOME_AE: NOME_AE}


POWERBI_COLORS = {
    "bg": "#f4f8f5",
    "panel": "#ffffff",
    "ink": "#153125",
    "muted": "#5d7168",
    "border": "#d7e5dc",
    "accent": "#f6b728",
    "accent_dark": "#996b00",
    "blue": "#0057a6",
    "green": "#00843d",
    "red": "#c92a2a",
    "orange": "#e68a00",
    "gray": "#6b7280",
}

TYPE_COLORS = {
    NOME_API: "#0057a6",
    NOME_ANC: "#f6b728",
    NOME_AE: "#00843d",
}

SHORT_TYPE_COLORS = {
    "API": "#0057a6",
    "ANC": "#f6b728",
    "AE": "#00843d",
}

STATUS_COLORS = {
    "Vencida": "#dc2626",
    "A vencer em 15 dias": "#ea8a00",
    "A vencer em 30 dias": "#ea8a00",
    "Dentro do prazo": "#0057a6",
    "Cumprido": "#00843d",
    "Finalizada": "#00843d",
    "Sem prazo informado": "#6b7280",
}

PLOTLY_TEMPLATE = "plotly_white"


def aplicar_estilo_powerbi():
    st.markdown(
        f"""
        <style>
            :root {{
                --antt-bg: {POWERBI_COLORS['bg']};
                --antt-panel: {POWERBI_COLORS['panel']};
                --antt-ink: {POWERBI_COLORS['ink']};
                --antt-muted: {POWERBI_COLORS['muted']};
                --antt-border: {POWERBI_COLORS['border']};
                --antt-accent: {POWERBI_COLORS['accent']};
                --antt-accent-dark: {POWERBI_COLORS['accent_dark']};
                --antt-blue: {POWERBI_COLORS['blue']};
                --antt-green: {POWERBI_COLORS['green']};
                --antt-red: {POWERBI_COLORS['red']};
                --antt-orange: {POWERBI_COLORS['orange']};
            }}
            .stApp {{ background: var(--antt-bg); color: var(--antt-ink); }}
            .block-container {{ padding-top: 1.25rem; padding-bottom: 2rem; max-width: 1480px; }}
            h1, h2, h3 {{ letter-spacing: 0; color: var(--antt-ink); }}
            h1 {{ font-size: 1.8rem; font-weight: 750; margin-bottom: .25rem; }}
            h2, h3 {{ font-weight: 700; }}

            [data-testid="stSidebar"] {{
                background: #ffffff;
                border-right: 1px solid var(--antt-border);
                box-shadow: 10px 0 28px rgba(0, 132, 61, 0.08);
            }}
            [data-testid="stSidebar"] > div:first-child {{ padding-top: 1rem; }}
            [data-testid="stSidebar"] * {{ color: var(--antt-ink); }}
            [data-testid="stSidebar"] label,
            [data-testid="stSidebar"] .stMarkdown p,
            [data-testid="stSidebar"] .stCaptionContainer {{ color: var(--antt-muted); }}
            [data-testid="stSidebar"] [role="radiogroup"] label {{
                min-height: 40px;
                border-radius: 8px;
                padding: 8px 10px;
                margin-bottom: 5px;
                border: 1px solid transparent;
                transition: background .15s ease, border-color .15s ease, transform .15s ease;
            }}
            [data-testid="stSidebar"] [role="radiogroup"] label:hover {{
                background: #eef8f1;
                border-color: #cde8d5;
                transform: translateX(2px);
            }}
            [data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked) {{
                background: #e4f4e9 !important;
                border-color: rgba(0, 132, 61, 0.35);
                box-shadow: inset 4px 0 0 var(--antt-green);
                font-weight: 800;
            }}

            .antt-sidebar-brand {{
                background: linear-gradient(180deg, #ffffff 0%, #f2faf5 100%);
                border: 1px solid #cfe8d7;
                border-radius: 10px;
                padding: 14px;
                margin: 0 0 12px;
                text-align: center;
            }}
            .antt-logo-img {{ display: block; height: auto; max-width: 100%; margin: 0 auto; }}
            .antt-sidebar-logo .antt-logo-img {{ width: min(100%, 224px); margin-bottom: 10px; }}
            .antt-sidebar-brand-title {{ color: var(--antt-green); font-size: 1.02rem; font-weight: 900; line-height: 1.1; }}
            .antt-sidebar-brand-subtitle {{ color: var(--antt-muted); font-size: .78rem; margin-top: 6px; line-height: 1.35; }}
            .antt-sidebar-profile {{
                background: #f8fbf9;
                border: 1px solid var(--antt-border);
                border-radius: 8px;
                padding: 10px 12px;
                margin: 0 0 12px;
            }}
            .antt-sidebar-profile-label {{ color: var(--antt-green); font-size: .68rem; font-weight: 850; text-transform: uppercase; letter-spacing: .06em; margin-top: 8px; }}
            .antt-sidebar-profile-value {{ color: var(--antt-ink); font-size: .82rem; font-weight: 650; word-break: break-word; }}

            div[data-testid="stMetric"],
            div[data-testid="stDataFrame"],
            div[data-testid="stExpander"],
            div[data-testid="stForm"] {{
                background: var(--antt-panel);
                border: 1px solid var(--antt-border);
                border-radius: 8px;
                box-shadow: 0 10px 24px rgba(21, 49, 37, 0.06);
            }}
            div[data-testid="stMetric"] {{ padding: 14px 16px; min-height: 112px; }}
            div[data-testid="stMetricLabel"] p {{ color: var(--antt-muted); font-size: .82rem; font-weight: 650; }}
            div[data-testid="stMetricValue"] {{ color: var(--antt-ink); font-weight: 800; }}
            div[data-testid="stForm"] {{ padding: 18px; }}
            div[data-testid="stDataFrame"] {{ overflow: hidden; }}

            .antt-page-header {{
                display: flex;
                justify-content: space-between;
                gap: 18px;
                align-items: flex-end;
                padding: 4px 0 12px 0;
                border-bottom: 1px solid var(--antt-border);
                margin-bottom: 16px;
            }}
            .antt-eyebrow {{ color: var(--antt-green); font-weight: 900; text-transform: uppercase; font-size: .72rem; letter-spacing: .08em; margin-bottom: 2px; }}
            .antt-title {{ font-size: 1.85rem; font-weight: 850; color: var(--antt-ink); line-height: 1.15; }}
            .antt-subtitle {{ color: var(--antt-muted); font-size: .94rem; margin-top: 4px; }}
            .antt-chip-row {{ display: flex; gap: 8px; flex-wrap: wrap; justify-content: flex-end; }}
            .antt-chip {{
                background: #fff7dc;
                color: #725000;
                border: 1px solid #f4d073;
                border-radius: 999px;
                padding: 6px 10px;
                font-size: .78rem;
                font-weight: 750;
                white-space: nowrap;
            }}
            .antt-kpi {{
                background: var(--antt-panel);
                border: 1px solid var(--antt-border);
                border-radius: 8px;
                box-shadow: 0 10px 24px rgba(21, 49, 37, 0.06);
                padding: 15px 16px 13px;
                min-height: 118px;
                position: relative;
                overflow: hidden;
            }}
            .antt-kpi::before {{ content: ""; position: absolute; left: 0; top: 0; bottom: 0; width: 5px; background: var(--kpi-color, var(--antt-green)); }}
            .antt-kpi-label {{ color: var(--antt-muted); font-size: .78rem; font-weight: 750; text-transform: uppercase; letter-spacing: .04em; margin-bottom: 8px; }}
            .antt-kpi-value {{ color: var(--antt-ink); font-size: 2rem; line-height: 1; font-weight: 850; }}
            .antt-kpi-help {{ color: var(--antt-muted); font-size: .80rem; margin-top: 9px; }}
            .antt-alert-card {{
                background: #fff7ed;
                border: 1px solid #fed7aa;
                border-left: 5px solid var(--antt-orange);
                color: #7c2d12;
                border-radius: 8px;
                padding: 13px 15px;
                margin: 4px 0 14px;
                font-weight: 650;
            }}
            .antt-filter-panel {{
                background: #ffffff;
                border: 1px solid var(--antt-border);
                border-radius: 8px;
                padding: 12px 12px 14px;
                box-shadow: 0 10px 24px rgba(21, 49, 37, 0.06);
                margin-bottom: 12px;
            }}
            .st-key-dash_right_filters {{
                position: fixed;
                top: 0;
                right: 0;
                bottom: 0;
                width: 252px;
                overflow-y: auto;
                z-index: 999;
                background: #ffffff;
                border-left: 1px solid var(--antt-border);
                box-shadow: -10px 0 28px rgba(0, 132, 61, 0.08);
                padding: 14px 12px 22px;
            }}
            .st-key-dash_right_filters:has(.antt-filter-collapsed) {{
                width: 84px;
                padding: 12px 8px;
                overflow: hidden;
            }}
            .st-key-dash_right_filters [data-testid="stVerticalBlock"] {{ gap: .55rem; }}
            .st-key-dash_right_filters .stButton > button {{
                min-height: 34px;
                padding: 6px 10px;
                font-size: .82rem;
                border-radius: 6px;
            }}
            .block-container:has(.st-key-dash_right_filters) {{
                padding-right: 276px;
            }}
            .block-container:has(.st-key-dash_right_filters:has(.antt-filter-collapsed)) {{
                padding-right: 108px;
            }}
            .antt-filter-title {{ color: var(--antt-green); font-size: .88rem; font-weight: 900; text-transform: uppercase; letter-spacing: .06em; margin-bottom: 4px; }}
            .antt-filter-help {{ color: var(--antt-muted); font-size: .78rem; line-height: 1.3; margin-bottom: 8px; }}
            .antt-section-title {{ color: var(--antt-ink); font-weight: 850; font-size: 1.05rem; margin: 18px 0 8px; }}
            .antt-nav-chart {{ background: #ffffff; border: 1px solid var(--antt-border); border-radius: 8px; box-shadow: 0 10px 24px rgba(21, 49, 37, 0.05); padding: 12px; margin-bottom: 14px; }}
            .antt-nav-chart-title {{ color: var(--antt-ink); font-weight: 850; font-size: 1rem; margin-bottom: 3px; }}
            .antt-nav-chart-help {{ color: var(--antt-muted); font-size: .78rem; margin-bottom: 10px; }}
            .antt-bar-track {{ height: 14px; background: #e4eee7; border-radius: 999px; overflow: hidden; margin-top: 11px; }}
            .antt-bar-fill {{ height: 14px; border-radius: 999px; min-width: 4px; }}
            .js-plotly-plot, .stPlotlyChart {{ background: var(--antt-panel); border-radius: 8px; }}
            .stPlotlyChart {{ border: 1px solid var(--antt-border); box-shadow: 0 10px 24px rgba(21, 49, 37, 0.05); padding: 8px 10px 2px; }}
            .stButton > button,
            .stDownloadButton > button,
            button[kind="primary"] {{ background: var(--antt-green); color: #ffffff; border: 1px solid var(--antt-green); border-radius: 6px; font-weight: 750; }}
            .stButton > button:hover,
            .stDownloadButton > button:hover {{ border-color: var(--antt-blue); color: #ffffff; background: var(--antt-blue); }}
            [data-testid="stSidebar"] .stButton > button {{ background: #ffffff; color: var(--antt-green); border-color: #cfe8d7; }}
            [data-testid="stSidebar"] .stButton > button:hover {{ background: #e4f4e9; color: var(--antt-green); border-color: var(--antt-green); }}
            [data-baseweb="select"] > div,
            [data-baseweb="input"] > div,
            textarea {{ border-radius: 6px !important; }}
            @media (max-width: 900px) {{
                .antt-page-header {{ align-items: flex-start; flex-direction: column; }}
                .antt-chip-row {{ justify-content: flex-start; }}
                .antt-title {{ font-size: 1.55rem; }}
                .st-key-dash_right_filters {{
                    position: static;
                    width: auto;
                    max-height: none;
                    overflow: visible;
                    border-left: 0;
                    box-shadow: none;
                    padding: 0;
                    margin-bottom: 12px;
                }}
                .block-container:has(.st-key-dash_right_filters) {{
                    padding-right: 1rem;
                }}
            }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def antt_logo_svg(width=300):
    return f"""
    <svg viewBox="0 0 360 176" width="{width}" role="img" aria-label="ANTT - Agência Nacional de Transportes Terrestres" xmlns="http://www.w3.org/2000/svg">
        <g transform="translate(0 0)">
            <polygon points="70,0 35,66 105,66" fill="#0057a6"/>
            <polygon points="36,62 0,124 74,124" fill="#00843d"/>
            <polygon points="105,62 69,124 143,124" fill="#0057a6"/>
            <polygon points="45,60 113,44 78,104" fill="#f6b728" stroke="#ffffff" stroke-width="2.4"/>
        </g>
        <g transform="translate(150 48)" fill="#00843d">
            <text x="0" y="67" font-family="Arial Black, Arial, Helvetica, sans-serif" font-size="64" font-weight="900" letter-spacing="0">ANTT</text>
        </g>
        <text x="23" y="148" font-family="Arial, Helvetica, sans-serif" font-size="18" font-weight="800" fill="#2f2f35" letter-spacing="1.2">AGÊNCIA NACIONAL DE</text>
        <text x="0" y="172" font-family="Arial, Helvetica, sans-serif" font-size="18" font-weight="800" fill="#2f2f35" letter-spacing="1.2">TRANSPORTES TERRESTRES</text>
    </svg>
    """

def antt_logo_html(width=150, class_name=""):
    if LOGO_ANTT_PATH.exists():
        encoded = base64.b64encode(LOGO_ANTT_PATH.read_bytes()).decode("ascii")
        return (
            f'<div class="{class_name}">'
            f'<img class="antt-logo-img" src="data:image/svg+xml;base64,{encoded}" '
            f'width="{width}" alt="ANTT - Agência Nacional de Transportes Terrestres">'
            f'</div>'
        )
    return f'<div class="{class_name}">{antt_logo_svg(width)}</div>'


def page_header(titulo, subtitulo="", chips=None):
    chips = chips or []
    chips_html = "".join([f'<span class="antt-chip">{chip}</span>' for chip in chips])
    st.markdown(
        f"""
        <div class="antt-page-header">
            <div>
                <div class="antt-eyebrow">Controle de Medidas Preventivas</div>
                <div class="antt-title">{titulo}</div>
                <div class="antt-subtitle">{subtitulo}</div>
            </div>
            <div class="antt-chip-row">{chips_html}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def kpi_card(label, value, helper="", color="#f2c811"):
    st.markdown(
        f"""
        <div class="antt-kpi" style="--kpi-color:{color}">
            <div class="antt-kpi-label">{label}</div>
            <div class="antt-kpi-value">{value}</div>
            <div class="antt-kpi-help">{helper}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def page_actions(items):
    st.markdown('<div class="antt-page-actions">', unsafe_allow_html=True)
    cols = st.columns(max(1, len(items)))
    for col, item in zip(cols, items):
        with col:
            if st.button(item["label"], key=item["key"], use_container_width=True):
                item["callback"]()
    st.markdown('</div>', unsafe_allow_html=True)


def meta_card(label, value, helper=""):
    st.markdown(
        f"""
        <div class="antt-list-meta">
            <div class="antt-list-meta-label">{label}</div>
            <div class="antt-list-meta-value">{value}</div>
            <div class="antt-list-meta-help">{helper}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def resumo_lista(df, tipo="medidas"):
    total = 0 if df is None else len(df)
    concessionarias = df["concessionaria"].nunique() if df is not None and not df.empty and "concessionaria" in df.columns else 0
    disciplinas = df["disciplina"].nunique() if df is not None and not df.empty and "disciplina" in df.columns else 0
    criticas = df["situacao_prazo"].isin(["Vencida", "A vencer em 15 dias", "A vencer em 30 dias"]).sum() if df is not None and not df.empty and "situacao_prazo" in df.columns else 0
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        meta_card("Registros", total, tipo)
    with c2:
        meta_card("Concessionárias", concessionarias, "no resultado")
    with c3:
        meta_card("Disciplinas", disciplinas, "no resultado")
    with c4:
        meta_card("Atenção", int(criticas), "vencidas ou próximas")


def aviso_selecao(texto="Selecione uma linha para abrir detalhes e ações."):
    st.markdown(f'<div class="antt-detail-callout">{texto}</div>', unsafe_allow_html=True)


def contexto_tabela(df, nome):
    st.markdown(f'<div class="antt-section-title">{nome}</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="antt-muted-note">{len(df)} registros exibidos com os filtros atuais.</div>', unsafe_allow_html=True)


COLUNAS_TABELA_API = [
    "id_api", "disciplina", "ano_concessao", "concessionaria", "data_emissao",
    "status_emissao", "processo_sei", "documento_sei", "objeto", "observacao",
]
COLUNAS_TABELA_ANC = [
    "id_anc", "disciplina", "ano_concessao", "concessionaria", "prazo", "situacao_prazo",
    "status", "data_emissao", "processo_sei", "documento_sei", "documento_sei_comprovacao",
    "observacao_fiscalizacao",
]
COLUNAS_TABELA_AE = [
    "id_ae", "disciplina", "ano_concessao", "concessionaria", "prazo_acao_educativa",
    "situacao_prazo", "finalizada", "data_emissao", "processo_sei", "documento_sei_instaurador",
    "escopo", "observacao",
]
COLUNAS_TABELA_GERAL = [
    "tipo", "id_medida", "disciplina", "ano_concessao", "concessionaria", "data_emissao",
    "prazo_referencia", "situacao_prazo", "status", "finalizada", "processo_sei", "documento_referencia",
    "resumo",
]

RENOMEAR_COLUNAS_TABELA = {
    "tipo": "Tipo",
    "id_medida": "ID da medida",
    "id_api": "ID da API",
    "id_anc": "ID da ANC",
    "id_ae": "ID da AE",
    "disciplina": "Disciplina",
    "ano_concessao": "Ano concessão",
    "concessionaria": "Concessionária",
    "data_emissao": "Data de emissão",
    "prazo": "Prazo ANC",
    "prazo_acao_educativa": "Prazo AE",
    "prazo_referencia": "Prazo",
    "situacao_prazo": "Situação prazo",
    "status": "Status ANC",
    "status_emissao": "Status emissão",
    "finalizada": "AE finalizada",
    "processo_sei": "Processo SEI",
    "documento_sei": "Documento SEI",
    "documento_sei_instaurador": "Documento SEI instaurador",
    "documento_referencia": "Documento SEI",
    "documento_sei_comprovacao": "Documento SEI comprovação",
    "objeto": "Objeto",
    "escopo": "Escopo",
    "observacao": "Observação",
    "observacao_fiscalizacao": "Observação fiscalização",
    "resumo": "Resumo",
}


def preparar_tabela(df, colunas):
    if df is None or df.empty:
        return df
    colunas_presentes = [c for c in colunas if c in df.columns]
    # Mantem a seleção funcional e evita campos técnicos/duplicados por padrão.
    out = df[colunas_presentes].copy()
    if "ano_concessao" in out.columns:
        out["ano_concessao"] = out["ano_concessao"].apply(formatar_ano_concessao)
    out = formatar_colunas_data(out)
    return out.rename(columns={k: v for k, v in RENOMEAR_COLUNAS_TABELA.items() if k in out.columns})


def preparar_tabela_api(df):
    return preparar_tabela(df, COLUNAS_TABELA_API)


def preparar_tabela_anc(df):
    return preparar_tabela(df, COLUNAS_TABELA_ANC)


def preparar_tabela_ae(df):
    return preparar_tabela(df, COLUNAS_TABELA_AE)


def preparar_tabela_geral(df):
    if df is None or df.empty:
        return df
    out = df.copy()
    if "documento_referencia" not in out.columns:
        out["documento_referencia"] = ""
        if "documento_sei" in out.columns:
            out["documento_referencia"] = out["documento_sei"].fillna("")
        if "documento_sei_instaurador" in out.columns:
            mask = out["documento_referencia"].astype(str).str.strip() == ""
            out.loc[mask, "documento_referencia"] = out.loc[mask, "documento_sei_instaurador"].fillna("")
    if "resumo" not in out.columns:
        out["resumo"] = ""
        for col in ["objeto", "escopo", "observacao_fiscalizacao", "observacao"]:
            if col in out.columns:
                mask = out["resumo"].astype(str).str.strip() == ""
                out.loc[mask, "resumo"] = out.loc[mask, col].fillna("")
    return preparar_tabela(out, COLUNAS_TABELA_GERAL)


COLUNAS_PENDENCIA_ANC = [
    "id", "id_anc", "disciplina", "ano_concessao", "concessionaria", "prazo", "situacao_prazo",
    "status_resposta", "data_resposta", "documento_sei_comprovacao", "processo_sei", "observacao_concessionaria",
]
COLUNAS_PENDENCIA_AE = [
    "id", "id_ae", "disciplina", "ano_concessao", "concessionaria", "prazo_acao_educativa", "situacao_prazo",
    "mes_acompanhamento", "status_acompanhamento", "data_envio", "documento_sei_acompanhamento",
    "processo_sei", "observacao_concessionaria",
]
RENOMEAR_COLUNAS_TABELA.update({
    "id": "ID pendência",
    "id_anc": "ID da ANC",
    "id_ae": "ID da AE",
    "status_resposta": "Status validação",
    "status_acompanhamento": "Status validação",
    "data_resposta": "Data envio",
    "data_envio": "Data envio",
    "mes_acompanhamento": "Mês acompanhamento",
    "documento_sei_acompanhamento": "Documento SEI acompanhamento",
    "observacao_concessionaria": "Observação concessionária",
})


def preparar_tabela_pendencia_anc(df):
    return preparar_tabela(df, COLUNAS_PENDENCIA_ANC)


def preparar_tabela_pendencia_ae(df):
    return preparar_tabela(df, COLUNAS_PENDENCIA_AE)


def id_selecionado(selected, coluna_exibida, coluna_original=None):
    if not selected:
        return None
    if coluna_exibida in selected:
        return selected[coluna_exibida]
    if coluna_original and coluna_original in selected:
        return selected[coluna_original]
    return None

# ============================================================
# Utilitários
# ============================================================

def hash_password(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def config_value(nome, padrao=""):
    env_value = os.environ.get(nome)
    if env_value:
        return env_value
    try:
        secret_value = st.secrets.get(nome)
        if secret_value:
            return str(secret_value)
    except Exception:
        pass
    return padrao


def connect_db():
    return sqlite3.connect(DB_PATH, check_same_thread=False)


def execute(query, params=()):
    conn = connect_db()
    cur = conn.cursor()
    cur.execute(query, params)
    conn.commit()
    conn.close()


def query_df(query, params=()):
    conn = connect_db()
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df


def query_one(query, params=()):
    conn = connect_db()
    cur = conn.cursor()
    cur.execute(query, params)
    row = cur.fetchone()
    conn.close()
    return row


def safe_text(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()


def safe_int(value):
    text = safe_text(value)
    if not text:
        return None
    nums = re.sub(r"\D", "", text)
    if not nums:
        return None
    try:
        return int(nums)
    except Exception:
        return None


def safe_date(value):
    if value is None or safe_text(value) == "":
        return ""
    data = parse_date(value)
    return data.isoformat() if data else ""


def parse_date(value):
    if value is None or safe_text(value) == "":
        return None
    text = safe_text(value)
    try:
        return datetime.fromisoformat(text).date()
    except Exception:
        try:
            if re.match(r"^\d{1,2}/\d{1,2}/\d{4}$", text):
                return datetime.strptime(text, "%d/%m/%Y").date()
            return pd.to_datetime(value, dayfirst=True).date()
        except Exception:
            return None


def formatar_data_br(value):
    data = parse_date(value)
    return data.strftime("%d/%m/%Y") if data else ""


def formatar_colunas_data(df):
    if df is None or df.empty:
        return df
    out = df.copy()
    for col in out.columns:
        col_lower = str(col).lower()
        if "data" in col_lower or "prazo" in col_lower:
            out[col] = out[col].apply(formatar_data_br)
    return out


def date_input_br(label, *args, **kwargs):
    kwargs.setdefault("format", "DD/MM/YYYY")
    return st.date_input(label, *args, **kwargs)


def apenas_numeros(valor):
    return re.sub(r"\D", "", safe_text(valor))


def formatar_processo_sei(processo):
    nums = apenas_numeros(processo)
    if len(nums) != 17:
        return ""
    return f"{nums[0:5]}.{nums[5:11]}/{nums[11:15]}-{nums[15:17]}"


def validar_processo_sei(processo):
    return len(apenas_numeros(processo)) == 17


def normalizar_documento_sei(documento):
    return apenas_numeros(documento)


def validar_documento_sei(documento):
    return bool(re.match(REGEX_DOCUMENTO_SEI, safe_text(documento)))


def gerar_id_api(documento):
    doc = normalizar_documento_sei(documento)
    return f"API-{doc}" if doc else ""


def gerar_id_anc(documento):
    doc = normalizar_documento_sei(documento)
    return f"ANC-{doc}" if doc else ""


def gerar_id_ae(documento):
    doc = normalizar_documento_sei(documento)
    return f"AE-{doc}" if doc else ""


def formatar_ano_concessao(ano):
    numero = safe_int(ano)
    if numero is None:
        return ""
    return f"{numero}º"


def limpar_selecao():
    st.session_state["grid_nonce"] = st.session_state.get("grid_nonce", 0) + 1


def key_grid(base):
    return f"{base}_{st.session_state.get('grid_nonce', 0)}"


def ir_para(pagina, filtros=None):
    # Guarda o destino para ser aplicado antes da criação do st.sidebar.radio na próxima execução.
    st.session_state["proxima_pagina"] = pagina

    if filtros:
        for k, v in filtros.items():
            st.session_state[k] = v

    limpar_selecao()
    st.rerun()


def filtros_para_tela(prefixo, conc=None, disc=None, ano_conc=None, ano_emissao=None, mes_emissao=None, situacao=None):
    filtros = {
        f"{prefixo}_conc": conc or [],
        f"{prefixo}_disc": disc or [],
        f"{prefixo}_ano_conc": ano_conc or [],
        f"{prefixo}_ano_emissao": ano_emissao or [],
        f"{prefixo}_mes_emissao": mes_emissao or [],
    }
    if situacao:
        filtros[f"{prefixo}_situacao"] = situacao
    return filtros


def botao_drillthrough(label, pagina, prefixo, conc=None, disc=None, ano_conc=None, ano_emissao=None, mes_emissao=None, situacao=None, key=None):
    if st.button(label, key=key, use_container_width=True):
        ir_para(pagina, filtros_para_tela(prefixo, conc, disc, ano_conc, ano_emissao, mes_emissao, situacao))


def valor_ponto_plotly(ponto, campo="x"):
    if ponto is None:
        return None
    if hasattr(ponto, "get"):
        valor = ponto.get(campo)
        if valor is not None:
            return valor
        if campo == "customdata" and ponto.get("customdata"):
            return ponto.get("customdata")
    try:
        valor = getattr(ponto, campo)
        if valor is not None:
            return valor
    except Exception:
        pass
    return None


def botoes_drillthrough_tipos(conc, disc, ano_conc, ano_emissao, mes_emissao, prefixo="chart_tipo"):
    b1, b2, b3 = st.columns(3)
    with b1:
        botao_drillthrough("Abrir APIs", NOME_API, "api", conc, disc, ano_conc, ano_emissao, mes_emissao, key=f"{prefixo}_api")
    with b2:
        botao_drillthrough("Abrir ANCs", NOME_ANC, "anc", conc, disc, ano_conc, ano_emissao, mes_emissao, key=f"{prefixo}_anc")
    with b3:
        botao_drillthrough("Abrir AEs", NOME_AE, "ae", conc, disc, ano_conc, ano_emissao, mes_emissao, key=f"{prefixo}_ae")


def destaque_sei(label, valor):
    valor = normalizar_documento_sei(valor)
    texto = valor if valor else "Não informado"
    st.markdown(
        f'<div class="antt-detail-callout"><strong>{label}:</strong> {texto}</div>',
        unsafe_allow_html=True,
    )


def cor_barra(valor, mapa=None, padrao="#2563eb"):
    if mapa and valor in mapa:
        return mapa[valor]
    return padrao


def grafico_barra_basico(df, x, y, titulo, color=None, color_map=None, hover_cols=None, key=None):
    if df is None or df.empty:
        st.info("Sem dados para exibir.")
        return
    if px is None:
        st.bar_chart(df.set_index(x)[y])
        return
    fig = px.bar(
        df,
        x=x,
        y=y,
        color=color,
        text=y,
        template=PLOTLY_TEMPLATE,
        color_discrete_map=color_map,
        hover_data=hover_cols,
    )
    if color is None and color_map:
        fig.update_traces(marker_color=[color_map.get(v, POWERBI_COLORS["green"]) for v in df[x].tolist()])
    fig.update_traces(
        textposition="outside",
        cliponaxis=False,
        marker_line_width=0,
        hovertemplate="<b>%{x}</b><br>Quantidade: %{y}<extra></extra>",
    )
    fig.update_layout(
        title=dict(text=titulo, font=dict(size=17, color=POWERBI_COLORS["ink"]), x=0.02),
        height=340,
        margin=dict(l=18, r=18, t=52, b=42),
        paper_bgcolor="white",
        plot_bgcolor="white",
        font=dict(color=POWERBI_COLORS["ink"], family="Segoe UI, Arial"),
        xaxis=dict(title="", showgrid=False, tickfont=dict(size=11)),
        yaxis=dict(title="Quantidade", gridcolor="#e4eee7", zeroline=False),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        legend_title_text="",
    )
    st.plotly_chart(fig, use_container_width=True, key=key)


def grafico_barras_navegavel(df, label_col, valor_col, titulo, key, on_click, color_map=None, max_rows=12):
    st.markdown(
        f'<div class="antt-nav-chart"><div class="antt-nav-chart-title">{titulo}</div>'
        '<div class="antt-nav-chart-help">Use o botão da linha para abrir a tabela já filtrada.</div>',
        unsafe_allow_html=True,
    )
    if df is None or df.empty:
        st.info("Sem dados para exibir.")
        st.markdown('</div>', unsafe_allow_html=True)
        return
    view = df.copy().head(max_rows)
    max_val = max(float(view[valor_col].max()), 1.0)
    for idx, row in view.reset_index(drop=True).iterrows():
        label = safe_text(row[label_col]) or "Não informado"
        valor = int(row[valor_col]) if not pd.isna(row[valor_col]) else 0
        pct = max(4, min(100, round(valor / max_val * 100, 1)))
        color = cor_barra(label, color_map, "#2563eb")
        c1, c2 = st.columns([0.42, 0.58], vertical_alignment="center")
        with c1:
            if st.button(f"{label} · {valor}", key=f"{key}_{idx}_{label}", use_container_width=True):
                on_click(row)
        with c2:
            st.markdown(
                f'<div class="antt-bar-track"><div class="antt-bar-fill" style="width:{pct}%; background:{color};"></div></div>',
                unsafe_allow_html=True,
            )
    st.markdown('</div>', unsafe_allow_html=True)


def sidebar_brand():
    st.sidebar.markdown(
        f"""
        <div class="antt-sidebar-brand">
            {antt_logo_html(224, "antt-sidebar-logo")}
            <div class="antt-sidebar-brand-title">Controle de Medidas Preventivas</div>
            <div class="antt-sidebar-brand-subtitle">Apoio operacional, prazos e validação SEI</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def sidebar_profile():
    usuario = st.session_state.get("email", "")
    perfil = st.session_state.get("perfil", "")
    concessionaria = st.session_state.get("concessionaria", "")
    conc_html = ""
    if perfil == "Concessionária":
        conc_html = f'<div class="antt-sidebar-profile-label">Concessionária</div><div class="antt-sidebar-profile-value">{concessionaria}</div>'
    st.sidebar.markdown(
        f"""
        <div class="antt-sidebar-profile">
            <div class="antt-sidebar-profile-label">Usuário</div><div class="antt-sidebar-profile-value">{usuario}</div>
            <div class="antt-sidebar-profile-label">Perfil</div><div class="antt-sidebar-profile-value">{perfil}</div>
            {conc_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def filtro_situacao_prazo(df, prefixo, opcoes=None):
    if df is None or df.empty or "situacao_prazo" not in df.columns:
        return df
    opcoes = opcoes or sorted(df["situacao_prazo"].dropna().unique().tolist())
    selecionadas = st.multiselect("Situação de prazo", opcoes, key=f"{prefixo}_situacao")
    if selecionadas:
        return df[df["situacao_prazo"].isin(selecionadas)]
    return df

def campo_processo_sei(chave, valor_inicial=""):
    digitado = st.text_input(
        "Processo SEI",
        value=apenas_numeros(valor_inicial),
        placeholder="Digite somente números. Ex: 50505136555202461",
        key=chave,
    )
    formatado = formatar_processo_sei(digitado)
    if digitado:
        st.caption(f"Processo formatado: {formatado}" if formatado else "Digite 17 números para formar 00000.000000/0000-00.")
    return digitado, formatado


def campo_documento_sei(label, chave, valor_inicial="", placeholder="Somente números"):
    digitado = st.text_input(label, value=normalizar_documento_sei(valor_inicial), placeholder=placeholder, key=chave)
    normalizado = normalizar_documento_sei(digitado)
    if digitado:
        st.caption(f"Documento SEI considerado: {normalizado}" if normalizado else "Informe somente números.")
    return digitado, normalizado


def situacao_prazo_anc(status, prazo):
    status = safe_text(status).lower()
    prazo_dt = parse_date(prazo)
    if status == "cumprido":
        return "Cumprido"
    if prazo_dt is None:
        return "Sem prazo informado"
    if prazo_dt < date.today():
        return "Vencida"
    if prazo_dt <= date.today() + pd.Timedelta(days=15):
        return "A vencer em 15 dias"
    return "Dentro do prazo"


def situacao_prazo_ae(finalizada, prazo_ae):
    finalizada = safe_text(finalizada).lower()
    prazo_dt = parse_date(prazo_ae)
    if finalizada == "sim":
        return "Finalizada"
    if prazo_dt is None:
        return "Sem prazo informado"
    if prazo_dt < date.today():
        return "Vencida"
    if prazo_dt <= date.today() + pd.Timedelta(days=30):
        return "A vencer em 30 dias"
    return "Dentro do prazo"


def audit(usuario, perfil, acao, tabela, registro_id, valor_anterior="", valor_novo=""):
    execute(
        """
        INSERT INTO auditoria (usuario, perfil, data_hora, acao, tabela, registro_id, valor_anterior, valor_novo)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (usuario, perfil, datetime.now().isoformat(timespec="seconds"), acao, tabela, registro_id, valor_anterior, valor_novo),
    )


def salvar_upload(uploaded_file, prefixo):
    if uploaded_file is None:
        return ""
    ext = uploaded_file.name.split(".")[-1].lower()
    if ext not in EXTENSOES_UPLOAD:
        raise ValueError("Somente arquivos PDF ou HTML são permitidos.")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{prefixo}_{timestamp}_{uploaded_file.name}"
    path = UPLOAD_DIR / filename
    with open(path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return str(path)


def mostrar_arquivo_upload(caminho, titulo="Arquivo enviado"):
    caminho = safe_text(caminho)
    if not caminho:
        st.caption("Nenhum arquivo anexado.")
        return
    path = Path(caminho)
    if not path.exists():
        st.warning(f"Arquivo não encontrado no servidor: {caminho}")
        return
    ext = path.suffix.lower()
    dados = path.read_bytes()
    mime = "application/pdf" if ext == ".pdf" else "text/html"
    st.download_button(f"Abrir/Baixar {titulo}", data=dados, file_name=path.name, mime=mime, key=f"download_{titulo}_{path.name}")
    with st.expander(f"Visualizar {titulo}", expanded=False):
        if ext == ".pdf":
            b64 = base64.b64encode(dados).decode("utf-8")
            st.markdown(f'<iframe src="data:application/pdf;base64,{b64}" width="100%" height="650"></iframe>', unsafe_allow_html=True)
        elif ext in [".html", ".htm"]:
            html = dados.decode("utf-8", errors="ignore")
            components.html(html, height=650, scrolling=True)
        else:
            st.info("Visualização disponível apenas para PDF ou HTML.")

# ============================================================
# Tabelas e filtros
# ============================================================

def tabela_interativa(df, key, height=430):
    if df is None or df.empty:
        st.info("Nenhum registro encontrado.")
        return None
    df_show = df.copy()
    if AgGrid is None:
        st.dataframe(df_show, use_container_width=True, height=height)
        st.caption("Para seleção por clique e tooltip avançado, instale streamlit-aggrid.")
        return None
    gb = GridOptionsBuilder.from_dataframe(df_show)
    gb.configure_default_column(
        sortable=True,
        filter=True,
        resizable=True,
        wrapText=True,
        autoHeight=False,
        minWidth=150,
        width=190,
    )
    for col in df_show.columns:
        if col in ["objeto", "observacao", "observacao_fiscalizacao", "escopo", "observacao_concessionaria", "observacao_validacao_antt"]:
            gb.configure_column(col, width=420, wrapText=True, autoHeight=False)
        elif col in ["processo_sei", "id_api", "id_anc", "id_ae", "documento_sei", "documento_sei_instaurador"]:
            gb.configure_column(col, width=230)
        elif "data" in col or "prazo" in col:
            gb.configure_column(col, width=170)
    gb.configure_selection(selection_mode="single", use_checkbox=False)
    grid_options = gb.build()
    grid_options["rowHeight"] = 78  # aproximadamente três linhas
    for col_def in grid_options.get("columnDefs", []):
        field = col_def.get("field")
        if field:
            col_def["tooltipField"] = field
            col_def["cellStyle"] = {"whiteSpace": "normal", "lineHeight": "22px", "overflow": "hidden"}
    grid_options["enableBrowserTooltips"] = True
    response = AgGrid(
        df_show,
        gridOptions=grid_options,
        update_mode=GridUpdateMode.SELECTION_CHANGED,
        fit_columns_on_grid_load=False,
        height=height,
        allow_unsafe_jscode=True,
        key=key,
    )
    selected = response.get("selected_rows", [])
    if isinstance(selected, pd.DataFrame) and not selected.empty:
        return selected.iloc[0].to_dict()
    if isinstance(selected, list) and selected:
        return selected[0]
    return None


def filtros_medidas(df, prefixo="filtro", layout="horizontal", mostrar_titulo=True):
    if df is None:
        df = pd.DataFrame()
    if layout == "vertical":
        titulo_html = '<div class="antt-filter-title">Filtros</div>' if mostrar_titulo else ''
        st.markdown(
            titulo_html + '<div class="antt-filter-help">Refine a visão geral antes de abrir as planilhas.</div>',
            unsafe_allow_html=True,
        )
        c1 = c2 = c3 = c4 = c5 = st
    else:
        st.subheader("Filtros")
        c1, c2, c3, c4, c5 = st.columns(5)
    if st.session_state.get("perfil") == "Concessionária":
        conc_padrao = [st.session_state.get("concessionaria")]
        c1.info(f"Concessionária: {conc_padrao[0]}")
        conc = conc_padrao
    else:
        conc = c1.multiselect("Concessionária", CONCESSIONARIAS, key=f"{prefixo}_conc")
    disc = c2.multiselect("Disciplina", DISCIPLINAS, key=f"{prefixo}_disc")
    anos_base = []
    if not df.empty and "ano_concessao_limpo" in df.columns:
        anos_base = sorted([int(x) for x in df["ano_concessao_limpo"].dropna().unique() if safe_text(x) != ""])
    ano_conc = c3.multiselect("Ano da concessão", anos_base, key=f"{prefixo}_ano_conc")
    anos_emissao = []
    datas_emissao = pd.Series(dtype="datetime64[ns]")
    if not df.empty and "data_emissao" in df.columns:
        datas_emissao = pd.to_datetime(df["data_emissao"], errors="coerce").dropna()
        anos_emissao = sorted(datas_emissao.dt.year.unique().astype(int).tolist())
    ano_emissao = c4.multiselect("Ano da emissão", anos_emissao, key=f"{prefixo}_ano_emissao")
    meses_emissao = []
    if ano_emissao and not datas_emissao.empty:
        datas_do_ano = datas_emissao[datas_emissao.dt.year.isin([int(x) for x in ano_emissao])]
        meses_emissao = sorted(datas_do_ano.dt.month.unique().astype(int).tolist())
    mes_key = f"{prefixo}_mes_emissao"
    if not ano_emissao and mes_key in st.session_state:
        st.session_state[mes_key] = []
    elif mes_key in st.session_state:
        st.session_state[mes_key] = [m for m in st.session_state[mes_key] if int(m) in meses_emissao]
    mes_emissao = c5.multiselect(
        "Mês da emissão",
        meses_emissao,
        format_func=lambda m: f"{int(m):02d}",
        key=mes_key,
        disabled=not bool(ano_emissao),
        placeholder="Selecione o ano primeiro",
    )
    return conc, disc, ano_conc, ano_emissao, mes_emissao


def filtros_medidas_estado(prefixo="filtro"):
    if st.session_state.get("perfil") == "Concessionária":
        conc = [st.session_state.get("concessionaria")]
    else:
        conc = st.session_state.get(f"{prefixo}_conc", [])
    return (
        conc,
        st.session_state.get(f"{prefixo}_disc", []),
        st.session_state.get(f"{prefixo}_ano_conc", []),
        st.session_state.get(f"{prefixo}_ano_emissao", []),
        st.session_state.get(f"{prefixo}_mes_emissao", []),
    )


def aplicar_filtros(df, conc=None, disc=None, ano_conc=None, ano_emissao=None, mes_emissao=None):
    if df is None or df.empty:
        return df
    out = df.copy()
    if conc and "concessionaria" in out.columns:
        out = out[out["concessionaria"].isin(conc)]
    if disc and "disciplina" in out.columns:
        out = out[out["disciplina"].isin(disc)]
    if ano_conc and "ano_concessao_limpo" in out.columns:
        out = out[out["ano_concessao_limpo"].fillna(-1).astype(int).isin([int(x) for x in ano_conc])]
    if (ano_emissao or mes_emissao) and "data_emissao" in out.columns:
        datas = pd.to_datetime(out["data_emissao"], errors="coerce")
        mask = pd.Series(True, index=out.index)
        if ano_emissao:
            mask &= datas.dt.year.isin([int(x) for x in ano_emissao])
        if mes_emissao:
            mask &= datas.dt.month.isin([int(x) for x in mes_emissao])
        out = out[mask]
    return out

# ============================================================
# Banco
# ============================================================

def init_db():
    execute("""
        CREATE TABLE IF NOT EXISTS usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT UNIQUE,
            senha_hash TEXT,
            perfil TEXT,
            concessionaria TEXT,
            ativo INTEGER DEFAULT 1
        )
    """)
    execute("""
        CREATE TABLE IF NOT EXISTS api (
            id_api TEXT PRIMARY KEY, processo_sei TEXT, documento_sei TEXT, concessionaria TEXT,
            ano_concessao TEXT, ano_concessao_limpo INTEGER, disciplina TEXT, data_emissao TEXT,
            status_emissao TEXT, objeto TEXT, observacao TEXT, tipo_medida TEXT
        )
    """)
    execute("""
        CREATE TABLE IF NOT EXISTS anc (
            id_anc TEXT PRIMARY KEY, processo_sei TEXT, documento_sei TEXT, concessionaria TEXT,
            ano_concessao TEXT, ano_concessao_limpo INTEGER, disciplina TEXT, data_emissao TEXT,
            data_recebimento_oficio TEXT, prazo_dias_uteis INTEGER, prazo TEXT, status TEXT,
            documento_sei_comprovacao TEXT, observacao_fiscalizacao TEXT, tipo_medida TEXT
        )
    """)
    execute("""
        CREATE TABLE IF NOT EXISTS ae (
            id_ae TEXT PRIMARY KEY, processo_sei TEXT, documento_sei_instaurador TEXT, concessionaria TEXT,
            ano_concessao TEXT, ano_concessao_limpo INTEGER, disciplina TEXT, data_emissao TEXT,
            data_alinhamento_escopo TEXT, escopo TEXT, prazo_envio_cronograma TEXT, prazo_acao_educativa TEXT,
            documento_sei_cronograma TEXT, documento_sei_acomp_1 TEXT, documento_sei_acomp_2 TEXT,
            documento_sei_acomp_3 TEXT, documento_sei_acomp_4 TEXT, documento_sei_acomp_5 TEXT,
            status_automatico TEXT, finalizada TEXT, observacao TEXT, tipo_medida TEXT
        )
    """)
    execute("""
        CREATE TABLE IF NOT EXISTS anc_respostas (
            id INTEGER PRIMARY KEY AUTOINCREMENT, id_anc TEXT, concessionaria TEXT, data_resposta TEXT,
            usuario_responsavel TEXT, documento_sei_comprovacao TEXT, observacao_concessionaria TEXT,
            arquivo_link TEXT, status_resposta TEXT, data_validacao_antt TEXT, usuario_antt_validador TEXT,
            observacao_validacao_antt TEXT
        )
    """)
    execute("""
        CREATE TABLE IF NOT EXISTS ae_acompanhamentos (
            id INTEGER PRIMARY KEY AUTOINCREMENT, id_ae TEXT, concessionaria TEXT, mes_acompanhamento INTEGER,
            data_envio TEXT, usuario_responsavel TEXT, documento_sei_acompanhamento TEXT,
            observacao_concessionaria TEXT, arquivo_link TEXT, status_acompanhamento TEXT,
            data_validacao_antt TEXT, usuario_antt_validador TEXT, observacao_validacao_antt TEXT
        )
    """)
    execute("""
        CREATE TABLE IF NOT EXISTS auditoria (
            id INTEGER PRIMARY KEY AUTOINCREMENT, usuario TEXT, perfil TEXT, data_hora TEXT, acao TEXT,
            tabela TEXT, registro_id TEXT, valor_anterior TEXT, valor_novo TEXT
        )
    """)
    seed_users()
    normalizar_anos_concessao_db()



def normalizar_anos_concessao_db():
    chaves = {"api": "id_api", "anc": "id_anc", "ae": "id_ae"}
    for tabela, chave in chaves.items():
        try:
            df = query_df(f"SELECT {chave}, ano_concessao FROM {tabela}")
        except Exception:
            continue
        for _, row in df.iterrows():
            ano_formatado = formatar_ano_concessao(row.get("ano_concessao"))
            if ano_formatado and safe_text(row.get("ano_concessao")) != ano_formatado:
                execute(f"UPDATE {tabela} SET ano_concessao=? WHERE {chave}=?", (ano_formatado, row[chave]))


def seed_users():
    count = query_one("SELECT COUNT(*) FROM usuarios")[0]
    if count > 0:
        return
    usuarios = [
        (config_value("ADMIN_EMAIL", "admin@antt.gov.br"), config_value("ADMIN_PASSWORD", "admin123"), "ANTT", ""),
        (config_value("VIEWER_EMAIL", "visualizador@antt.gov.br"), config_value("VIEWER_PASSWORD", "visualizar123"), "Visualizador", ""),
        (config_value("LITORAL_EMAIL", "litoralsul@concessionaria.com"), config_value("LITORAL_PASSWORD", "litoral123"), "Concessionária", "Litoral Sul"),
        (config_value("PLANALTO_EMAIL", "planaltosul@concessionaria.com"), config_value("PLANALTO_PASSWORD", "planalto123"), "Concessionária", "Planalto Sul"),
        (config_value("VIACOSTEIRA_EMAIL", "viacosteira@concessionaria.com"), config_value("VIACOSTEIRA_PASSWORD", "via123"), "Concessionária", "ViaCosteira"),
    ]
    for email, senha, perfil, concessionaria in usuarios:
        execute("INSERT INTO usuarios (email, senha_hash, perfil, concessionaria, ativo) VALUES (?, ?, ?, ?, 1)",
                (email, hash_password(senha), perfil, concessionaria))

# ============================================================
# Importação Excel
# ============================================================

def obter_aba_excel(xls, preferencial, alternativa):
    if preferencial in xls.sheet_names:
        return preferencial
    if alternativa in xls.sheet_names:
        return alternativa
    return None


def importar_excel_se_necessario():
    total_api = query_one("SELECT COUNT(*) FROM api")[0]
    total_anc = query_one("SELECT COUNT(*) FROM anc")[0]
    total_ae = query_one("SELECT COUNT(*) FROM ae")[0]
    if total_api + total_anc + total_ae > 0:
        return
    if not os.path.exists(EXCEL_PATH):
        return
    try:
        xls = pd.ExcelFile(EXCEL_PATH, engine="openpyxl")
        importar_planilha_xls_para_banco(xls, "Acrescentar ou atualizar registros", False, EXCEL_PATH)
    except Exception:
        return


def importar_planilha_xls_para_banco(xls, modo_importacao, registrar_auditoria=True, nome_arquivo="arquivo_excel"):
    aba_api = obter_aba_excel(xls, "Fato_API_BI", "API")
    aba_anc = obter_aba_excel(xls, "Fato_ANC_BI", "ANC")
    aba_ae = obter_aba_excel(xls, "Fato_AE_BI", "AcaoEducativa")
    if modo_importacao == "Substituir medidas existentes":
        execute("DELETE FROM api")
        execute("DELETE FROM anc")
        execute("DELETE FROM ae")
        if registrar_auditoria:
            audit(st.session_state["email"], st.session_state["perfil"], "Substituiu base via Excel", "importacao_excel", "TODAS", "", nome_arquivo)
    total_api = total_anc = total_ae = 0

    if aba_api:
        df_api = xls.parse(aba_api)
        for _, r in df_api.iterrows():
            documento = normalizar_documento_sei(r.get("Documento SEI"))
            id_api = gerar_id_api(documento)
            if not id_api:
                continue
            processo_raw = safe_text(r.get("Processo SEI"))
            processo = formatar_processo_sei(processo_raw) or processo_raw
            ano_limpo = safe_int(r.get("Ano Concessao Limpo")) or safe_int(r.get("Ano Concessão"))
            ano_texto = formatar_ano_concessao(ano_limpo or r.get("Ano Concessão"))
            execute("""
                INSERT OR REPLACE INTO api
                (id_api, processo_sei, documento_sei, concessionaria, ano_concessao, ano_concessao_limpo,
                 disciplina, data_emissao, status_emissao, objeto, observacao, tipo_medida)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (id_api, processo, documento, safe_text(r.get("Concessionária")), ano_texto, ano_limpo,
                  safe_text(r.get("Disciplina")), safe_date(r.get("Data de Emissão")), safe_text(r.get("Status de Emissão")),
                  safe_text(r.get("Objeto")), safe_text(r.get("Observação")), "API"))
            total_api += 1

    if aba_anc:
        df_anc = xls.parse(aba_anc)
        for _, r in df_anc.iterrows():
            documento = normalizar_documento_sei(r.get("Documento SEI"))
            id_anc = gerar_id_anc(documento)
            if not id_anc:
                continue
            processo_raw = safe_text(r.get("Processo SEI"))
            processo = formatar_processo_sei(processo_raw) or processo_raw
            ano_limpo = safe_int(r.get("Ano Concessao Limpo")) or safe_int(r.get("Ano Concessão"))
            ano_texto = formatar_ano_concessao(ano_limpo or r.get("Ano Concessão"))
            execute("""
                INSERT OR REPLACE INTO anc
                (id_anc, processo_sei, documento_sei, concessionaria, ano_concessao, ano_concessao_limpo,
                 disciplina, data_emissao, data_recebimento_oficio, prazo_dias_uteis, prazo, status,
                 documento_sei_comprovacao, observacao_fiscalizacao, tipo_medida)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (id_anc, processo, documento, safe_text(r.get("Concessionária")), ano_texto, ano_limpo,
                  safe_text(r.get("Disciplina")), safe_date(r.get("Data de Emissão")), safe_date(r.get("Data de Recebimento do Ofício")),
                  safe_int(r.get("Prazo em Dias Úteis")), safe_date(r.get("Prazo")), safe_text(r.get("Status")),
                  normalizar_documento_sei(r.get("Documento SEI da Comprovação")), safe_text(r.get("Observação da Fiscalização")), "ANC"))
            total_anc += 1

    if aba_ae:
        df_ae = xls.parse(aba_ae)
        for _, r in df_ae.iterrows():
            documento = normalizar_documento_sei(r.get("Documento SEI Instaurador"))
            id_ae = gerar_id_ae(documento)
            if not id_ae:
                continue
            processo_raw = safe_text(r.get("Processo SEI"))
            processo = formatar_processo_sei(processo_raw) or processo_raw
            ano_limpo = safe_int(r.get("Ano Concessao Limpo")) or safe_int(r.get("Ano Concessão"))
            ano_texto = formatar_ano_concessao(ano_limpo or r.get("Ano Concessão"))
            execute("""
                INSERT OR REPLACE INTO ae
                (id_ae, processo_sei, documento_sei_instaurador, concessionaria, ano_concessao, ano_concessao_limpo,
                 disciplina, data_emissao, data_alinhamento_escopo, escopo, prazo_envio_cronograma, prazo_acao_educativa,
                 documento_sei_cronograma, documento_sei_acomp_1, documento_sei_acomp_2, documento_sei_acomp_3,
                 documento_sei_acomp_4, documento_sei_acomp_5, status_automatico, finalizada, observacao, tipo_medida)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (id_ae, processo, documento, safe_text(r.get("Concessionária")), ano_texto, ano_limpo,
                  safe_text(r.get("Disciplina")), safe_date(r.get("Data de Emissão")), safe_date(r.get("Data de Alinhamento do Escopo")),
                  safe_text(r.get("Escopo da Ação Educativa")), safe_date(r.get("Prazo para Envio do Cronograma")),
                  safe_date(r.get("Prazo da Ação Educativa")), normalizar_documento_sei(r.get("Documento SEI Cronograma")),
                  normalizar_documento_sei(r.get("Documento SEI Acompanhamento 1º Mês")),
                  normalizar_documento_sei(r.get("Documento SEI Acompanhamento 2º Mês")),
                  normalizar_documento_sei(r.get("Documento SEI Acompanhamento 3º Mês")),
                  normalizar_documento_sei(r.get("Documento SEI Acompanhamento 4º Mês")),
                  normalizar_documento_sei(r.get("Documento SEI Acompanhamento 5º Mês")),
                  safe_text(r.get("Status Automático")), safe_text(r.get("Finalizada?")), safe_text(r.get("Observação")), "Ação Educativa"))
            total_ae += 1

    if registrar_auditoria:
        audit(st.session_state["email"], st.session_state["perfil"], "Importou arquivo Excel", "importacao_excel", nome_arquivo, "", f"API={total_api}; ANC={total_anc}; AE={total_ae}")
    return total_api, total_anc, total_ae

# ============================================================
# Login e navegação
# ============================================================

def login_screen():
    c1, c2, c3 = st.columns([1, 1.25, 1])
    with c2:
        st.markdown(
            f"""
            <div style="text-align:center; padding: 22px 0 10px;">
                {antt_logo_html(330)}
                <div style="font-size:1.35rem; font-weight:850; color:#153125; margin-top:8px;">{APP_TITLE}</div>
                <div style="color:#5d7168; font-size:.92rem; margin-top:4px;">Acesso ao painel de apoio operacional e validação SEI</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        with st.form("login_form"):
            email = st.text_input("E-mail")
            senha = st.text_input("Senha", type="password")
            entrar = st.form_submit_button("Entrar", use_container_width=True)
    if entrar:
        usuario = query_one("SELECT email, perfil, concessionaria FROM usuarios WHERE email=? AND senha_hash=? AND ativo=1", (email, hash_password(senha)))
        if usuario:
            st.session_state["logged"] = True
            st.session_state["email"] = usuario[0]
            st.session_state["perfil"] = usuario[1]
            st.session_state["concessionaria"] = usuario[2]
            st.session_state["pagina_atual"] = "Dashboard Geral"
            st.rerun()
        else:
            st.error("Usuário ou senha inválidos.")


def logout_button():
    if st.sidebar.button("Sair"):
        st.session_state.clear()
        st.rerun()


def where_concessionaria():
    if st.session_state.get("perfil") == "Concessionária":
        return " WHERE concessionaria = ?", (st.session_state.get("concessionaria"),)
    return "", ()

# ============================================================
# Dashboard com avisos e gráficos navegáveis
# ============================================================

def pendencias_contadores():
    anc_p = query_one("SELECT COUNT(*) FROM anc_respostas WHERE status_resposta IN ('Aguardando validação ANTT','Complementação solicitada')")[0]
    ae_p = query_one("SELECT COUNT(*) FROM ae_acompanhamentos WHERE status_acompanhamento IN ('Aguardando validação ANTT','Complementação solicitada')")[0]
    return anc_p, ae_p


def plotly_bar_navegavel(df, x, y, title, key, color=None, custom_data=None):
    if px is None or df.empty:
        st.info("Instale plotly para gráficos navegáveis.")
        return None
    st.caption("Clique em uma barra para abrir a tabela correspondente com os filtros atuais.")
    color_map = None
    if color == "Tipo":
        color_map = TYPE_COLORS
    elif color == "Situação":
        color_map = STATUS_COLORS
    fig = px.bar(
        df,
        x=x,
        y=y,
        color=color,
        custom_data=custom_data,
        title=title,
        text=y,
        template=PLOTLY_TEMPLATE,
        color_discrete_map=color_map,
    )
    if color is None and title == "Medidas por tipo":
        fig.update_traces(marker_color=[SHORT_TYPE_COLORS.get(v, "#2563eb") for v in df[x].tolist()])
    fig.update_traces(textposition="outside", cliponaxis=False, marker_line_width=0)
    fig.update_layout(
        height=360,
        dragmode="select",
        bargap=0.28,
        margin=dict(l=18, r=18, t=52, b=42),
        title=dict(font=dict(size=17, color=POWERBI_COLORS["ink"]), x=0.02),
        paper_bgcolor="white",
        plot_bgcolor="white",
        font=dict(color=POWERBI_COLORS["ink"], family="Segoe UI, Arial"),
        xaxis=dict(title="", showgrid=False, tickfont=dict(size=11)),
        yaxis=dict(title="Quantidade", gridcolor="#e8eef7", zeroline=False),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        legend_title_text="",
    )
    event = st.plotly_chart(fig, use_container_width=True, on_select="rerun", selection_mode="points", key=key)
    try:
        points = event.selection.points
        if points:
            return points[0]
    except Exception:
        pass
    try:
        points = event["selection"]["points"]
        if points:
            return points[0]
    except Exception:
        pass
    return None


def dashboard():
    where, params = where_concessionaria()
    df_api = query_df(f"SELECT * FROM api {where}", params)
    df_anc = query_df(f"SELECT * FROM anc {where}", params)
    df_ae = query_df(f"SELECT * FROM ae {where}", params)

    base_filtros = pd.concat([
        df_api[["concessionaria", "disciplina", "ano_concessao_limpo", "data_emissao"]] if not df_api.empty else pd.DataFrame(columns=["concessionaria", "disciplina", "ano_concessao_limpo", "data_emissao"]),
        df_anc[["concessionaria", "disciplina", "ano_concessao_limpo", "data_emissao"]] if not df_anc.empty else pd.DataFrame(columns=["concessionaria", "disciplina", "ano_concessao_limpo", "data_emissao"]),
        df_ae[["concessionaria", "disciplina", "ano_concessao_limpo", "data_emissao"]] if not df_ae.empty else pd.DataFrame(columns=["concessionaria", "disciplina", "ano_concessao_limpo", "data_emissao"]),
    ], ignore_index=True)
    if "dash_filters_collapsed" not in st.session_state:
        st.session_state["dash_filters_collapsed"] = False
    with st.container(key="dash_right_filters"):
        if st.session_state["dash_filters_collapsed"]:
            st.markdown('<span class="antt-filter-collapsed"></span>', unsafe_allow_html=True)
            if st.button("Mostrar filtros", key="dash_expandir_filtros", use_container_width=True):
                st.session_state["dash_filters_collapsed"] = False
                st.rerun()
            conc, disc, ano_conc, ano_emissao, mes_emissao = filtros_medidas_estado("dash")
        else:
            c_filtro_titulo, c_filtro_acao = st.columns([0.58, 0.42], vertical_alignment="center")
            with c_filtro_titulo:
                st.markdown('<div class="antt-filter-title">Filtros</div>', unsafe_allow_html=True)
            with c_filtro_acao:
                recolher = st.button("Recolher", key="dash_recolher_filtros", use_container_width=True)
            if recolher:
                st.session_state["dash_filters_collapsed"] = True
                st.rerun()
            conc, disc, ano_conc, ano_emissao, mes_emissao = filtros_medidas(base_filtros, "dash", layout="vertical", mostrar_titulo=False)
    df_api = aplicar_filtros(df_api, conc, disc, ano_conc, ano_emissao, mes_emissao)
    df_anc = aplicar_filtros(df_anc, conc, disc, ano_conc, ano_emissao, mes_emissao)
    df_ae = aplicar_filtros(df_ae, conc, disc, ano_conc, ano_emissao, mes_emissao)

    if not df_anc.empty:
        df_anc["situacao_prazo"] = df_anc.apply(lambda r: situacao_prazo_anc(r["status"], r["prazo"]), axis=1)
    if not df_ae.empty:
        df_ae["situacao_prazo"] = df_ae.apply(lambda r: situacao_prazo_ae(r["finalizada"], r["prazo_acao_educativa"]), axis=1)

    total_api, total_anc, total_ae = len(df_api), len(df_anc), len(df_ae)
    total_medidas = total_api + total_anc + total_ae
    anc_vencidas = len(df_anc[df_anc["situacao_prazo"] == "Vencida"]) if not df_anc.empty else 0
    anc_a_vencer = len(df_anc[df_anc["situacao_prazo"] == "A vencer em 15 dias"]) if not df_anc.empty else 0
    ae_vencidas = len(df_ae[df_ae["situacao_prazo"] == "Vencida"]) if not df_ae.empty else 0
    anc_p, ae_p = pendencias_contadores()
    taxa_anc = round(len(df_anc[df_anc["status"] == "Cumprido"]) / total_anc * 100, 1) if total_anc and "status" in df_anc.columns else 0

    page_header(
        "Dashboard Geral",
        "Visão executiva das medidas, prazos críticos e pendências de validação.",
        [f"{total_medidas} medidas", f"{anc_p + ae_p} pendências ANTT"],
    )

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        kpi_card("Total", total_medidas, "Medidas filtradas", POWERBI_COLORS["accent"])
        botao_drillthrough("Ver tabela", "Tabela Geral", "geral", conc, disc, ano_conc, ano_emissao, mes_emissao, key="drill_total")
    with c2:
        kpi_card("API", total_api, "Alertas emitidos", POWERBI_COLORS["blue"])
        botao_drillthrough("Ver APIs", NOME_API, "api", conc, disc, ano_conc, ano_emissao, mes_emissao, key="drill_api")
    with c3:
        kpi_card("ANC", total_anc, f"{taxa_anc}% cumpridas", POWERBI_COLORS["accent"])
        botao_drillthrough("Ver ANCs", NOME_ANC, "anc", conc, disc, ano_conc, ano_emissao, mes_emissao, key="drill_anc")
    with c4:
        kpi_card("AE", total_ae, "Ações educativas", POWERBI_COLORS["green"])
        botao_drillthrough("Ver AEs", NOME_AE, "ae", conc, disc, ano_conc, ano_emissao, mes_emissao, key="drill_ae")
    with c5:
        kpi_card("Pendências", anc_p + ae_p, "Aguardando ANTT", POWERBI_COLORS["orange"])
        if st.button("Validar", key="drill_pendencias", use_container_width=True):
            ir_para("Pendências ANTT")

    c6, c7, c8 = st.columns(3)
    with c6:
        kpi_card("Críticas", anc_vencidas + ae_vencidas, "ANC/AE vencidas", POWERBI_COLORS["red"])
        b1, b2 = st.columns(2)
        with b1:
            botao_drillthrough("ANC", NOME_ANC, "anc", conc, disc, ano_conc, ano_emissao, mes_emissao, ["Vencida"], key="drill_crit_anc")
        with b2:
            botao_drillthrough("AE", NOME_AE, "ae", conc, disc, ano_conc, ano_emissao, mes_emissao, ["Vencida"], key="drill_crit_ae")
    with c7:
        kpi_card("ANCs a vencer", anc_a_vencer, "Próximos 15 dias", POWERBI_COLORS["orange"])
        botao_drillthrough("Ver tabela", NOME_ANC, "anc", conc, disc, ano_conc, ano_emissao, mes_emissao, ["A vencer em 15 dias"], key="drill_anc_a_vencer")
    with c8:
        kpi_card("AEs vencidas", ae_vencidas, "Ações educativas", POWERBI_COLORS["red"])
        botao_drillthrough("Ver tabela", NOME_AE, "ae", conc, disc, ano_conc, ano_emissao, mes_emissao, ["Vencida"], key="drill_ae_vencidas")

    if st.session_state.get("perfil") == "ANTT" and (anc_p + ae_p) > 0:
        st.markdown(f'<div class="antt-alert-card">Existem {anc_p + ae_p} pendências aguardando ação da ANTT.</div>', unsafe_allow_html=True)
        if st.button("Abrir Pendências ANTT"):
            ir_para("Pendências ANTT")

    st.markdown('<div class="antt-section-title">Análise visual</div>', unsafe_allow_html=True)
    st.caption("Passe o mouse sobre as barras para ver o resumo. Use os botões dos cards acima para abrir as planilhas com filtro.")
    col1, col2 = st.columns(2)
    with col1:
        medidas_tipo = pd.DataFrame({"Tipo": ["API", "ANC", "AE"], "Quantidade": [total_api, total_anc, total_ae]})
        grafico_barra_basico(medidas_tipo, "Tipo", "Quantidade", "Medidas por tipo", color_map=SHORT_TYPE_COLORS, key="graf_basico_tipo")
    with col2:
        if not df_anc.empty:
            sit = df_anc["situacao_prazo"].value_counts().reset_index()
            sit.columns = ["Situação", "Quantidade"]
            ordem = ["Vencida", "A vencer em 15 dias", "Dentro do prazo", "Cumprido", "Sem prazo informado"]
            sit["ordem"] = sit["Situação"].apply(lambda s: ordem.index(s) if s in ordem else 99)
            sit = sit.sort_values("ordem").drop(columns="ordem")
            grafico_barra_basico(sit, "Situação", "Quantidade", "ANCs por situação de prazo", color_map=STATUS_COLORS, key="graf_basico_situacao")
        else:
            st.info("Sem ANCs para exibir.")

    col3, col4 = st.columns(2)
    with col3:
        long = []
        if not df_api.empty:
            t = df_api.groupby("concessionaria").size().reset_index(name="Quantidade")
            t["Tipo"] = NOME_API
            long.append(t)
        if not df_anc.empty:
            t = df_anc.groupby("concessionaria").size().reset_index(name="Quantidade")
            t["Tipo"] = NOME_ANC
            long.append(t)
        if not df_ae.empty:
            t = df_ae.groupby("concessionaria").size().reset_index(name="Quantidade")
            t["Tipo"] = NOME_AE
            long.append(t)
        if long:
            medida_conc = pd.concat(long, ignore_index=True).rename(columns={"concessionaria": "Concessionária"})
            grafico_barra_basico(medida_conc, "Concessionária", "Quantidade", "Medidas por concessionária", color="Tipo", color_map=TYPE_COLORS, key="graf_basico_concessionaria")
        else:
            st.info("Sem dados para exibir.")
    with col4:
        if not df_anc.empty:
            disc_df = df_anc.groupby("disciplina").size().sort_values(ascending=False).head(12).reset_index(name="Quantidade")
            disc_df = disc_df.rename(columns={"disciplina": "Disciplina"})
            grafico_barra_basico(disc_df, "Disciplina", "Quantidade", "Top disciplinas em ANC", key="graf_basico_disciplina")
        else:
            st.info("Sem ANCs para exibir.")

# ============================================================
# Cadastro
# ============================================================

def tela_cadastrar_medida():
    page_header("Cadastrar Medida", "Crie APIs, ANCs e Ações Educativas a partir de um fluxo único.", ["ANTT", "Novo registro"])
    if st.session_state["perfil"] != "ANTT":
        st.warning("Acesso restrito aos usuários da ANTT.")
        return
    page_actions([
        {"label": "Ver dashboard", "key": "cad_dash", "callback": lambda: ir_para("Dashboard Geral")},
        {"label": "Ver tabela geral", "key": "cad_geral", "callback": lambda: ir_para("Tabela Geral")},
        {"label": "Importar Excel", "key": "cad_importar", "callback": lambda: ir_para("Importar Excel")},
    ])
    tipo_padrao = st.session_state.pop("tipo_cadastro_padrao", NOME_API)
    opcoes_tipo = [NOME_API, NOME_ANC, NOME_AE]
    tipo = st.selectbox("Selecione o tipo de medida", opcoes_tipo, index=opcoes_tipo.index(tipo_padrao) if tipo_padrao in opcoes_tipo else 0)
    st.divider()
    if tipo == NOME_API:
        form_cadastro_api()
    elif tipo == NOME_ANC:
        form_cadastro_anc()
    else:
        form_cadastro_ae()


def form_cadastro_api():
    st.subheader(f"Cadastrar {NOME_API}")
    with st.form("form_api"):
        proc_digit, processo = campo_processo_sei("api_proc")
        _, documento = campo_documento_sei("Documento SEI", "api_doc", "", "Ex: 32186480")
        concessionaria = st.selectbox("Concessionária", CONCESSIONARIAS)
        ano = st.number_input("Ano da Concessão", min_value=1, max_value=99, value=1, step=1)
        disciplina = st.selectbox("Disciplina", DISCIPLINAS)
        data_emissao = date_input_br("Data de Emissão")
        status = st.selectbox("Status de Emissão", STATUS_API)
        objeto = st.text_area("Objeto")
        observacao = st.text_area("Observação")
        salvar = st.form_submit_button("Salvar API")
    if salvar:
        erros = []
        if not validar_processo_sei(proc_digit): erros.append("Processo SEI deve conter 17 números.")
        if not validar_documento_sei(documento): erros.append("Documento SEI deve conter somente números.")
        if not objeto.strip(): erros.append("Objeto é obrigatório.")
        if erros:
            for e in erros: st.error(e)
            return
        id_api = gerar_id_api(documento)
        try:
            execute("""
                INSERT INTO api (id_api, processo_sei, documento_sei, concessionaria, ano_concessao, ano_concessao_limpo, disciplina, data_emissao, status_emissao, objeto, observacao, tipo_medida)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (id_api, processo, documento, concessionaria, formatar_ano_concessao(ano), int(ano), disciplina, data_emissao.isoformat(), status, objeto, observacao, "API"))
            audit(st.session_state["email"], st.session_state["perfil"], "Cadastrou nova API", "api", id_api, "", documento)
            limpar_selecao()
            st.success(f"API cadastrada com sucesso: {id_api}")
        except sqlite3.IntegrityError:
            st.error("Já existe uma API com esse Documento SEI.")
        except Exception as e:
            st.error(f"Erro ao cadastrar API: {e}")


def form_cadastro_anc():
    st.subheader(f"Cadastrar {NOME_ANC}")
    with st.form("form_anc"):
        proc_digit, processo = campo_processo_sei("anc_proc")
        _, documento = campo_documento_sei("Documento SEI", "anc_doc", "", "Ex: 32717059")
        concessionaria = st.selectbox("Concessionária", CONCESSIONARIAS)
        ano = st.number_input("Ano da Concessão", min_value=1, max_value=99, value=1, step=1)
        disciplina = st.selectbox("Disciplina", DISCIPLINAS)
        data_emissao = date_input_br("Data de Emissão")
        data_recebimento = date_input_br("Data de Recebimento do Ofício")
        prazo_dias = st.number_input("Prazo em Dias Úteis", min_value=1, max_value=180, value=15, step=1)
        prazo = date_input_br("Prazo Final")
        status = st.selectbox("Status", STATUS_ANC)
        _, doc_comp = campo_documento_sei("Documento SEI da Comprovação, se já houver", "anc_doc_comp", "", "Somente números")
        obs = st.text_area("Observação da Fiscalização")
        salvar = st.form_submit_button("Salvar ANC")
    if salvar:
        erros = []
        if not validar_processo_sei(proc_digit): erros.append("Processo SEI deve conter 17 números.")
        if not validar_documento_sei(documento): erros.append("Documento SEI da ANC deve conter somente números.")
        if doc_comp and not validar_documento_sei(doc_comp): erros.append("Documento SEI da comprovação deve conter somente números.")
        if prazo < data_recebimento: erros.append("Prazo final não pode ser anterior à data de recebimento.")
        if not obs.strip(): erros.append("Observação da fiscalização é obrigatória.")
        if erros:
            for e in erros: st.error(e)
            return
        id_anc = gerar_id_anc(documento)
        try:
            execute("""
                INSERT INTO anc (id_anc, processo_sei, documento_sei, concessionaria, ano_concessao, ano_concessao_limpo, disciplina, data_emissao, data_recebimento_oficio, prazo_dias_uteis, prazo, status, documento_sei_comprovacao, observacao_fiscalizacao, tipo_medida)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (id_anc, processo, documento, concessionaria, formatar_ano_concessao(ano), int(ano), disciplina, data_emissao.isoformat(), data_recebimento.isoformat(), int(prazo_dias), prazo.isoformat(), status, doc_comp, obs, "ANC"))
            audit(st.session_state["email"], st.session_state["perfil"], "Cadastrou nova ANC", "anc", id_anc, "", documento)
            limpar_selecao()
            st.success(f"ANC cadastrada com sucesso: {id_anc}")
        except sqlite3.IntegrityError:
            st.error("Já existe uma ANC com esse Documento SEI.")
        except Exception as e:
            st.error(f"Erro ao cadastrar ANC: {e}")


def form_cadastro_ae():
    st.subheader(f"Cadastrar {NOME_AE}")
    with st.form("form_ae"):
        proc_digit, processo = campo_processo_sei("ae_proc")
        _, documento = campo_documento_sei("Documento SEI Instaurador", "ae_doc", "", "Ex: 33333333")
        concessionaria = st.selectbox("Concessionária", CONCESSIONARIAS)
        ano = st.number_input("Ano da Concessão", min_value=1, max_value=99, value=1, step=1)
        disciplina = st.selectbox("Disciplina", DISCIPLINAS)
        data_emissao = date_input_br("Data de Emissão")
        data_alinhamento = date_input_br("Data de Alinhamento do Escopo")
        escopo = st.text_area("Escopo da Ação Educativa")
        prazo_cronograma = date_input_br("Prazo para Envio do Cronograma")
        prazo_ae = date_input_br("Prazo da Ação Educativa")
        _, doc_cron = campo_documento_sei("Documento SEI Cronograma, se já houver", "ae_doc_cron", "", "Somente números")
        status = st.selectbox("Status Automático", STATUS_AE)
        finalizada = st.selectbox("Finalizada?", SIM_NAO)
        obs = st.text_area("Observação")
        salvar = st.form_submit_button("Salvar Ação Educativa")
    if salvar:
        erros = []
        if not validar_processo_sei(proc_digit): erros.append("Processo SEI deve conter 17 números.")
        if not validar_documento_sei(documento): erros.append("Documento SEI Instaurador deve conter somente números.")
        if doc_cron and not validar_documento_sei(doc_cron): erros.append("Documento SEI Cronograma deve conter somente números.")
        if prazo_ae < data_emissao: erros.append("Prazo da AE não pode ser anterior à data de emissão.")
        if not escopo.strip(): erros.append("Escopo é obrigatório.")
        if erros:
            for e in erros: st.error(e)
            return
        id_ae = gerar_id_ae(documento)
        try:
            execute("""
                INSERT INTO ae (id_ae, processo_sei, documento_sei_instaurador, concessionaria, ano_concessao, ano_concessao_limpo, disciplina, data_emissao, data_alinhamento_escopo, escopo, prazo_envio_cronograma, prazo_acao_educativa, documento_sei_cronograma, documento_sei_acomp_1, documento_sei_acomp_2, documento_sei_acomp_3, documento_sei_acomp_4, documento_sei_acomp_5, status_automatico, finalizada, observacao, tipo_medida)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (id_ae, processo, documento, concessionaria, formatar_ano_concessao(ano), int(ano), disciplina, data_emissao.isoformat(), data_alinhamento.isoformat(), escopo, prazo_cronograma.isoformat(), prazo_ae.isoformat(), doc_cron, "", "", "", "", "", status, finalizada, obs, "Ação Educativa"))
            audit(st.session_state["email"], st.session_state["perfil"], "Cadastrou nova AE", "ae", id_ae, "", documento)
            limpar_selecao()
            st.success(f"Ação Educativa cadastrada com sucesso: {id_ae}")
        except sqlite3.IntegrityError:
            st.error("Já existe uma AE com esse Documento SEI Instaurador.")
        except Exception as e:
            st.error(f"Erro ao cadastrar AE: {e}")

# ============================================================
# Telas de medidas
# ============================================================

def tela_api():
    page_header(NOME_API, "Consulte, exporte e mantenha os alertas registrados.", ["Tabela", "Detalhe por clique"])
    if st.session_state.get("perfil") == "ANTT":
        if st.button("Cadastrar nova API", key="nova_api", use_container_width=False):
            st.session_state["tipo_cadastro_padrao"] = NOME_API
            ir_para("Cadastrar Medida")
    where, params = where_concessionaria()
    df = query_df(f"SELECT * FROM api {where} ORDER BY data_emissao DESC", params)
    if df.empty:
        st.info("Nenhuma API encontrada.")
        return
    with st.expander("Filtros da tabela", expanded=True):
        conc, disc, ano_conc, ano_emissao, mes_emissao = filtros_medidas(df, "api")
    df = aplicar_filtros(df, conc, disc, ano_conc, ano_emissao, mes_emissao)
    resumo_lista(df, "APIs")
    contexto_tabela(df, "Registros de API")
    selected = tabela_interativa(preparar_tabela_api(df), key_grid("grid_api"))
    st.download_button("Exportar APIs em CSV", df.to_csv(index=False).encode("utf-8-sig"), "apis.csv", "text/csv")
    if not selected:
        aviso_selecao("Selecione uma linha para visualizar detalhes, editar ou excluir a API.")
        return
    id_api = id_selecionado(selected, "ID da API", "id_api")
    item = df[df["id_api"] == id_api].iloc[0]
    st.divider()
    st.subheader("Detalhes da medida selecionada")
    st.write("**ID:**", item["id_api"])
    st.write("**Processo SEI:**", item["processo_sei"])
    st.write("**Documento SEI:**", item["documento_sei"])
    st.write("**Concessionária:**", item["concessionaria"])
    st.write("**Disciplina:**", item["disciplina"])
    st.write("**Data de Emissão:**", formatar_data_br(item["data_emissao"]))
    st.write("**Status:**", item["status_emissao"])
    st.write("**Objeto:**")
    st.info(item["objeto"])
    st.write("**Observação:**")
    st.info(item["observacao"])
    if st.session_state["perfil"] == "ANTT":
        st.divider()
        st.subheader("Ações da ANTT")
        with st.expander("Editar API", expanded=False):
            with st.form("editar_api"):
                proc_digit, processo = campo_processo_sei("edit_api_proc", item["processo_sei"])
                _, documento = campo_documento_sei("Documento SEI", "edit_api_doc", item["documento_sei"], "Somente números")
                concessionaria = st.selectbox("Concessionária", CONCESSIONARIAS, index=CONCESSIONARIAS.index(item["concessionaria"]) if item["concessionaria"] in CONCESSIONARIAS else 0, key="edit_api_conc")
                ano_valor = int(item["ano_concessao_limpo"]) if safe_text(item["ano_concessao_limpo"]) else 1
                ano_edit = st.number_input("Ano da Concessão", min_value=1, max_value=99, value=ano_valor, step=1, key="edit_api_ano")
                disciplina = st.selectbox("Disciplina", DISCIPLINAS, index=DISCIPLINAS.index(item["disciplina"]) if item["disciplina"] in DISCIPLINAS else 0, key="edit_api_disc")
                data_emissao = date_input_br("Data de Emissão", value=parse_date(item["data_emissao"]) or date.today(), key="edit_api_data")
                status = st.selectbox("Status de Emissão", STATUS_API, index=STATUS_API.index(item["status_emissao"]) if item["status_emissao"] in STATUS_API else 0, key="edit_api_status")
                objeto = st.text_area("Objeto", value=item["objeto"], key="edit_api_obj")
                observacao = st.text_area("Observação", value=item["observacao"], key="edit_api_obs")
                salvar = st.form_submit_button("Salvar alterações")
            if salvar:
                erros = []
                if not validar_processo_sei(proc_digit): erros.append("Processo SEI deve conter 17 números.")
                if not validar_documento_sei(documento): erros.append("Documento SEI deve conter somente números.")
                if not objeto.strip(): erros.append("Objeto é obrigatório.")
                if erros:
                    for e in erros: st.error(e)
                    return
                novo_id = gerar_id_api(documento)
                try:
                    execute("""
                        UPDATE api SET id_api=?, processo_sei=?, documento_sei=?, concessionaria=?, ano_concessao=?, ano_concessao_limpo=?, disciplina=?, data_emissao=?, status_emissao=?, objeto=?, observacao=? WHERE id_api=?
                    """, (novo_id, processo, documento, concessionaria, formatar_ano_concessao(ano_edit), int(ano_edit), disciplina, data_emissao.isoformat(), status, objeto, observacao, id_api))
                    audit(st.session_state["email"], st.session_state["perfil"], "Editou API", "api", id_api, str(item.to_dict()), novo_id)
                    limpar_selecao()
                    st.success("API atualizada com sucesso.")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("Já existe outra API com esse Documento SEI.")
                except Exception as e:
                    st.error(f"Erro ao atualizar API: {e}")
        with st.expander("Excluir API", expanded=False):
            st.warning("Esta ação não pode ser desfeita.")
            confirmar = st.checkbox(f"Confirmo que desejo excluir a API {id_api}", key="conf_del_api")
            texto = st.text_input("Digite EXCLUIR para confirmar", key="txt_del_api")
            if st.button("Excluir API"):
                if confirmar and texto == "EXCLUIR":
                    execute("DELETE FROM api WHERE id_api=?", (id_api,))
                    audit(st.session_state["email"], st.session_state["perfil"], "Excluiu API", "api", id_api, str(item.to_dict()), "")
                    limpar_selecao()
                    st.success("API excluída com sucesso.")
                    st.rerun()
                else:
                    st.error("Confirmação inválida. Marque a caixa e digite EXCLUIR.")


def tela_anc():
    page_header(NOME_ANC, "Acompanhe prazo, cumprimento, respostas e validação de não conformidades.", ["Prazos", "Validação"])
    if st.session_state.get("perfil") == "ANTT":
        if st.button("Cadastrar nova ANC", key="nova_anc", use_container_width=False):
            st.session_state["tipo_cadastro_padrao"] = NOME_ANC
            ir_para("Cadastrar Medida")
    where, params = where_concessionaria()
    df = query_df(f"SELECT * FROM anc {where} ORDER BY prazo ASC", params)
    if df.empty:
        st.info("Nenhuma ANC encontrada.")
        return
    df["situacao_prazo"] = df.apply(lambda r: situacao_prazo_anc(r["status"], r["prazo"]), axis=1)
    with st.expander("Filtros da tabela", expanded=True):
        conc, disc, ano_conc, ano_emissao, mes_emissao = filtros_medidas(df, "anc")
        df = aplicar_filtros(df, conc, disc, ano_conc, ano_emissao, mes_emissao)
        df = filtro_situacao_prazo(df, "anc", ["Vencida", "A vencer em 15 dias", "Dentro do prazo", "Cumprido", "Sem prazo informado"])
    resumo_lista(df, "ANCs")
    contexto_tabela(df, "Registros de ANC")
    selected = tabela_interativa(preparar_tabela_anc(df), key_grid("grid_anc"))
    st.download_button("Exportar ANCs em CSV", df.to_csv(index=False).encode("utf-8-sig"), "ancs.csv", "text/csv")
    if not selected:
        aviso_selecao("Selecione uma linha para visualizar detalhes e ações disponíveis.")
        return
    id_anc = id_selecionado(selected, "ID da ANC", "id_anc")
    anc = df[df["id_anc"] == id_anc].iloc[0]
    st.divider()
    st.subheader("Detalhes da medida selecionada")
    st.write("**ID:**", anc["id_anc"])
    st.write("**Processo SEI:**", anc["processo_sei"])
    st.write("**Documento SEI:**", anc["documento_sei"])
    st.write("**Concessionária:**", anc["concessionaria"])
    st.write("**Disciplina:**", anc["disciplina"])
    st.write("**Prazo:**", formatar_data_br(anc["prazo"]))
    st.write("**Status:**", anc["status"])
    st.write("**Situação:**", anc["situacao_prazo"])
    st.write("**Observação da Fiscalização:**")
    st.info(anc["observacao_fiscalizacao"])
    perfil = st.session_state["perfil"]
    if perfil == "Concessionária":
        st.divider()
        st.subheader("Ações disponíveis para a concessionária")
        with st.form("form_resposta_anc"):
            _, documento = campo_documento_sei("Documento SEI da comprovação (obrigatório)", "resp_anc_doc", "", "Somente números")
            observacao = st.text_area("Observação da concessionária")
            arquivo = st.file_uploader("Anexar documento ou evidência (PDF ou HTML)", type=EXTENSOES_UPLOAD)
            enviar = st.form_submit_button("Enviar cumprimento para validação ANTT")
        if enviar:
            if not validar_documento_sei(documento):
                st.error("Informe obrigatoriamente o número SEI da comprovação, usando somente números.")
                return
            try:
                arquivo_path = salvar_upload(arquivo, id_anc)
            except ValueError as e:
                st.error(str(e))
                return
            execute("""
                INSERT INTO anc_respostas (id_anc, concessionaria, data_resposta, usuario_responsavel, documento_sei_comprovacao, observacao_concessionaria, arquivo_link, status_resposta)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (id_anc, anc["concessionaria"], datetime.now().isoformat(timespec="seconds"), st.session_state["email"], documento, observacao, arquivo_path, "Aguardando validação ANTT"))
            audit(st.session_state["email"], perfil, "Enviou resposta ANC", "anc_respostas", id_anc, "", documento)
            limpar_selecao()
            st.success("Resposta enviada para validação da ANTT.")
            st.rerun()
    elif perfil == "ANTT":
        respostas = query_df("SELECT * FROM anc_respostas WHERE id_anc=? ORDER BY data_resposta DESC", (id_anc,))
        st.subheader("Respostas da concessionária")
        selected_resp = tabela_interativa(respostas, key_grid("grid_respostas_anc"), height=260) if not respostas.empty else None
        if selected_resp:
            mostrar_arquivo_upload(selected_resp.get("arquivo_link", ""), f"Resposta ANC {selected_resp.get('id')}")
        elif respostas.empty:
            st.info("Nenhuma resposta enviada.")
        validar_respostas_anc(respostas, id_anc, perfil)
        st.divider()
        st.subheader("Ações da ANTT")
        editar_excluir_anc(anc, id_anc, perfil)


def validar_respostas_anc(respostas, id_anc, perfil):
    pendentes = respostas[respostas["status_resposta"].isin(["Aguardando validação ANTT", "Complementação solicitada"])] if not respostas.empty else pd.DataFrame()
    if pendentes.empty:
        return
    st.subheader("Validar resposta da concessionária")
    resposta_id = st.selectbox("Selecione uma resposta pendente", pendentes["id"].tolist(), key=f"val_resp_anc_{id_anc}")
    resposta = pendentes[pendentes["id"] == resposta_id].iloc[0]
    with st.form(f"validar_anc_{id_anc}"):
        decisao = st.selectbox("Decisão ANTT", STATUS_VALIDACAO_ANC)
        obs_val = st.text_area("Observação da validação ANTT")
        validar = st.form_submit_button("Registrar validação")
    if validar:
        execute("""
            UPDATE anc_respostas SET status_resposta=?, data_validacao_antt=?, usuario_antt_validador=?, observacao_validacao_antt=? WHERE id=?
        """, (decisao, datetime.now().isoformat(timespec="seconds"), st.session_state["email"], obs_val, int(resposta_id)))
        if decisao == "Aceita":
            execute("UPDATE anc SET status='Cumprido', documento_sei_comprovacao=? WHERE id_anc=?", (resposta["documento_sei_comprovacao"], id_anc))
        audit(st.session_state["email"], perfil, f"Validou resposta ANC como {decisao}", "anc_respostas", id_anc, "", obs_val)
        limpar_selecao()
        st.success("Validação registrada.")
        st.rerun()


def editar_excluir_anc(anc, id_anc, perfil):
    with st.expander("Editar ANC", expanded=False):
        with st.form("editar_anc"):
            proc_digit, processo = campo_processo_sei("edit_anc_proc", anc["processo_sei"])
            _, documento = campo_documento_sei("Documento SEI", "edit_anc_doc", anc["documento_sei"], "Somente números")
            concessionaria = st.selectbox("Concessionária", CONCESSIONARIAS, index=CONCESSIONARIAS.index(anc["concessionaria"]) if anc["concessionaria"] in CONCESSIONARIAS else 0, key="edit_anc_conc")
            ano_valor = int(anc["ano_concessao_limpo"]) if safe_text(anc["ano_concessao_limpo"]) else 1
            ano_edit = st.number_input("Ano da Concessão", min_value=1, max_value=99, value=ano_valor, step=1, key="edit_anc_ano")
            disciplina = st.selectbox("Disciplina", DISCIPLINAS, index=DISCIPLINAS.index(anc["disciplina"]) if anc["disciplina"] in DISCIPLINAS else 0, key="edit_anc_disc")
            data_emissao = date_input_br("Data de Emissão", value=parse_date(anc["data_emissao"]) or date.today(), key="edit_anc_data_em")
            data_rec = date_input_br("Data de Recebimento do Ofício", value=parse_date(anc["data_recebimento_oficio"]) or date.today(), key="edit_anc_data_rec")
            prazo_dias = st.number_input("Prazo em Dias Úteis", min_value=1, max_value=180, value=int(anc["prazo_dias_uteis"]) if safe_text(anc["prazo_dias_uteis"]) else 15, step=1, key="edit_anc_prazo_dias")
            prazo = date_input_br("Prazo Final", value=parse_date(anc["prazo"]) or date.today(), key="edit_anc_prazo")
            status = st.selectbox("Status", STATUS_ANC, index=STATUS_ANC.index(anc["status"]) if anc["status"] in STATUS_ANC else 0, key="edit_anc_status")
            _, doc_comp = campo_documento_sei("Documento SEI da Comprovação", "edit_anc_doc_comp", anc["documento_sei_comprovacao"] or "", "Somente números")
            obs_edit = st.text_area("Observação da Fiscalização", value=anc["observacao_fiscalizacao"], key="edit_anc_obs")
            salvar = st.form_submit_button("Salvar alterações")
        if salvar:
            erros = []
            if not validar_processo_sei(proc_digit): erros.append("Processo SEI deve conter 17 números.")
            if not validar_documento_sei(documento): erros.append("Documento SEI deve conter somente números.")
            if doc_comp and not validar_documento_sei(doc_comp): erros.append("Documento SEI da comprovação deve conter somente números.")
            if prazo < data_rec: erros.append("Prazo final não pode ser anterior à data de recebimento.")
            if not obs_edit.strip(): erros.append("Observação da fiscalização é obrigatória.")
            if erros:
                for e in erros: st.error(e)
                return
            novo_id = gerar_id_anc(documento)
            try:
                execute("""
                    UPDATE anc SET id_anc=?, processo_sei=?, documento_sei=?, concessionaria=?, ano_concessao=?, ano_concessao_limpo=?, disciplina=?, data_emissao=?, data_recebimento_oficio=?, prazo_dias_uteis=?, prazo=?, status=?, documento_sei_comprovacao=?, observacao_fiscalizacao=? WHERE id_anc=?
                """, (novo_id, processo, documento, concessionaria, formatar_ano_concessao(ano_edit), int(ano_edit), disciplina, data_emissao.isoformat(), data_rec.isoformat(), int(prazo_dias), prazo.isoformat(), status, doc_comp, obs_edit, id_anc))
                execute("UPDATE anc_respostas SET id_anc=? WHERE id_anc=?", (novo_id, id_anc))
                audit(st.session_state["email"], perfil, "Editou ANC", "anc", id_anc, str(anc.to_dict()), novo_id)
                limpar_selecao()
                st.success("ANC atualizada com sucesso.")
                st.rerun()
            except sqlite3.IntegrityError:
                st.error("Já existe outra ANC com esse Documento SEI.")
            except Exception as e:
                st.error(f"Erro ao atualizar ANC: {e}")
    with st.expander("Excluir ANC", expanded=False):
        st.warning("Esta ação apagará também respostas vinculadas à ANC.")
        confirmar = st.checkbox(f"Confirmo que desejo excluir a ANC {id_anc}", key="conf_del_anc")
        texto = st.text_input("Digite EXCLUIR para confirmar", key="txt_del_anc")
        if st.button("Excluir ANC"):
            if confirmar and texto == "EXCLUIR":
                execute("DELETE FROM anc_respostas WHERE id_anc=?", (id_anc,))
                execute("DELETE FROM anc WHERE id_anc=?", (id_anc,))
                audit(st.session_state["email"], perfil, "Excluiu ANC", "anc", id_anc, str(anc.to_dict()), "")
                limpar_selecao()
                st.success("ANC excluída com sucesso.")
                st.rerun()
            else:
                st.error("Confirmação inválida. Marque a caixa e digite EXCLUIR.")


def tela_ae():
    page_header(NOME_AE, "Gerencie ações educativas, cronogramas e acompanhamentos enviados.", ["Acompanhamento", "Cronograma"])
    if st.session_state.get("perfil") == "ANTT":
        if st.button("Cadastrar nova AE", key="nova_ae", use_container_width=False):
            st.session_state["tipo_cadastro_padrao"] = NOME_AE
            ir_para("Cadastrar Medida")
    where, params = where_concessionaria()
    df = query_df(f"SELECT * FROM ae {where} ORDER BY prazo_acao_educativa ASC", params)
    if df.empty:
        st.info("Nenhuma Ação Educativa encontrada.")
        return
    df["situacao_prazo"] = df.apply(lambda r: situacao_prazo_ae(r["finalizada"], r["prazo_acao_educativa"]), axis=1)
    with st.expander("Filtros da tabela", expanded=True):
        conc, disc, ano_conc, ano_emissao, mes_emissao = filtros_medidas(df, "ae")
        df = aplicar_filtros(df, conc, disc, ano_conc, ano_emissao, mes_emissao)
        df = filtro_situacao_prazo(df, "ae", ["Vencida", "A vencer em 30 dias", "Dentro do prazo", "Finalizada", "Sem prazo informado"])
    resumo_lista(df, "AEs")
    contexto_tabela(df, "Registros de AE")
    selected = tabela_interativa(preparar_tabela_ae(df), key_grid("grid_ae"))
    st.download_button("Exportar AEs em CSV", df.to_csv(index=False).encode("utf-8-sig"), "acoes_educativas.csv", "text/csv")
    if not selected:
        aviso_selecao("Selecione uma linha para visualizar detalhes e ações disponíveis.")
        return
    id_ae = id_selecionado(selected, "ID da AE", "id_ae")
    ae = df[df["id_ae"] == id_ae].iloc[0]
    st.divider()
    st.subheader("Detalhes da medida selecionada")
    st.write("**ID:**", ae["id_ae"])
    st.write("**Processo SEI:**", ae["processo_sei"])
    st.write("**Documento SEI Instaurador:**", ae["documento_sei_instaurador"])
    st.write("**Concessionária:**", ae["concessionaria"])
    st.write("**Disciplina:**", ae["disciplina"])
    st.write("**Prazo da AE:**", formatar_data_br(ae["prazo_acao_educativa"]))
    st.write("**Finalizada?:**", ae["finalizada"])
    st.write("**Situação:**", ae["situacao_prazo"])
    st.write("**Escopo:**")
    st.info(ae["escopo"])
    perfil = st.session_state["perfil"]
    if perfil == "Concessionária":
        st.divider()
        st.subheader("Ações disponíveis para a concessionária")
        with st.form("form_acomp_ae"):
            mes = st.selectbox("Mês de acompanhamento", [1, 2, 3, 4, 5])
            _, documento = campo_documento_sei("Documento SEI do acompanhamento (obrigatório)", "acomp_ae_doc", "", "Somente números")
            observacao = st.text_area("Observação da concessionária")
            arquivo = st.file_uploader("Anexar documento ou evidência (PDF ou HTML)", type=EXTENSOES_UPLOAD)
            enviar = st.form_submit_button("Enviar acompanhamento")
        if enviar:
            if not validar_documento_sei(documento):
                st.error("Informe obrigatoriamente o número SEI do acompanhamento, usando somente números.")
                return
            try:
                arquivo_path = salvar_upload(arquivo, id_ae)
            except ValueError as e:
                st.error(str(e))
                return
            execute("""
                INSERT INTO ae_acompanhamentos (id_ae, concessionaria, mes_acompanhamento, data_envio, usuario_responsavel, documento_sei_acompanhamento, observacao_concessionaria, arquivo_link, status_acompanhamento)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (id_ae, ae["concessionaria"], int(mes), datetime.now().isoformat(timespec="seconds"), st.session_state["email"], documento, observacao, arquivo_path, "Aguardando validação ANTT"))
            audit(st.session_state["email"], perfil, "Enviou acompanhamento AE", "ae_acompanhamentos", id_ae, "", documento)
            limpar_selecao()
            st.success("Acompanhamento enviado para validação da ANTT.")
            st.rerun()
    elif perfil == "ANTT":
        acompanhamentos = query_df("SELECT * FROM ae_acompanhamentos WHERE id_ae=? ORDER BY data_envio DESC", (id_ae,))
        st.subheader("Acompanhamentos enviados")
        selected_acomp = tabela_interativa(acompanhamentos, key_grid("grid_acomp_ae"), height=260) if not acompanhamentos.empty else None
        if selected_acomp:
            mostrar_arquivo_upload(selected_acomp.get("arquivo_link", ""), f"Acompanhamento AE {selected_acomp.get('id')}")
        elif acompanhamentos.empty:
            st.info("Nenhum acompanhamento enviado.")
        validar_acompanhamentos_ae(acompanhamentos, id_ae, perfil)
        st.divider()
        st.subheader("Ações da ANTT")
        editar_excluir_ae(ae, id_ae, perfil)


def validar_acompanhamentos_ae(acompanhamentos, id_ae, perfil):
    pendentes = acompanhamentos[acompanhamentos["status_acompanhamento"].isin(["Aguardando validação ANTT", "Complementação solicitada"])] if not acompanhamentos.empty else pd.DataFrame()
    if pendentes.empty:
        return
    st.subheader("Validar acompanhamento da concessionária")
    acomp_id = st.selectbox("Selecione um acompanhamento pendente", pendentes["id"].tolist(), key=f"val_acomp_ae_{id_ae}")
    acomp = pendentes[pendentes["id"] == acomp_id].iloc[0]
    with st.form(f"validar_acomp_ae_{id_ae}"):
        decisao = st.selectbox("Decisão ANTT", STATUS_VALIDACAO_AE)
        obs_val = st.text_area("Observação da validação ANTT")
        validar = st.form_submit_button("Registrar validação")
    if validar:
        execute("""
            UPDATE ae_acompanhamentos SET status_acompanhamento=?, data_validacao_antt=?, usuario_antt_validador=?, observacao_validacao_antt=? WHERE id=?
        """, (decisao, datetime.now().isoformat(timespec="seconds"), st.session_state["email"], obs_val, int(acomp_id)))
        if decisao == "Aceito":
            mes = int(acomp["mes_acompanhamento"])
            if mes in [1, 2, 3, 4, 5]:
                execute(f"UPDATE ae SET documento_sei_acomp_{mes}=? WHERE id_ae=?", (acomp["documento_sei_acompanhamento"], id_ae))
        audit(st.session_state["email"], perfil, f"Validou acompanhamento AE como {decisao}", "ae_acompanhamentos", id_ae, "", obs_val)
        limpar_selecao()
        st.success("Validação registrada.")
        st.rerun()


def editar_excluir_ae(ae, id_ae, perfil):
    with st.expander("Editar Ação Educativa", expanded=False):
        with st.form("editar_ae"):
            proc_digit, processo = campo_processo_sei("edit_ae_proc", ae["processo_sei"])
            _, documento = campo_documento_sei("Documento SEI Instaurador", "edit_ae_doc", ae["documento_sei_instaurador"], "Somente números")
            concessionaria = st.selectbox("Concessionária", CONCESSIONARIAS, index=CONCESSIONARIAS.index(ae["concessionaria"]) if ae["concessionaria"] in CONCESSIONARIAS else 0, key="edit_ae_conc")
            ano_valor = int(ae["ano_concessao_limpo"]) if safe_text(ae["ano_concessao_limpo"]) else 1
            ano_edit = st.number_input("Ano da Concessão", min_value=1, max_value=99, value=ano_valor, step=1, key="edit_ae_ano")
            disciplina = st.selectbox("Disciplina", DISCIPLINAS, index=DISCIPLINAS.index(ae["disciplina"]) if ae["disciplina"] in DISCIPLINAS else 0, key="edit_ae_disc")
            data_emissao = date_input_br("Data de Emissão", value=parse_date(ae["data_emissao"]) or date.today(), key="edit_ae_data_em")
            data_alinhamento = date_input_br("Data de Alinhamento do Escopo", value=parse_date(ae["data_alinhamento_escopo"]) or date.today(), key="edit_ae_data_al")
            escopo = st.text_area("Escopo da Ação Educativa", value=ae["escopo"], key="edit_ae_escopo")
            prazo_cron = date_input_br("Prazo para Envio do Cronograma", value=parse_date(ae["prazo_envio_cronograma"]) or date.today(), key="edit_ae_prazo_cron")
            prazo_ae = date_input_br("Prazo da Ação Educativa", value=parse_date(ae["prazo_acao_educativa"]) or date.today(), key="edit_ae_prazo")
            _, doc_cron = campo_documento_sei("Documento SEI Cronograma", "edit_ae_doc_cron", ae["documento_sei_cronograma"] or "", "Somente números")
            status = st.selectbox("Status Automático", STATUS_AE, index=STATUS_AE.index(ae["status_automatico"]) if ae["status_automatico"] in STATUS_AE else 0, key="edit_ae_status")
            finalizada = st.selectbox("Finalizada?", SIM_NAO, index=SIM_NAO.index(ae["finalizada"]) if ae["finalizada"] in SIM_NAO else 0, key="edit_ae_finalizada")
            obs_edit = st.text_area("Observação", value=ae["observacao"], key="edit_ae_obs")
            salvar = st.form_submit_button("Salvar alterações")
        if salvar:
            erros = []
            if not validar_processo_sei(proc_digit): erros.append("Processo SEI deve conter 17 números.")
            if not validar_documento_sei(documento): erros.append("Documento SEI Instaurador deve conter somente números.")
            if doc_cron and not validar_documento_sei(doc_cron): erros.append("Documento SEI Cronograma deve conter somente números.")
            if prazo_ae < data_emissao: erros.append("Prazo da AE não pode ser anterior à data de emissão.")
            if not escopo.strip(): erros.append("Escopo é obrigatório.")
            if erros:
                for e in erros: st.error(e)
                return
            novo_id = gerar_id_ae(documento)
            try:
                execute("""
                    UPDATE ae SET id_ae=?, processo_sei=?, documento_sei_instaurador=?, concessionaria=?, ano_concessao=?, ano_concessao_limpo=?, disciplina=?, data_emissao=?, data_alinhamento_escopo=?, escopo=?, prazo_envio_cronograma=?, prazo_acao_educativa=?, documento_sei_cronograma=?, status_automatico=?, finalizada=?, observacao=? WHERE id_ae=?
                """, (novo_id, processo, documento, concessionaria, formatar_ano_concessao(ano_edit), int(ano_edit), disciplina, data_emissao.isoformat(), data_alinhamento.isoformat(), escopo, prazo_cron.isoformat(), prazo_ae.isoformat(), doc_cron, status, finalizada, obs_edit, id_ae))
                execute("UPDATE ae_acompanhamentos SET id_ae=? WHERE id_ae=?", (novo_id, id_ae))
                audit(st.session_state["email"], perfil, "Editou AE", "ae", id_ae, str(ae.to_dict()), novo_id)
                limpar_selecao()
                st.success("Ação Educativa atualizada com sucesso.")
                st.rerun()
            except sqlite3.IntegrityError:
                st.error("Já existe outra AE com esse Documento SEI Instaurador.")
            except Exception as e:
                st.error(f"Erro ao atualizar AE: {e}")
    with st.expander("Excluir Ação Educativa", expanded=False):
        st.warning("Esta ação apagará também acompanhamentos vinculados à AE.")
        confirmar = st.checkbox(f"Confirmo que desejo excluir a AE {id_ae}", key="conf_del_ae")
        texto = st.text_input("Digite EXCLUIR para confirmar", key="txt_del_ae")
        if st.button("Excluir Ação Educativa"):
            if confirmar and texto == "EXCLUIR":
                execute("DELETE FROM ae_acompanhamentos WHERE id_ae=?", (id_ae,))
                execute("DELETE FROM ae WHERE id_ae=?", (id_ae,))
                audit(st.session_state["email"], perfil, "Excluiu AE", "ae", id_ae, str(ae.to_dict()), "")
                limpar_selecao()
                st.success("Ação Educativa excluída com sucesso.")
                st.rerun()
            else:
                st.error("Confirmação inválida. Marque a caixa e digite EXCLUIR.")

# ============================================================
# Pendências ANTT
# ============================================================

def tela_pendencias_antt():
    page_header("Pendências ANTT", "Fila de validações aguardando análise e decisão da ANTT.", ["Validação", "Prioridade"])
    if st.session_state["perfil"] != "ANTT":
        st.warning("Acesso restrito aos usuários da ANTT.")
        return

    df_anc = query_df("""
        SELECT r.*, a.processo_sei, a.documento_sei, a.disciplina, a.ano_concessao, a.prazo, a.status
        FROM anc_respostas r
        LEFT JOIN anc a ON a.id_anc = r.id_anc
        WHERE r.status_resposta IN ('Aguardando validação ANTT','Complementação solicitada')
        ORDER BY r.data_resposta ASC
    """)
    if not df_anc.empty:
        df_anc["situacao_prazo"] = df_anc.apply(lambda r: situacao_prazo_anc(r["status"], r["prazo"]), axis=1)

    df_ae = query_df("""
        SELECT r.*, a.processo_sei, a.documento_sei_instaurador, a.disciplina, a.ano_concessao,
               a.prazo_acao_educativa, a.finalizada
        FROM ae_acompanhamentos r
        LEFT JOIN ae a ON a.id_ae = r.id_ae
        WHERE r.status_acompanhamento IN ('Aguardando validação ANTT','Complementação solicitada')
        ORDER BY r.data_envio ASC
    """)
    if not df_ae.empty:
        df_ae["situacao_prazo"] = df_ae.apply(lambda r: situacao_prazo_ae(r["finalizada"], r["prazo_acao_educativa"]), axis=1)

    c1, c2, c3 = st.columns(3)
    with c1:
        meta_card("Pendências ANC", len(df_anc), "respostas aguardando decisão")
    with c2:
        meta_card("Pendências AE", len(df_ae), "acompanhamentos aguardando decisão")
    with c3:
        meta_card("Total", len(df_anc) + len(df_ae), "itens na fila ANTT")

    st.markdown('<div class="antt-section-title">Respostas de ANC aguardando validação</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="antt-muted-note">{len(df_anc)} registros exibidos.</div>', unsafe_allow_html=True)
    sel_anc = tabela_interativa(preparar_tabela_pendencia_anc(df_anc), key_grid("grid_pend_anc"), height=300) if not df_anc.empty else None
    if sel_anc:
        pend_id = id_selecionado(sel_anc, "ID pendência", "id")
        sel_anc_original = df_anc[df_anc["id"] == int(pend_id)].iloc[0].to_dict()
        destaque_sei("Documento SEI da comprovação", sel_anc_original.get("documento_sei_comprovacao", ""))
        mostrar_arquivo_upload(sel_anc_original.get("arquivo_link", ""), f"Resposta ANC {sel_anc_original.get('id')}")
        with st.form("validar_pendencia_anc"):
            decisao = st.selectbox("Decisão ANTT", STATUS_VALIDACAO_ANC, key="pend_anc_decisao")
            obs = st.text_area("Observação da validação ANTT", key="pend_anc_obs")
            salvar = st.form_submit_button("Registrar validação da ANC")
        if salvar:
            execute("""
                UPDATE anc_respostas SET status_resposta=?, data_validacao_antt=?, usuario_antt_validador=?, observacao_validacao_antt=? WHERE id=?
            """, (decisao, datetime.now().isoformat(timespec="seconds"), st.session_state["email"], obs, int(sel_anc_original["id"])))
            if decisao == "Aceita":
                execute("UPDATE anc SET status='Cumprido', documento_sei_comprovacao=? WHERE id_anc=?", (sel_anc_original["documento_sei_comprovacao"], sel_anc_original["id_anc"]))
            audit(st.session_state["email"], "ANTT", f"Validou resposta ANC como {decisao}", "anc_respostas", sel_anc_original["id_anc"], "", obs)
            limpar_selecao()
            st.success("Validação registrada.")
            st.rerun()
    elif df_anc.empty:
        st.info("Nenhuma resposta de ANC aguardando validação.")

    st.markdown('<div class="antt-section-title">Acompanhamentos de Ação Educativa aguardando validação</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="antt-muted-note">{len(df_ae)} registros exibidos.</div>', unsafe_allow_html=True)
    sel_ae = tabela_interativa(preparar_tabela_pendencia_ae(df_ae), key_grid("grid_pend_ae"), height=300) if not df_ae.empty else None
    if sel_ae:
        pend_id = id_selecionado(sel_ae, "ID pendência", "id")
        sel_ae_original = df_ae[df_ae["id"] == int(pend_id)].iloc[0].to_dict()
        destaque_sei("Documento SEI do acompanhamento", sel_ae_original.get("documento_sei_acompanhamento", ""))
        mostrar_arquivo_upload(sel_ae_original.get("arquivo_link", ""), f"Acompanhamento AE {sel_ae_original.get('id')}")
        with st.form("validar_pendencia_ae"):
            decisao = st.selectbox("Decisão ANTT", STATUS_VALIDACAO_AE, key="pend_ae_decisao")
            obs = st.text_area("Observação da validação ANTT", key="pend_ae_obs")
            salvar = st.form_submit_button("Registrar validação da AE")
        if salvar:
            execute("""
                UPDATE ae_acompanhamentos SET status_acompanhamento=?, data_validacao_antt=?, usuario_antt_validador=?, observacao_validacao_antt=? WHERE id=?
            """, (decisao, datetime.now().isoformat(timespec="seconds"), st.session_state["email"], obs, int(sel_ae_original["id"])))
            if decisao == "Aceito":
                mes = int(sel_ae_original["mes_acompanhamento"])
                if mes in [1, 2, 3, 4, 5]:
                    execute(f"UPDATE ae SET documento_sei_acomp_{mes}=? WHERE id_ae=?", (sel_ae_original["documento_sei_acompanhamento"], sel_ae_original["id_ae"]))
            audit(st.session_state["email"], "ANTT", f"Validou acompanhamento AE como {decisao}", "ae_acompanhamentos", sel_ae_original["id_ae"], "", obs)
            limpar_selecao()
            st.success("Validação registrada.")
            st.rerun()
    elif df_ae.empty:
        st.info("Nenhum acompanhamento de AE aguardando validação.")

# ============================================================
# Outras telas
# ============================================================

def tela_importar_excel():
    page_header("Importar Dados", "Atualize a base por planilha Excel mantendo auditoria das cargas.", ["Excel", "Base oficial"])
    if st.session_state["perfil"] != "ANTT":
        st.warning("Acesso restrito aos usuários da ANTT.")
        return
    st.info("Envie uma planilha Excel com abas Fato_API_BI, Fato_ANC_BI e Fato_AE_BI. Também são aceitas API, ANC e AcaoEducativa.")
    modo = st.radio("Modo de importação", ["Acrescentar ou atualizar registros", "Substituir medidas existentes"])
    uploaded_file = st.file_uploader("Selecione o arquivo Excel", type=["xlsx"])
    confirmar = st.checkbox("Confirmo que revisei o arquivo e desejo importar as informações.")
    if st.button("Importar Excel"):
        if uploaded_file is None:
            st.error("Selecione um arquivo Excel.")
            return
        if not confirmar:
            st.error("Marque a confirmação antes de importar.")
            return
        try:
            xls = pd.ExcelFile(uploaded_file, engine="openpyxl")
            total_api, total_anc, total_ae = importar_planilha_xls_para_banco(xls, modo, True, uploaded_file.name)
            limpar_selecao()
            st.success(f"Importação concluída: {total_api} APIs, {total_anc} ANCs e {total_ae} Ações Educativas.")
        except Exception as e:
            st.error(f"Erro ao importar Excel: {e}")


def area_concessionaria():
    page_header("Área da Concessionária", "Painel operacional para respostas de ANC e acompanhamentos de AE.", ["Concessionária", "Operação"])
    concessionaria = st.session_state["concessionaria"]
    df_anc = query_df("SELECT * FROM anc WHERE concessionaria=?", (concessionaria,))
    df_ae = query_df("SELECT * FROM ae WHERE concessionaria=?", (concessionaria,))
    if not df_anc.empty:
        df_anc["situacao_prazo"] = df_anc.apply(lambda r: situacao_prazo_anc(r["status"], r["prazo"]), axis=1)
    if not df_ae.empty:
        df_ae["situacao_prazo"] = df_ae.apply(lambda r: situacao_prazo_ae(r["finalizada"], r["prazo_acao_educativa"]), axis=1)
    base = pd.concat([
        df_anc[["concessionaria", "disciplina", "ano_concessao_limpo", "data_emissao"]] if not df_anc.empty else pd.DataFrame(columns=["concessionaria", "disciplina", "ano_concessao_limpo", "data_emissao"]),
        df_ae[["concessionaria", "disciplina", "ano_concessao_limpo", "data_emissao"]] if not df_ae.empty else pd.DataFrame(columns=["concessionaria", "disciplina", "ano_concessao_limpo", "data_emissao"]),
    ], ignore_index=True)
    conc, disc, ano_conc, ano_emissao, mes_emissao = filtros_medidas(base, "area_concessionaria")
    df_anc = aplicar_filtros(df_anc, conc, disc, ano_conc, ano_emissao, mes_emissao)
    df_ae = aplicar_filtros(df_ae, conc, disc, ano_conc, ano_emissao, mes_emissao)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Minhas ANCs", len(df_anc))
    c2.metric("ANCs vencidas", len(df_anc[df_anc["situacao_prazo"] == "Vencida"]) if not df_anc.empty else 0)
    c3.metric("ANCs a vencer", len(df_anc[df_anc["situacao_prazo"] == "A vencer em 15 dias"]) if not df_anc.empty else 0)
    c4.metric("Ações Educativas ativas", len(df_ae[df_ae["finalizada"].str.lower() != "sim"]) if not df_ae.empty else 0)
    st.subheader("Selecione uma ANC para informar cumprimento")
    selected_anc = tabela_interativa(preparar_tabela_anc(df_anc), key_grid("grid_area_conc_anc"), 320)
    if selected_anc:
        st.info("Para responder esta ANC, abra a tela 'Aviso de Não Conformidade (ANC)'.")
        if st.button("Ir para a ANC selecionada"):
            ir_para(NOME_ANC)
    st.subheader("Selecione uma Ação Educativa para enviar acompanhamento")
    selected_ae = tabela_interativa(preparar_tabela_ae(df_ae), key_grid("grid_area_conc_ae"), 320)
    if selected_ae:
        st.info("Para enviar acompanhamento desta Ação Educativa, abra a tela 'Ação Educativa (AE)'.")
        if st.button("Ir para a Ação Educativa selecionada"):
            ir_para(NOME_AE)



def tela_tabela_geral():
    page_header("Tabela Geral", "Visão consolidada de API, ANC e AE com os mesmos filtros do dashboard.", ["Consolidado", "Exportação"])
    where, params = where_concessionaria()
    df_api = query_df(f"SELECT * FROM api {where}", params)
    df_anc = query_df(f"SELECT * FROM anc {where}", params)
    df_ae = query_df(f"SELECT * FROM ae {where}", params)

    if not df_api.empty:
        df_api = df_api.copy()
        df_api["tipo"] = "API"
        df_api["id_medida"] = df_api["id_api"]
        df_api["situacao_prazo"] = "-"
        df_api["prazo_referencia"] = ""
    if not df_anc.empty:
        df_anc = df_anc.copy()
        df_anc["tipo"] = "ANC"
        df_anc["id_medida"] = df_anc["id_anc"]
        df_anc["situacao_prazo"] = df_anc.apply(lambda r: situacao_prazo_anc(r["status"], r["prazo"]), axis=1)
        df_anc["prazo_referencia"] = df_anc["prazo"]
    if not df_ae.empty:
        df_ae = df_ae.copy()
        df_ae["tipo"] = "AE"
        df_ae["id_medida"] = df_ae["id_ae"]
        df_ae["situacao_prazo"] = df_ae.apply(lambda r: situacao_prazo_ae(r["finalizada"], r["prazo_acao_educativa"]), axis=1)
        df_ae["prazo_referencia"] = df_ae["prazo_acao_educativa"]

    frames = [df for df in [df_api, df_anc, df_ae] if not df.empty]
    if not frames:
        st.info("Nenhuma medida encontrada.")
        return

    base = pd.concat(frames, ignore_index=True, sort=False)
    with st.expander("Filtros da tabela", expanded=True):
        conc, disc, ano_conc, ano_emissao, mes_emissao = filtros_medidas(base, "geral")
        base = aplicar_filtros(base, conc, disc, ano_conc, ano_emissao, mes_emissao)
        base = filtro_situacao_prazo(base, "geral")

    resumo_lista(base, "medidas")
    contexto_tabela(base, "Medidas consolidadas")
    tabela_geral = preparar_tabela_geral(base)
    tabela_interativa(tabela_geral, key_grid("grid_tabela_geral"), height=520)
    st.download_button("Exportar tabela geral em CSV", tabela_geral.to_csv(index=False).encode("utf-8-sig"), "tabela_geral_medidas.csv", "text/csv")

def tela_auditoria():
    page_header("Histórico de Auditoria", "Rastreie alterações, validações e importações realizadas no sistema.", ["Controle", "Governança"])
    if st.session_state["perfil"] != "ANTT":
        st.warning("Acesso restrito à ANTT.")
        return
    df = query_df("SELECT * FROM auditoria ORDER BY data_hora DESC")
    if df.empty:
        st.info("Nenhum registro de auditoria.")
    else:
        tabela_interativa(df, key_grid("grid_auditoria"), 500)
        st.download_button("Exportar auditoria em CSV", df.to_csv(index=False).encode("utf-8-sig"), "auditoria.csv", "text/csv")


def tela_admin():
    page_header("Administração", "Gerencie usuários e permissões de acesso ao aplicativo.", ["Usuários", "Acesso"])
    if st.session_state["perfil"] != "ANTT":
        st.warning("Acesso restrito à ANTT.")
        return
    st.subheader("Usuários cadastrados")
    df = query_df("SELECT id, email, perfil, concessionaria, ativo FROM usuarios ORDER BY perfil, email")
    tabela_interativa(df, key_grid("grid_usuarios"), 320)
    st.subheader("Criar novo usuário")
    with st.form("novo_usuario"):
        email = st.text_input("E-mail")
        senha = st.text_input("Senha inicial", type="password")
        perfil = st.selectbox("Perfil", ["ANTT", "Concessionária", "Visualizador"])
        concessionaria = st.selectbox("Concessionária", [""] + CONCESSIONARIAS)
        criar = st.form_submit_button("Criar usuário")
    if criar:
        if not email or not senha:
            st.error("Informe e-mail e senha.")
            return
        if perfil == "Concessionária" and not concessionaria:
            st.error("Usuário concessionária deve estar vinculado a uma concessionária.")
            return
        try:
            execute("INSERT INTO usuarios (email, senha_hash, perfil, concessionaria, ativo) VALUES (?, ?, ?, ?, 1)", (email, hash_password(senha), perfil, concessionaria))
            st.success("Usuário criado com sucesso.")
        except sqlite3.IntegrityError:
            st.error("Já existe usuário com esse e-mail.")
        except Exception as e:
            st.error(f"Erro ao criar usuário: {e}")



MENU_LABELS = {
    "Dashboard Geral": "Visão geral",
    "Tabela Geral": "Planilha geral",
    "Pendências ANTT": "Pendências SEI",
    "Cadastrar Medida": "Cadastrar medida",
    "Importar Excel": "Importar planilha",
    NOME_API: "APIs",
    NOME_ANC: "ANCs",
    NOME_AE: "Ações educativas",
    "Auditoria": "Auditoria",
    "Administração": "Usuários e acesso",
    "Área da Concessionária": "Minha área",
}


def rotulo_menu(pagina):
    return MENU_LABELS.get(pagina, pagina)

# ============================================================
# Main
# ============================================================

def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide", initial_sidebar_state="expanded")
    aplicar_estilo_powerbi()
    if "grid_nonce" not in st.session_state:
        st.session_state["grid_nonce"] = 0
    init_db()
    importar_excel_se_necessario()
    if "logged" not in st.session_state:
        login_screen()
        return
    sidebar_brand()
    sidebar_profile()
    logout_button()
    perfil = st.session_state["perfil"]
    if perfil == "ANTT":
        opcoes = ["Dashboard Geral", "Tabela Geral", "Pendências ANTT", "Cadastrar Medida", "Importar Excel", NOME_API, NOME_ANC, NOME_AE, "Auditoria", "Administração"]
    elif perfil == "Concessionária":
        opcoes = ["Área da Concessionária", "Dashboard Geral", "Tabela Geral", NOME_API, NOME_ANC, NOME_AE]
    else:
        opcoes = ["Dashboard Geral", "Tabela Geral", NOME_API, NOME_ANC, NOME_AE]
    if "pagina_atual" not in st.session_state or st.session_state["pagina_atual"] not in opcoes:
        st.session_state["pagina_atual"] = opcoes[0]
    destino = st.session_state.pop("proxima_pagina", None)
    if destino in opcoes:
        st.session_state["pagina_atual"] = destino

    pagina = st.sidebar.radio("Menu", opcoes, key="pagina_atual", format_func=rotulo_menu)
    if pagina == "Dashboard Geral":
        dashboard()
    elif pagina == "Tabela Geral":
        tela_tabela_geral()
    elif pagina == "Pendências ANTT":
        tela_pendencias_antt()
    elif pagina == "Cadastrar Medida":
        tela_cadastrar_medida()
    elif pagina == "Importar Excel":
        tela_importar_excel()
    elif pagina == "Área da Concessionária":
        area_concessionaria()
    elif pagina == NOME_API:
        tela_api()
    elif pagina == NOME_ANC:
        tela_anc()
    elif pagina == NOME_AE:
        tela_ae()
    elif pagina == "Auditoria":
        tela_auditoria()
    elif pagina == "Administração":
        tela_admin()

if __name__ == "__main__":
    main()
