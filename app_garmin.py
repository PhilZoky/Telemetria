import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import simplekml
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
import re
import os
import io
import json
import requests
from datetime import datetime
import streamlit.components.v1 as components

try:
    import google.generativeai as genai
except ImportError:
    genai = None

# ==========================================
# CONFIGURAÇÃO DA PÁGINA E PALETA AERONÁUTICA
# ==========================================
st.set_page_config(
    page_title="AeroTelemetry | Flight Data Monitoring",
    page_icon="✈️",
    layout="wide",
    initial_sidebar_state="expanded"
)

if "tema_ui" not in st.session_state:
    st.session_state.tema_ui = "Escuro (Cockpit)"

def obter_paleta_cores():
    modo_claro = "Claro" in st.session_state.tema_ui
    if modo_claro:
        return {
            "modo": "light",
            "bg_solid": "#e6f2ff",
            "bg_app": "linear-gradient(180deg, #dbeafe 0%, #f0f9ff 45%, #e0f2fe 100%)",
            "bg_sidebar": "#d9eaff",
            "bg_card": "#ffffff",
            "bg_card_grad": "linear-gradient(180deg, #ffffff 0%, #f8fbff 100%)",
            "bg_header": "linear-gradient(90deg, #172554 0%, #1e40af 50%, #1d4ed8 100%)",
            "txt_header": "#ffffff",
            "sub_header": "#bae6fd",
            "bg_badge_header": "rgba(255, 255, 255, 0.16)",
            "bg_input": "#ffffff",
            "txt_main": "#0f172a",
            "txt_sec": "#1e3a8a",
            "txt_muted": "#475569",
            "border": "#93c5fd",
            "accent": "#1d4ed8",
            "accent_light": "#0284c7",
            "plotly_template": "plotly_white",
            "plotly_bg": "#ffffff",
            "map_style": "carto-positron",
            "tab_bg": "#cfe3ff",
            "tab_txt": "#000000"
        }
    else:
        return {
            "modo": "dark",
            "bg_solid": "#0b192c",
            "bg_app": "linear-gradient(180deg, #091526 0%, #0d213a 50%, #0b192c 100%)",
            "bg_sidebar": "#081220",
            "bg_card": "#102542",
            "bg_card_grad": "linear-gradient(160deg, #132c4e 0%, #0d1f38 100%)",
            "bg_header": "linear-gradient(90deg, #0f294a 0%, #1d4ed8 50%, #0f294a 100%)",
            "txt_header": "#f8fafc",
            "sub_header": "#93c5fd",
            "bg_badge_header": "rgba(11, 25, 44, 0.65)",
            "bg_input": "#0f243e",
            "txt_main": "#f1f5f9",
            "txt_sec": "#cbd5e1",
            "txt_muted": "#94a3b8",
            "border": "#1e3a5f",
            "accent": "#38bdf8",
            "accent_light": "#60a5fa",
            "plotly_template": "plotly_dark",
            "plotly_bg": "#0f243e",
            "map_style": "carto-darkmatter",
            "tab_bg": "#0f243e",
            "tab_txt": "#ffffff"
        }

T = obter_paleta_cores()

def injetar_estilo_profissional(t):
    esquema_cor = "light" if t['modo'] == 'light' else "dark"
    bg_barra_abas = t['bg_header'] if t['modo'] == 'light' else t['tab_bg']
    cor_aba_txt = "#000000" if t['modo'] == 'light' else "#ffffff"
    css = f"""<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@500;600;700&display=swap');
:root, html, body, [data-testid="stApp"] {{
    color-scheme: {esquema_cor} !important;
    --primary-color: {t['accent']} !important;
    --background-color: {t['bg_solid']} !important;
    --secondary-background-color: {t['bg_sidebar']} !important;
    --text-color: {t['txt_main']} !important;
}}
html, body, [data-testid="stApp"], .stApp, [data-testid="stAppViewContainer"], [data-testid="stMain"] {{
    font-family: 'Inter', 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
    background-color: {t['bg_solid']} !important;
    background: {t['bg_app']} !important;
    color: {t['txt_main']} !important;
}}
#MainMenu, [data-testid="stToolbar"], [data-testid="stDecoration"], footer {{
    display: none !important;
    visibility: hidden !important;
}}
[data-testid="stHeader"] {{
    background-color: transparent !important;
}}
.block-container {{
    padding-top: 2.2rem !important;
    padding-bottom: 3rem !important;
    max-width: 96% !important;
}}
html body section[data-testid="stSidebar"],
html body section[data-testid="stSidebar"] > div,
html body [data-testid="stSidebarContent"],
html body [data-testid="stSidebarUserContent"] {{
    background-color: {t['bg_sidebar']} !important;
    background: {t['bg_sidebar']} !important;
    border-right: 1px solid {t['border']} !important;
}}
html body section[data-testid="stSidebar"] * {{
    color: {t['txt_main']} !important;
}}
label, .stMarkdown p, .stRadio label, [data-testid="stWidgetLabel"] p {{
    color: {t['txt_main']} !important;
    font-weight: 600 !important;
}}
/* Abas */
.stTabs [data-baseweb="tab-list"],
div[role="tablist"] {{
    gap: 8px !important;
    background: {bg_barra_abas} !important;
    padding: 8px 10px !important;
    border-radius: 10px !important;
    border: 1px solid {t['border']} !important;
}}
.stTabs button[data-baseweb="tab"],
div[role="tablist"] button[role="tab"],
button[data-baseweb="tab"],
button[role="tab"] {{
    height: 42px !important;
    border-radius: 7px !important;
    padding: 0 18px !important;
    background-color: {t['bg_card']} !important;
    border: 1px solid {t['border']} !important;
    color: {cor_aba_txt} !important;
    -webkit-text-fill-color: {cor_aba_txt} !important;
    font-weight: 700 !important;
    font-size: 12.5px !important;
}}
.stTabs button[data-baseweb="tab"] *,
div[role="tablist"] button[role="tab"] *,
button[data-baseweb="tab"] p,
button[data-baseweb="tab"] span,
button[data-baseweb="tab"] div,
button[role="tab"] p,
button[role="tab"] span,
button[role="tab"] div {{
    color: {cor_aba_txt} !important;
    -webkit-text-fill-color: {cor_aba_txt} !important;
    font-weight: 700 !important;
}}
.stTabs button[data-baseweb="tab"][aria-selected="true"],
div[role="tablist"] button[role="tab"][aria-selected="true"],
button[data-baseweb="tab"][aria-selected="true"],
button[role="tab"][aria-selected="true"] {{
    background: linear-gradient(135deg, #1d4ed8 0%, #1e40af 100%) !important;
    border: 1px solid #1e3a8a !important;
    color: #ffffff !important;
    -webkit-text-fill-color: #ffffff !important;
}}
.stTabs button[data-baseweb="tab"][aria-selected="true"] *,
div[role="tablist"] button[role="tab"][aria-selected="true"] *,
button[data-baseweb="tab"][aria-selected="true"] p,
button[role="tab"][aria-selected="true"] p {{
    color: #ffffff !important;
    -webkit-text-fill-color: #ffffff !important;
}}
.stTabs [data-baseweb="tab-highlight"],
.stTabs [data-baseweb="tab-border"],
div[data-baseweb="tab-highlight"],
div[data-baseweb="tab-border"] {{
    display: none !important;
}}
/* Botões Principais */
.stButton > button, .stDownloadButton > button {{
    background: linear-gradient(135deg, #1d4ed8 0%, #1e40af 100%) !important;
    color: #ffffff !important;
    -webkit-text-fill-color: #ffffff !important;
    border: 1px solid #1e3a8a !important;
    border-radius: 8px !important;
    padding: 0.55rem 1.2rem !important;
    font-weight: 600 !important;
    font-size: 13px !important;
}}
.stButton > button *, .stDownloadButton > button * {{
    color: #ffffff !important;
    -webkit-text-fill-color: #ffffff !important;
}}
/* File Uploader */
[data-testid="stFileUploader"] {{
    background-color: {t['bg_card']} !important;
    border: 1.5px dashed {t['accent']} !important;
    border-radius: 10px;
    padding: 10px;
}}
[data-testid="stFileUploader"] section,
[data-testid="stFileUploader"] ul,
[data-testid="stFileUploader"] li,
[data-testid="stFileUploader"] [role="listitem"],
[data-testid="stFileUploaderFile"] {{
    background-color: {t['bg_input']} !important;
    color: {t['txt_main']} !important;
    border-radius: 8px !important;
}}
[data-testid="stFileUploader"] li *,
[data-testid="stFileUploader"] section * {{
    color: {t['txt_main']} !important;
    -webkit-text-fill-color: {t['txt_main']} !important;
}}
[data-testid="stFileUploader"] button {{
    background-color: {t['accent']} !important;
    color: #ffffff !important;
    border: none !important;
}}
[data-testid="stFileUploader"] button * {{
    color: #ffffff !important;
    -webkit-text-fill-color: #ffffff !important;
}}
/* Perfil de Envelope na barra lateral: fundo azul-marinho e texto BRANCO */
section[data-testid="stSidebar"] [data-baseweb="select"] > div {{
    background-color: #1e3a8a !important;
    border: 1px solid #1d4ed8 !important;
    border-radius: 8px !important;
}}
section[data-testid="stSidebar"] [data-baseweb="select"] *,
section[data-testid="stSidebar"] [data-baseweb="select"] div,
section[data-testid="stSidebar"] [data-baseweb="select"] span,
section[data-testid="stSidebar"] [data-baseweb="select"] svg {{
    color: #ffffff !important;
    -webkit-text-fill-color: #ffffff !important;
    fill: #ffffff !important;
}}
/* Inputs, Senha, Autofill e TextAreas */
.stTextInput > div > div, .stTextInput > div > div > input,
.stNumberInput > div > div > input,
.stTextArea > div > div, .stTextArea > div > div > textarea {{
    background-color: {t['bg_input']} !important;
    border-color: {t['border']} !important;
    color: {t['txt_main']} !important;
    -webkit-text-fill-color: {t['txt_main']} !important;
    border-radius: 8px !important;
    font-weight: 500 !important;
}}
input:-webkit-autofill, input:-webkit-autofill:hover, input:-webkit-autofill:focus {{
    -webkit-box-shadow: 0 0 0 1000px {t['bg_input']} inset !important;
    -webkit-text-fill-color: {t['txt_main']} !important;
}}
[data-testid="stExpander"], .streamlit-expanderHeader {{
    background-color: {t['bg_card']} !important;
    border: 1px solid {t['border']} !important;
    color: {t['txt_main']} !important;
    border-radius: 8px !important;
}}
h1, h2, h3, h4 {{
    color: {t['txt_main']} !important;
    font-weight: 700 !important;
}}
hr {{
    border-color: {t['border']} !important;
    margin: 1.4rem 0 !important;
}}
</style>"""
    st.markdown(css, unsafe_allow_html=True)

    st.markdown("""<style>
/* Força todas as abas (Login e Painel) para BRANCO PURO (#ffffff) no Claro e no Escuro */
button[data-baseweb="tab"],
button[data-baseweb="tab"] [data-testid="stMarkdownContainer"],
button[data-baseweb="tab"] [data-testid="stMarkdownContainer"] p,
button[role="tab"],
button[role="tab"] p,
button[role="tab"] span {
    color: #ffffff !important;
    -webkit-text-fill-color: #ffffff !important;
    opacity: 1 !important;
    font-weight: 600 !important;
}
button[data-baseweb="tab"][aria-selected="true"] [data-testid="stMarkdownContainer"] p,
button[role="tab"][aria-selected="true"] p {
    color: #ffffff !important;
    -webkit-text-fill-color: #ffffff !important;
    font-weight: 700 !important;
}
</style>""", unsafe_allow_html=True)

    if t['modo'] == 'light':
        st.markdown("""<style>
[data-testid="stFileUploaderFile"],
[data-testid="stFileUploaderFile"] div,
[data-testid="stFileUploaderFileData"],
[data-testid="stFileUploader"] ul li,
[data-testid="stFileUploader"] ul li div {
    background-color: #ffffff !important;
    background: #ffffff !important;
    color: #0f172a !important;
    -webkit-text-fill-color: #0f172a !important;
}
[data-testid="stFileUploaderFile"] {
    border: 1px solid #93c5fd !important;
    border-radius: 8px !important;
    padding: 4px 8px !important;
}
section[data-testid="stSidebar"] [data-testid="stSelectbox"] [data-baseweb="select"],
section[data-testid="stSidebar"] [data-testid="stSelectbox"] [data-baseweb="select"] > div,
section[data-testid="stSidebar"] [data-testid="stSelectbox"] [data-baseweb="select"] > div > div {
    background-color: #ffffff !important;
    background: #ffffff !important;
    border-color: #93c5fd !important;
    color: #0f172a !important;
    -webkit-text-fill-color: #0f172a !important;
}
section[data-testid="stSidebar"] [data-testid="stSelectbox"] [data-baseweb="select"] svg {
    fill: #0f172a !important;
    color: #0f172a !important;
}
[data-testid="stTextArea"] [data-baseweb="textarea"],
[data-testid="stTextArea"] [data-baseweb="base-input"],
[data-testid="stTextArea"] textarea,
textarea {
    background-color: #ffffff !important;
    background: #ffffff !important;
    color: #0f172a !important;
    -webkit-text-fill-color: #0f172a !important;
    caret-color: #0f172a !important;
    border-color: #93c5fd !important;
    border-radius: 8px !important;
}
[data-testid="InputInstructions"] {
    color: #475569 !important;
    -webkit-text-fill-color: #475569 !important;
}
</style>""", unsafe_allow_html=True)

injetar_estilo_profissional(T)
st.markdown("""<style>
[data-baseweb="tab-list"] * {
    color: #ffffff !important;
    -webkit-text-fill-color: #ffffff !important;
    opacity: 1 !important;
}
[data-baseweb="tab-highlight"] {
    background-color: #ffffff !important;
}
</style>""", unsafe_allow_html=True)


# ==========================================
# PERSISTÊNCIA EM NUVEM (GITHUB GIST SECRETO)
# ==========================================
def _obter_config_gist():
    try:
        if "GITHUB_TOKEN" in st.secrets and "GIST_ID" in st.secrets:
            token = str(st.secrets["GITHUB_TOKEN"]).strip()
            gist_id = str(st.secrets["GIST_ID"]).strip()
            if token and gist_id:
                return token, gist_id
    except Exception:
        pass
    return None, None

def ler_json_nuvem(nome_arquivo, padrao_fn):
    if "_cloud_cache" not in st.session_state:
        st.session_state["_cloud_cache"] = {}
    if nome_arquivo in st.session_state["_cloud_cache"]:
        return json.loads(json.dumps(st.session_state["_cloud_cache"][nome_arquivo]))

    token, gist_id = _obter_config_gist()
    if token and gist_id:
        try:
            headers = {
                "Authorization": f"Bearer {token}",
                "Accept": "application/vnd.github+json"
            }
            resp = requests.get(f"https://api.github.com/gists/{gist_id}", headers=headers, timeout=10)
            if resp.status_code == 200:
                files = resp.json().get("files", {})
                if nome_arquivo in files and files[nome_arquivo].get("content"):
                    dados = json.loads(files[nome_arquivo]["content"])
                    if isinstance(dados, dict) and dados:
                        st.session_state["_cloud_cache"][nome_arquivo] = dados
                        return json.loads(json.dumps(dados))
        except Exception:
            pass

    if os.path.exists(nome_arquivo):
        try:
            with open(nome_arquivo, "r", encoding="utf-8") as f:
                dados_local = json.load(f)
            if isinstance(dados_local, dict) and dados_local:
                st.session_state["_cloud_cache"][nome_arquivo] = dados_local
                return json.loads(json.dumps(dados_local))
        except Exception:
            pass

    padrao = padrao_fn()
    st.session_state["_cloud_cache"][nome_arquivo] = json.loads(json.dumps(padrao))
    return padrao

def salvar_json_nuvem(nome_arquivo, dados):
    if "_cloud_cache" not in st.session_state:
        st.session_state["_cloud_cache"] = {}
    st.session_state["_cloud_cache"][nome_arquivo] = json.loads(json.dumps(dados))

    try:
        with open(nome_arquivo, "w", encoding="utf-8") as f:
            json.dump(dados, f, indent=4, ensure_ascii=False)
    except Exception:
        pass

    token, gist_id = _obter_config_gist()
    if token and gist_id:
        try:
            headers = {
                "Authorization": f"Bearer {token}",
                "Accept": "application/vnd.github+json"
            }
            payload = {
                "files": {
                    nome_arquivo: {
                        "content": json.dumps(dados, indent=4, ensure_ascii=False)
                    }
                }
            }
            resp = requests.patch(f"https://api.github.com/gists/{gist_id}", headers=headers, json=payload, timeout=15)
            return resp.status_code == 200
        except Exception:
            return False
    return False


# ==========================================
# CONFIGURAÇÃO DA INTELIGÊNCIA ARTIFICIAL
# ==========================================
PADRAO_API_KEY = ""

SYSTEM_PROMPT_FDM = """Você é um Engenheiro de Ensaios em Voo e Analista de Telemetria Aeronáutica (FDM).
REGRAS INVIOLÁVEIS:
1. USE APENAS OS DADOS FORNECIDOS NO PROMPT. Nunca invente horários, durações, clima, pressão atmosférica (hPa), peso/toneladas ou valores de sensores.
2. UNIDADES PERMITIDAS EXCLUSIVAMENTE: ft, kt, G, º (graus de Pitch/Roll), lb.ft (Torque), ºC (T5/ITT e OilT), % (Ng e Nh), PSI (OilP) e m (metros na formatura).
3. ESTRITAMENTE PROIBIDO mencionar ou calcular: kW, HP, potência elétrica, desempenho aeroelétrico, kg/s, toneladas, hPa, OAT/temperatura ambiente ou qualquer parâmetro que não esteja nos dados enviados.
4. NÃO sugira resolução de panes, troca de peças, boroscopia ou ações corretivas de manutenção. Descreva exclusivamente o voo, a aeronave e a performance entregue pelo motor.
5. Escreva em português técnico correto e fluente, sem repetir seções e sem usar emojis."""

def obter_credencial_ia():
    try:
        if "GEMINI_API_KEY" in st.secrets:
            return str(st.secrets["GEMINI_API_KEY"]).strip()
    except Exception:
        pass
    return PADRAO_API_KEY.strip()

def gerar_parecer_contingencia(max_g, min_g, pico_tq, pico_itt, pico_ng, pico_nh, alertas):
    """Gera análise descritiva estruturada automaticamente caso esteja offline ou sem chave ativa."""
    status_geral = "OPERAÇÃO INTEGRAL DENTRO DO ENVELOPE NOMINAL" if not alertas else "REGISTRO DE EXCEDÊNCIA OPERACIONAL NO PERFIL"
    bloco_alertas = "\n".join([f"* {a}" for a in alertas]) if alertas else "* **Sem Ocorrências:** Todos os sensores de célula e motor mantiveram-se dentro dos limites previstos para o perfil configurado."

    return f"""# ANÁLISE TÉCNICA DETALHADA DE VOO E PERFORMANCE (SOLO)
## 1. PERFIL OPERACIONAL E SÍNTESE DO VOO
* **Enquadramento Global:** **{status_geral}**
* **Envelope Estrutural (Força G):** A aeronave operou entre fator de carga mínimo de **{min_g:.2f}G** e máximo de **+{max_g:.2f}G** durante a missão.

## 2. PERFORMANCE DO GRUPO MOTOPROPULSOR
* **Torque e Propulsão:** Pico registrado de **{pico_tq:.0f} lb.ft** de Torque e rotação máxima de hélice (**Nh**) de **{pico_nh:.1f}%**.
* **Regime Termodinâmico:** Temperatura inter-turbinas (**T5/ITT**) atingiu pico de **{pico_itt:.0f} ºC**, acompanhada por rotação máxima do gerador de gás (**Ng**) de **{pico_ng:.1f}%**.

## 3. EVENTOS E EXCEDÊNCIAS REGISTRADAS
{bloco_alertas}"""

def gerar_texto_ia(prompt_texto):
    credencial = obter_credencial_ia()
    if not credencial:
        raise RuntimeError("Chave de API não configurada em st.secrets['GEMINI_API_KEY'].")

    if credencial.startswith("gsk_"):
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {credencial}"
        }
        modelos_preferidos = [
            "llama-3.3-70b-versatile",
            "meta-llama/llama-4-maverick-17b-128e-instruct",
            "meta-llama/llama-4-scout-17b-16e-instruct",
            "qwen/qwen3-32b",
            "openai/gpt-oss-120b",
            "llama3-70b-8192",
            "llama-3.1-8b-instant"
        ]
        try:
            r_mod = requests.get("https://api.groq.com/openai/v1/models", headers=headers, timeout=10)
            if r_mod.status_code == 200:
                bloqueados = ["whisper", "guard", "tts", "vision", "embed", "playai", "allam", "compound"]
                ativos = [
                    m["id"] for m in r_mod.json().get("data", [])
                    if not any(x in m["id"].lower() for x in bloqueados)
                ]
                ordenados = [m for m in modelos_preferidos if m in ativos]
                restantes_grandes = [
                    m for m in ativos
                    if m not in ordenados and not any(p in m.lower() for p in ["8b", "9b", "7b", "3b", "1b"])
                ]
                restantes_pequenos = [m for m in ativos if m not in ordenados and m not in restantes_grandes]
                if ordenados or restantes_grandes or restantes_pequenos:
                    modelos_preferidos = ordenados + restantes_grandes + restantes_pequenos
        except Exception:
            pass

        ultimo_erro_groq = ""
        url = "https://api.groq.com/openai/v1/chat/completions"

        for mod in modelos_preferidos:
            payload = {
                "model": mod,
                "messages": [
                    {"role": "system", "content": SYSTEM_PROMPT_FDM},
                    {"role": "user", "content": prompt_texto}
                ],
                "temperature": 0.1
            }
            resp = requests.post(url, headers=headers, json=payload, timeout=60)
            if resp.status_code == 200:
                st.session_state["ultimo_modelo_ia"] = mod
                return resp.json()["choices"][0]["message"]["content"]
            ultimo_erro_groq = f"Erro na Groq ({resp.status_code}) [{mod}]: {resp.text}"

        raise RuntimeError(ultimo_erro_groq)

    modelos_tentativa = ["gemini-2.5-flash", "gemini-2.5-pro", "gemini-2.0-flash"]
    payload = {
        "systemInstruction": {"parts": [{"text": SYSTEM_PROMPT_FDM}]},
        "contents": [{"parts": [{"text": prompt_texto}]}],
        "generationConfig": {"temperature": 0.1}
    }
    ultimo_erro = ""

    for mod in modelos_tentativa:
        url_key = f"https://generativelanguage.googleapis.com/v1beta/models/{mod}:generateContent?key={credencial}"
        resp = requests.post(url_key, headers={"Content-Type": "application/json"}, json=payload, timeout=60)
        if resp.status_code == 200:
            st.session_state["ultimo_modelo_ia"] = mod
            return resp.json()["candidates"][0]["content"]["parts"][0]["text"]
        if credencial.startswith("AQ."):
            url_bearer = f"https://generativelanguage.googleapis.com/v1beta/models/{mod}:generateContent"
            headers_b = {"Content-Type": "application/json", "Authorization": f"Bearer {credencial}"}
            resp_b = requests.post(url_bearer, headers=headers_b, json=payload, timeout=60)
            if resp_b.status_code == 200:
                st.session_state["ultimo_modelo_ia"] = mod
                return resp_b.json()["candidates"][0]["content"]["parts"][0]["text"]
        ultimo_erro = f"Status {resp.status_code} [{mod}]: {resp.text}"

    raise RuntimeError(f"Não foi possível autenticar na API. Detalhe: {ultimo_erro}")


# ==========================================
# 0. SISTEMA DE AUTENTICAÇÃO
# ==========================================
ARQUIVO_USUARIOS = "banco_usuarios.json"

def criar_admin_padrao():
    return {
        "Filipe_Iack": {
            "senha": "210772",
            "nome": "Filipe Iack",
            "email": "admin@telemetria.com",
            "organizacao": "AFA - Master"
        }
    }

def carregar_usuarios():
    dados = ler_json_nuvem(ARQUIVO_USUARIOS, criar_admin_padrao)
    if not isinstance(dados, dict) or not dados:
        dados = criar_admin_padrao()
    if "Filipe_Iack" not in dados:
        dados["Filipe_Iack"] = criar_admin_padrao()["Filipe_Iack"]
    return dados

def salvar_usuarios(usuarios):
    return salvar_json_nuvem(ARQUIVO_USUARIOS, usuarios)

usuarios_permitidos = carregar_usuarios()

if "usuario_logado" not in st.session_state:
    st.session_state.usuario_logado = None
if "nome_logado" not in st.session_state:
    st.session_state.nome_logado = None

if st.session_state.usuario_logado is None:
    st.markdown("<br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1.2, 1.8, 1.2])

    with col2:
        def _trocar_tema_login():
            st.session_state.tema_ui = st.session_state.radio_tema_login

        st.radio(
            "Tema Visual:",
            ["Claro (Diurno)", "Escuro (Cockpit)"],
            index=0 if "Claro" in st.session_state.tema_ui else 1,
            horizontal=True,
            label_visibility="collapsed",
            key="radio_tema_login",
            on_change=_trocar_tema_login
        )

        st.markdown(f"""
        <div style="background: {T['bg_card_grad']}; border: 1px solid {T['border']}; border-top: 4px solid {T['accent']}; border-radius: 14px; padding: 30px 32px; box-shadow: 0 12px 30px rgba(0,0,0,0.12); text-align: center; margin-bottom: 20px;">
            <div style="display: inline-block; background: rgba(29, 78, 216, 0.12); border: 1px solid rgba(29, 78, 216, 0.3); color: {T['accent']}; font-size: 11px; font-weight: 700; letter-spacing: 1.5px; padding: 4px 12px; border-radius: 20px; margin-bottom: 12px;">
                FLIGHT DATA MONITORING SYSTEM
            </div>
            <div style="margin: 0; color: {T['txt_main']}; font-size: 26px; font-weight: 700;">AeroTelemetry Tactical</div>
            <p style="margin: 6px 0 0 0; color: {T['txt_sec']}; font-size: 13.5px;">Plataforma Integrada de Análise de Voo, Motores e Inspetoria Técnica</p>
        </div>
        """, unsafe_allow_html=True)

        aba_login, aba_cadastro = st.tabs(["ACESSAR CREDENCIAIS", "CADASTRAR OPERADOR"])

        with aba_login:
            st.markdown("<div style='height: 8px'></div>", unsafe_allow_html=True)
            usuario_input = st.text_input("Identificação de Usuário (ID)", key="log_user", placeholder="Ex: Filipe_Iack")
            senha_input = st.text_input("Chave de Segurança (Senha)", type="password", key="log_pass", placeholder="••••••••")
            st.markdown("<div style='height: 6px'></div>", unsafe_allow_html=True)

            if st.button("AUTENTICAR NO SISTEMA", use_container_width=True):
                if usuario_input in usuarios_permitidos:
                    dados_usuario = usuarios_permitidos[usuario_input]
                    senha_correta = False
                    if isinstance(dados_usuario, str) and dados_usuario == senha_input:
                        senha_correta = True
                        st.session_state.nome_logado = usuario_input
                    elif isinstance(dados_usuario, dict) and dados_usuario.get("senha") == senha_input:
                        senha_correta = True
                        st.session_state.nome_logado = dados_usuario.get("nome", usuario_input)

                    if senha_correta:
                        st.session_state.usuario_logado = usuario_input
                        st.rerun()
                    else:
                        st.error("Credenciais inválidas. Verifique seu ID e Senha.")
                else:
                    st.error("Operador não localizado na base de dados.")

        with aba_cadastro:
            st.markdown("<div style='height: 8px'></div>", unsafe_allow_html=True)
            novo_nome = st.text_input("Nome e Posto / Graduação", key="cad_nome")
            novo_email = st.text_input("E-mail Institucional / Contato", key="cad_email")
            nova_org = st.text_input("Esquadrão / Unidade Aérea", key="cad_org")

            c_cad1, c_cad2 = st.columns(2)
            with c_cad1:
                novo_usuario = st.text_input("ID de Login", key="cad_user")
            with c_cad2:
                nova_senha = st.text_input("Definir Senha", type="password", key="cad_pass")

            if st.button("REGISTRAR OPERADOR", use_container_width=True):
                if novo_usuario in usuarios_permitidos:
                    st.warning("Este ID de usuário já está em uso.")
                elif not novo_usuario or not nova_senha or not novo_nome or not novo_email:
                    st.warning("Preencha todos os campos obrigatórios.")
                else:
                    usuarios_permitidos[novo_usuario] = {
                        "senha": nova_senha,
                        "nome": novo_nome,
                        "email": novo_email,
                        "organizacao": nova_org,
                        "data_cadastro": datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    }
                    salvou_nuvem = salvar_usuarios(usuarios_permitidos)
                    st.success(f"Credencial de {novo_nome} habilitada{' na nuvem' if salvou_nuvem else ''}. Acesse pela aba de Login.")

    st.stop()


# ==========================================
# 1. GERENCIADOR DE PERFIS (COM PERSISTÊNCIA EM NUVEM)
# ==========================================
_user_safe = re.sub(r'[^a-zA-Z0-9_-]', '_', str(st.session_state.usuario_logado))
ARQUIVO_PERFIS = f"perfis_{_user_safe}.json"

PERFIS_PADRAO = {
    "T-27M (Oficial)": {
        "tipo_aeronave": "Asa Fixa",
        "qtd_motores": 1,
        "g_max": 6.0, "g_min": -3.0,
        "tq_red": 1795.0, "tq_yel": 1795.0,
        "itt_red": 790.0, "itt_yel": 740.0,
        "ng_red": 101.5, "ng_yel": 100.0,
        "nh_red": 100.0, "nh_yel": 100.0,
        "oilt_red": 99.0, "oilt_min_green": 10.0,
        "oilp_red": 95.0, "oilp_min_green": 40.0
    },
    "Multi-Motor (Padrão)": {
        "tipo_aeronave": "Asa Fixa",
        "qtd_motores": 2,
        "g_max": 2.5, "g_min": 0.0,
        "tq_red": 1500.0, "tq_yel": 1400.0,
        "itt_red": 850.0, "itt_yel": 800.0,
        "ng_red": 100.0, "ng_yel": 98.0,
        "nh_red": 100.0, "nh_yel": 98.0,
        "oilt_red": 100.0, "oilt_min_green": 20.0,
        "oilp_red": 90.0, "oilp_min_green": 60.0
    }
}

def carregar_perfis():
    dados = ler_json_nuvem(ARQUIVO_PERFIS, lambda: json.loads(json.dumps(PERFIS_PADRAO)))
    if not isinstance(dados, dict) or not dados:
        dados = json.loads(json.dumps(PERFIS_PADRAO))
    if "T-27M (Oficial)" not in dados:
        dados["T-27M (Oficial)"] = PERFIS_PADRAO["T-27M (Oficial)"].copy()
    return dados

def salvar_perfis(perfis):
    return salvar_json_nuvem(ARQUIVO_PERFIS, perfis)

perfis_salvos = carregar_perfis()


# ==========================================
# 2. INICIALIZAÇÃO DE VARIÁVEIS DE SESSÃO E SINCRONIZAÇÃO DE IA
# ==========================================
obs_keys = [
    'obs_resumo',
    'obs_graficos',
    'obs_manobras',
    'obs_distancia',
    'obs_comparativo',
    'obs_3d'
]

for k in obs_keys:
    if k not in st.session_state:
        st.session_state[k] = ""

for sk in ["status_ia_solo", "status_ia_form", "status_ia_trend"]:
    if sk not in st.session_state:
        st.session_state[sk] = None

# Sincroniza as keys dos st.text_area ANTES de os widgets serem desenhados na tela
if st.session_state.pop("_sync_solo_ia", False):
    st.session_state["k_graf_rel"] = st.session_state.obs_graficos
    st.session_state["k_graf_manual"] = st.session_state.obs_graficos

if st.session_state.pop("_sync_form_ia", False):
    st.session_state["k_form_rel_ia"] = st.session_state.obs_comparativo
    st.session_state["k_comp"] = st.session_state.obs_comparativo

if st.session_state.pop("_sync_trend_ia", False):
    st.session_state["k_trend_rel_ia"] = st.session_state.obs_graficos
    st.session_state["k_comp_graficos"] = st.session_state.obs_graficos

for _wk, _orig in [
    ("k_graf_rel", "obs_graficos"),
    ("k_graf_manual", "obs_graficos"),
    ("k_form_rel_ia", "obs_comparativo"),
    ("k_comp", "obs_comparativo"),
    ("k_trend_rel_ia", "obs_graficos"),
    ("k_comp_graficos", "obs_graficos"),
]:
    if _wk not in st.session_state:
        st.session_state[_wk] = st.session_state[_orig]


# ==========================================
# 3. BARRA LATERAL (PAINEL DE COMANDO E INGESTÃO)
# ==========================================
_nuvem_ativa = _obter_config_gist()[0] is not None
_badge_nuvem = "<span style='color: #10b981; font-weight: 700;'>● Operacional (Nuvem)</span>" if _nuvem_ativa else "<span style='color: #10b981; font-weight: 700;'>● Operacional</span>"

st.sidebar.markdown(f"""
<div style="background: {T['bg_card_grad']}; border: 1px solid {T['border']}; border-left: 4px solid {T['accent']}; border-radius: 10px; padding: 14px 16px; margin-bottom: 12px; box-shadow: 0 2px 6px rgba(0,0,0,0.05);">
    <div style="font-size: 10px; font-weight: 700; color: {T['accent']}; letter-spacing: 1.2px; text-transform: uppercase;">INSPETOR CONECTADO</div>
    <div style="font-size: 15px; font-weight: 700; color: {T['txt_main']}; margin-top: 3px;">{st.session_state.nome_logado}</div>
    <div style="font-size: 11px; color: {T['txt_sec']}; margin-top: 2px;">Status: {_badge_nuvem}</div>
</div>
""", unsafe_allow_html=True)

if st.sidebar.button("ENCERRAR SESSÃO", use_container_width=True):
    st.session_state.usuario_logado = None
    st.session_state.nome_logado = None
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.markdown(f"<div style='font-size: 11px; font-weight: 700; color: {T['txt_sec']}; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 6px;'>INGESTÃO DE TELEMETRIA (.CSV)</div>", unsafe_allow_html=True)
arquivos_up = st.sidebar.file_uploader(
    "Importar Log(s) Garmin:",
    type=['csv'],
    accept_multiple_files=True,
    label_visibility="collapsed"
)

st.sidebar.markdown("---")
st.sidebar.markdown(f"<div style='font-size: 11px; font-weight: 700; color: {T['txt_sec']}; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 6px;'>CONFIGURAÇÃO DA AERONAVE</div>", unsafe_allow_html=True)

lista_perfis = list(perfis_salvos.keys())
prox_perfil = st.session_state.pop("_proximo_perfil", None)
if prox_perfil and prox_perfil in lista_perfis:
    st.session_state["perfil_ativo_sel"] = prox_perfil
elif "perfil_ativo_sel" in st.session_state and st.session_state["perfil_ativo_sel"] not in lista_perfis:
    st.session_state["perfil_ativo_sel"] = lista_perfis[0]

projeto_atual = st.sidebar.selectbox("Perfil de Envelope:", lista_perfis, key="perfil_ativo_sel")
p_dados = perfis_salvos[projeto_atual]

tipo_projeto = p_dados.get("tipo_aeronave", "Asa Fixa")
qtd_motores_config = int(p_dados.get("qtd_motores", 1))

st.sidebar.markdown(f"""
<div style="display: flex; gap: 8px; margin-top: 6px; margin-bottom: 10px;">
    <div style="flex: 1; background: {T['bg_card']}; border: 1px solid {T['border']}; border-radius: 8px; padding: 8px 10px; text-align: center;">
        <div style="font-size: 10px; color: {T['txt_sec']}; font-weight: 600; text-transform: uppercase;">Categoria</div>
        <div style="font-size: 13px; font-weight: 700; color: {T['txt_main']}; margin-top: 2px;">{tipo_projeto}</div>
    </div>
    <div style="flex: 1; background: {T['bg_card']}; border: 1px solid {T['border']}; border-radius: 8px; padding: 8px 10px; text-align: center;">
        <div style="font-size: 10px; color: {T['txt_sec']}; font-weight: 600; text-transform: uppercase;">Propulsão</div>
        <div style="font-size: 13px; font-weight: 700; color: {T['accent']}; margin-top: 2px;">{qtd_motores_config} Motor(es)</div>
    </div>
</div>
""", unsafe_allow_html=True)

st.sidebar.markdown("---")
st.sidebar.markdown(f"<div style='font-size: 11px; font-weight: 700; color: {T['txt_sec']}; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 6px;'>MODO DE MISSÃO</div>", unsafe_allow_html=True)
modo_analise = st.sidebar.radio(
    "Selecione o perfil de análise:",
    ["Voo Solo / Trend Monitoring", "Voo em Formatura (Líder e Ala)"],
    label_visibility="collapsed"
)

st.sidebar.markdown("---")
st.sidebar.markdown(f"<div style='font-size: 11px; font-weight: 700; color: {T['txt_sec']}; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 6px;'>APARÊNCIA DO PAINEL</div>", unsafe_allow_html=True)
escolha_tema = st.sidebar.radio(
    "Tema Visual:",
    ["Claro (Diurno)", "Escuro (Cockpit)"],
    index=0 if "Claro" in st.session_state.tema_ui else 1,
    label_visibility="collapsed"
)
if escolha_tema != st.session_state.tema_ui:
    st.session_state.tema_ui = escolha_tema
    st.rerun()

st.sidebar.markdown("---")

with st.sidebar.expander("Cadastrar Nova Aeronave", expanded=False):
    novo_projeto = st.text_input("Designação da Aeronave:", key="cad_anv_nome").strip()
    nova_classe = st.selectbox("Categoria:", ["Asa Fixa", "Asa Rotativa"], key="cad_anv_classe")
    qtd_mot_novo = st.number_input("Nº de Motores:", min_value=1, max_value=6, value=1, key="cad_anv_mot")
    if st.button("Salvar Novo Perfil", use_container_width=True):
        if not novo_projeto:
            st.warning("Informe a designação da aeronave.")
        elif novo_projeto in perfis_salvos:
            st.warning("Já existe um perfil com esse nome.")
        else:
            perfis_salvos[novo_projeto] = {
                "tipo_aeronave": nova_classe,
                "qtd_motores": int(qtd_mot_novo),
                "g_max": 6.0, "g_min": -3.0,
                "tq_red": 1795.0, "tq_yel": 1795.0,
                "itt_red": 790.0, "itt_yel": 740.0,
                "ng_red": 101.5, "ng_yel": 100.0,
                "nh_red": 100.0, "nh_yel": 100.0,
                "oilt_red": 99.0, "oilt_min_green": 10.0,
                "oilp_red": 95.0, "oilp_min_green": 40.0
            }
            salvar_perfis(perfis_salvos)
            st.session_state["_proximo_perfil"] = novo_projeto
            st.rerun()

with st.sidebar.expander("Editar / Excluir Perfil Ativo", expanded=False):
    st.caption("Renomeie, ajuste limites ou exclua o perfil selecionado acima.")
    e_perfil_oficial = (projeto_atual == "T-27M (Oficial)")

    nome_editado = st.text_input(
        "Nome do Perfil:",
        value=projeto_atual,
        disabled=e_perfil_oficial,
        key=f"edit_nome_{projeto_atual}"
    ).strip()

    novo_tipo_aeronave = st.selectbox(
        "Categoria Atual:",
        ["Asa Fixa", "Asa Rotativa"],
        index=0 if tipo_projeto == "Asa Fixa" else 1,
        key=f"edit_cat_{projeto_atual}"
    )
    nova_qtd_mot = st.number_input(
        "Quantidade de Motores Atual:",
        min_value=1,
        max_value=6,
        value=qtd_motores_config,
        key=f"edit_mot_{projeto_atual}"
    )
    st.markdown("---")

    def ler_limite(label, chave):
        val_ini = p_dados.get(chave, "")
        val = st.text_input(
            label,
            value="" if val_ini is None else str(val_ini),
            key=f"lim_{projeto_atual}_{chave}"
        )
        try:
            return float(val) if val.strip() != "" else None
        except:
            return None

    l_g_max = ler_limite("Força G Máxima (+)", "g_max")
    l_g_min = ler_limite("Força G Mínima (-)", "g_min")
    l_tq_red = ler_limite("Torque Red Line", "tq_red")
    l_tq_yel = ler_limite("Torque Início Arco Amarelo", "tq_yel")
    l_ng_red = ler_limite("Ng Red Line", "ng_red")
    l_ng_yel = ler_limite("Ng Início Arco Amarelo", "ng_yel")
    l_nh_red = ler_limite("Nh Red Line", "nh_red")
    l_nh_yel = ler_limite("Nh Início Arco Amarelo", "nh_yel")
    l_itt_red = ler_limite("T5 / ITT Red Line", "itt_red")
    l_itt_yel = ler_limite("T5 / ITT Início Arco Amarelo", "itt_yel")
    l_oilt_red = ler_limite("Temp Óleo Red Line", "oilt_red")
    l_oilt_min = ler_limite("Temp Óleo Mínima (Arco Verde)", "oilt_min_green")
    l_oilp_red = ler_limite("Pressão Óleo Red Line", "oilp_red")
    l_oilp_min = ler_limite("Pressão Óleo Mínima (Luz Press Óleo)", "oilp_min_green")

    if st.button("Salvar Alterações do Perfil", use_container_width=True, key="btn_salvar_perfil"):
        nome_destino = projeto_atual if e_perfil_oficial else nome_editado
        if not nome_destino:
            st.warning("O nome do perfil não pode ficar vazio.")
        elif nome_destino != projeto_atual and nome_destino in perfis_salvos:
            st.warning(f"Já existe um perfil chamado '{nome_destino}'.")
        else:
            dados_novos = {
                "tipo_aeronave": novo_tipo_aeronave,
                "qtd_motores": int(nova_qtd_mot),
                "g_max": l_g_max,
                "g_min": l_g_min,
                "tq_red": l_tq_red,
                "tq_yel": l_tq_yel,
                "itt_red": l_itt_red,
                "itt_yel": l_itt_yel,
                "ng_red": l_ng_red,
                "ng_yel": l_ng_yel,
                "nh_red": l_nh_red,
                "nh_yel": l_nh_yel,
                "oilt_red": l_oilt_red,
                "oilt_min_green": l_oilt_min,
                "oilp_red": l_oilp_red,
                "oilp_min_green": l_oilp_min
            }
            dados_novos = {k: (v if v is not None else "") for k, v in dados_novos.items()}
            if nome_destino != projeto_atual:
                del perfis_salvos[projeto_atual]
            perfis_salvos[nome_destino] = dados_novos
            salvar_perfis(perfis_salvos)
            st.session_state["_proximo_perfil"] = nome_destino
            st.rerun()

    st.markdown("---")
    if e_perfil_oficial:
        st.caption("O perfil padrão T-27M (Oficial) é protegido e não pode ser excluído.")
    else:
        conf_del = st.checkbox(
            f"Confirmar exclusão de '{projeto_atual}'",
            key=f"conf_del_{projeto_atual}"
        )
        if st.button(
            "Excluir Perfil Selecionado",
            use_container_width=True,
            disabled=not conf_del,
            key="btn_excluir_perfil"
        ):
            del perfis_salvos[projeto_atual]
            salvar_perfis(perfis_salvos)
            st.session_state["_proximo_perfil"] = "T-27M (Oficial)"
            st.rerun()


# ==========================================
# 4. CABEÇALHO EXECUTIVO E FUNÇÕES MATEMÁTICAS
# ==========================================
st.markdown(f"""
<div style="background: {T['bg_header']}; border: 1px solid {T['border']}; border-bottom: 3px solid #3b82f6; border-radius: 12px; padding: 18px 24px; margin-bottom: 18px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 15px; box-shadow: 0 4px 15px rgba(0,0,0,0.12);">
    <div>
        <div style="display: flex; align-items: center; gap: 10px;">
            <span style="background: #2563eb; color: #fff; font-size: 10px; font-weight: 800; padding: 3px 8px; border-radius: 4px; letter-spacing: 1px;">GARMIN G1000 / G3000</span>
            <span style="color: {T['sub_header']}; font-size: 12px; font-family: 'JetBrains Mono', Consolas, monospace;">FDM ENGINE v4.0 • TACTICAL EDITION</span>
        </div>
        <div style="margin: 6px 0 0 0; font-size: 26px; font-weight: 700; color: {T['txt_header']}; letter-spacing: -0.3px;">Centro de Telemetria e Inspetoria de Voo</div>
    </div>
    <div style="display: flex; gap: 12px; align-items: center;">
        <div style="background: {T['bg_badge_header']}; border: 1px solid rgba(255,255,255,0.2); border-radius: 8px; padding: 8px 16px; text-align: right;">
            <div style="font-size: 10px; color: {T['sub_header']}; text-transform: uppercase; letter-spacing: 0.8px;">Aeronave Ativa</div>
            <div style="font-size: 15px; font-weight: 700; color: #ffffff; font-family: 'JetBrains Mono', Consolas, monospace;">{projeto_atual}</div>
        </div>
        <div style="background: {T['bg_badge_header']}; border: 1px solid rgba(255,255,255,0.2); border-radius: 8px; padding: 8px 16px; text-align: right;">
            <div style="font-size: 10px; color: {T['sub_header']}; text-transform: uppercase; letter-spacing: 0.8px;">Modo Operacional</div>
            <div style="font-size: 14px; font-weight: 600; color: #ffffff;">{"Formatura Tática" if "Formatura" in modo_analise else "Solo / Trend Monitoring"}</div>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

def formatar_duracao(td):
    total_seg = int(td.total_seconds())
    if total_seg < 0:
        return "00:00:00"
    horas = total_seg // 3600
    minutos = (total_seg % 3600) // 60
    segundos = total_seg % 60
    return f"{horas:02d}:{minutos:02d}:{segundos:02d}"

@st.cache_data
def processar_dados(file_bytes):
    df = pd.read_csv(io.BytesIO(file_bytes), skiprows=2, low_memory=False)
    df.columns = df.columns.str.strip()

    if 'Lcl Date' in df.columns:
        df['Lcl Date'] = df['Lcl Date'].astype(str).str.strip()
    if 'Lcl Time' in df.columns:
        df['Lcl Time'] = df['Lcl Time'].astype(str).str.strip()

    if 'Lcl Date' in df.columns and 'Lcl Time' in df.columns:
        df['Datetime'] = pd.to_datetime(df['Lcl Date'] + ' ' + df['Lcl Time'], errors='coerce')
    else:
        df['Datetime'] = pd.to_datetime(df['Lcl Time'], errors='coerce')

    colunas_numericas = [
        'AltMSL', 'IAS', 'Pitch', 'Roll', 'NormAc', 'LatAc',
        'Latitude', 'Longitude'
    ]

    for i in range(1, 7):
        colunas_numericas.extend([
            f'E{i} Torq ft lbs', f'E{i} ITT', f'E{i} NP %',
            f'E{i} NG', f'E{i} FFlow', f'E{i} OilP', f'E{i} OilT'
        ])

    for col in colunas_numericas:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')

    for i in range(1, 7):
        if f'E{i} OilT' in df.columns:
            df[f'E{i} OilT_Celsius'] = (df[f'E{i} OilT'] - 32) / 1.8

    df = df.dropna(subset=['Datetime']).sort_values('Datetime')
    return df

def formula_haversine(lat1, lon1, lat2, lon2):
    phi1 = np.deg2rad(lat1)
    phi2 = np.deg2rad(lat2)
    dphi = np.deg2rad(lat2 - lat1)
    dlambda = np.deg2rad(lon2 - lon1)

    a = np.sin(dphi / 2.0)**2 + np.cos(phi1) * np.cos(phi2) * np.sin(dlambda / 2.0)**2
    c = 2.0 * np.arctan2(np.sqrt(a), np.sqrt(1.0 - a))

    R = 6371000.0
    return R * c

def verificar_temporizador(df, coluna, limite_valor, segundos_necessarios):
    if coluna not in df.columns:
        return False
    acima_limite = df[coluna] > limite_valor
    consecutivos = acima_limite.groupby((~acima_limite).cumsum()).sum()
    return consecutivos.max() >= segundos_necessarios

def calcular_minutos_voo(df):
    if 'IAS' in df.columns:
        voo = df[df['IAS'] > 40]
        if not voo.empty:
            t0 = voo['Datetime'].iloc[0]
            return (df['Datetime'] - t0).dt.total_seconds() / 60.0
    return (df['Datetime'] - df['Datetime'].iloc[0]).dt.total_seconds() / 60.0


# ==========================================
# 5. GERADORES DE INTERFACE AVIONICS E RELATÓRIOS (.DOCX)
# ==========================================
def exibir_card_arcos(titulo, pico, red_line, yellow_line, min_green, unidade, formato=":.1f", e_minimo=False):
    cor = "#10b981"
    bg_badge = "rgba(16, 185, 129, 0.14)"
    status = "ARCO VERDE"

    if red_line is None:
        cor = "#64748b"
        bg_badge = "rgba(100, 116, 139, 0.15)"
        status = "SEM LIMITE"
    else:
        if e_minimo:
            if pico < red_line:
                cor, bg_badge, status = "#ef4444", "rgba(239, 68, 68, 0.18)", "LINHA VERMELHA"
            elif yellow_line is not None and pico <= yellow_line:
                cor, bg_badge, status = "#f59e0b", "rgba(245, 158, 11, 0.18)", "ARCO AMARELO"
            elif min_green is not None and pico < min_green:
                cor, bg_badge, status = "#f59e0b", "rgba(245, 158, 11, 0.18)", "ARCO AMARELO"
        else:
            if pico >= red_line:
                cor, bg_badge, status = "#ef4444", "rgba(239, 68, 68, 0.18)", "LINHA VERMELHA"
            elif yellow_line is not None and pico >= yellow_line:
                cor, bg_badge, status = "#f59e0b", "rgba(245, 158, 11, 0.18)", "ARCO AMARELO"
            elif min_green is not None and pico < min_green:
                cor, bg_badge, status = "#f59e0b", "rgba(245, 158, 11, 0.18)", "ARCO AMARELO"

    pico_str = ("{"+formato+"}").format(pico)
    limite_str = ("{"+formato+"}").format(red_line) if red_line is not None else 'N/A'

    pct_barra = 50
    if red_line is not None and red_line != 0:
        if e_minimo:
            pct_barra = min(100, max(8, int((abs(pico) / abs(red_line)) * 100)))
        else:
            pct_barra = min(100, max(8, int((pico / red_line) * 100)))

    html = f"""
    <div style="background: {T['bg_card_grad']}; padding: 16px 18px; border-radius: 12px; border: 1px solid {T['border']}; border-top: 4px solid {cor}; margin-bottom: 16px; box-shadow: 0 4px 12px rgba(0,0,0,0.08); position: relative; overflow: hidden;">
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
            <span style="font-size: 11.5px; color: {T['txt_sec']}; font-weight: 700; text-transform: uppercase; letter-spacing: 0.6px;">{titulo}</span>
            <span style="font-size: 10px; color: {cor}; background: {bg_badge}; border: 1px solid {cor}55; padding: 2px 8px; border-radius: 12px; font-weight: 700; letter-spacing: 0.5px;">{status}</span>
        </div>
        <div style="display: flex; align-items: baseline; gap: 6px; margin-bottom: 12px;">
            <span style="color: {T['txt_main']}; font-size: 30px; font-weight: 700; font-family: 'JetBrains Mono', Consolas, monospace; letter-spacing: -0.5px;">{pico_str}</span>
            <span style="color: {T['txt_sec']}; font-size: 13px; font-weight: 600;">{unidade}</span>
        </div>
        <div style="width: 100%; background: {T['border']}; height: 5px; border-radius: 4px; overflow: hidden; margin-bottom: 8px;">
            <div style="width: {pct_barra}%; background: {cor}; height: 100%; border-radius: 4px;"></div>
        </div>
        <div style="display: flex; justify-content: space-between; align-items: center;">
            <span style="font-size: 11px; color: {T['txt_muted']}; font-weight: 600;">Red Line Limit</span>
            <span style="font-size: 11.5px; color: {T['txt_main']}; font-family: 'JetBrains Mono', Consolas, monospace; font-weight: 700;">{limite_str} {unidade}</span>
        </div>
    </div>
    """
    st.markdown(html, unsafe_allow_html=True)

def render_html_table(df_formatted, height=420):
    modo_claro = T['modo'] == 'light'
    bg_cont = "#ffffff" if modo_claro else "#0f243e"
    bg_th = "#e6f2ff" if modo_claro else "#091526"
    col_th = "#1e3a8a" if modo_claro else "#93c5fd"
    col_td = "#0f172a" if modo_claro else "#e2e8f0"
    borda_td = "#e2e8f0" if modo_claro else "#1e3a5f"
    even_row = "#f8fbff" if modo_claro else "rgba(19, 44, 78, 0.45)"

    html_table = df_formatted.to_html(index=False, justify='center', border=0, classes='tabela-grande')
    html_completo = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&family=JetBrains+Mono:wght@400;600&display=swap" rel="stylesheet">
    <style>
    body {{ background-color: transparent; color: {col_td}; font-family: 'Inter', 'Segoe UI', sans-serif; margin: 0; }}
    .tabela-container {{ max-height: {height - 20}px; overflow-y: auto; overflow-x: auto; border: 1px solid {T['border']}; border-radius: 10px; background: {bg_cont}; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }}
    .tabela-container::-webkit-scrollbar {{ width: 8px; height: 8px; }}
    .tabela-container::-webkit-scrollbar-track {{ background: {bg_cont}; }}
    .tabela-container::-webkit-scrollbar-thumb {{ background: #64748b; border-radius: 4px; }}
    .tabela-grande {{ width: 100%; font-size: 13.5px; text-align: center; border-collapse: collapse; white-space: nowrap; }}
    .tabela-grande th {{ background: {bg_th}; color: {col_th}; font-size: 11.5px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.7px; padding: 13px 14px; position: sticky; top: 0; z-index: 1; border-bottom: 2px solid {T['border']}; }}
    .tabela-grande td {{ padding: 10px 14px; border-bottom: 1px solid {borda_td}; font-family: 'JetBrains Mono', Consolas, monospace; color: {col_td}; font-weight: 500; }}
    .tabela-grande tr:nth-child(even) {{ background-color: {even_row}; }}
    .tabela-grande tr:hover {{ background-color: rgba(59, 130, 246, 0.14); }}
    </style>
    </head>
    <body>
        <div class="tabela-container">{html_table}</div>
    </body>
    </html>
    """
    components.html(html_completo, height=height)

def exibir_tabela_grande(df_view):
    traducao = {
        'Lcl Time': 'Hora Local', 'AltMSL': 'Altitude (ft)', 'IAS': 'Velocidade (kt)',
        'NormAc': 'Força G', 'Pitch': 'Arfagem (º)', 'Roll': 'Rolagem (º)'
    }
    for i in range(1, 7):
        traducao[f'E{i} Torq ft lbs'] = f'Torque M{i} (Lbs)'
        traducao[f'E{i} ITT'] = f'T5/ITT M{i} (ºC)'
        traducao[f'E{i} NG'] = f'Ng M{i} (%)'
        traducao[f'E{i} NP %'] = f'Nh M{i} (%)'
        traducao[f'E{i} OilT_Celsius'] = f'T.Óleo M{i} (ºC)'
        traducao[f'E{i} OilP'] = f'P.Óleo M{i} (PSI)'

    df_traduzido = df_view.rename(columns=traducao)
    for col in df_traduzido.columns:
        if 'Altitude' in col or 'Torque' in col or 'ITT' in col or 'T.Óleo' in col or 'P.Óleo' in col:
            df_traduzido[col] = df_traduzido[col].apply(lambda x: f"{x:.0f}" if pd.notnull(x) else "-")
        elif 'Velocidade' in col or 'Ng' in col or 'Nh' in col or 'Arfagem' in col or 'Rolagem' in col:
            df_traduzido[col] = df_traduzido[col].apply(lambda x: f"{x:.1f}" if pd.notnull(x) else "-")
        elif 'Força G' in col:
            df_traduzido[col] = df_traduzido[col].apply(lambda x: f"{x:.2f}" if pd.notnull(x) else "-")
    render_html_table(df_traduzido)

def injetar_graficos_word(doc, graficos_dict):
    if any(fig is not None for fig in graficos_dict.values()):
        doc.add_heading("ANEXOS GRÁFICOS", level=1)
        for titulo_grafico, fig in graficos_dict.items():
            if fig is not None:
                try:
                    doc.add_heading(titulo_grafico, level=2)
                    img_bytes = fig.to_image(format="png", width=900, height=500, scale=2)
                    doc.add_picture(io.BytesIO(img_bytes), width=Inches(6.0))
                except Exception:
                    doc.add_paragraph(f"[Aviso: Não foi possível exportar a imagem '{titulo_grafico}'].")


# ==========================================
# 6. GERADOR WORD OFICIAL (PADRÃO FAB / LIBREOFFICE)
# ==========================================
def set_cell_background(cell, hex_color):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def adicionar_paragrafo_com_markdown(doc_or_cell, texto, estilo=None, bullet=False):
    if bullet:
        p = doc_or_cell.add_paragraph(style='List Bullet')
    elif estilo:
        p = doc_or_cell.add_paragraph(style=estilo)
    else:
        p = doc_or_cell.add_paragraph()

    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15

    partes = re.split(r'(\*\*.*?\*\*)', texto)
    for parte in partes:
        if parte.startswith('**') and parte.endswith('**'):
            texto_limpo = parte[2:-2]
            run = p.add_run(texto_limpo)
            run.bold = True
        else:
            p.add_run(parte)
    return p

def _estilizar_bordas_tabela(table, cor_borda="CBD5E1"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="6" w:space="0" w:color="1B365D"/>'
        f'<w:bottom w:val="single" w:sz="6" w:space="0" w:color="1B365D"/>'
        f'<w:left w:val="none"/>'
        f'<w:right w:val="none"/>'
        f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="{cor_borda}"/>'
        f'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="{cor_borda}"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

def _preencher_celula_formatada(cell, texto, alinhamento, cor_texto, tamanho_pt=9.0, forcar_negrito=False):
    p = cell.paragraphs[0]
    p.alignment = alinhamento
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.08

    # Centraliza verticalmente o conteúdo na célula
    tcPr = cell._tc.get_or_add_tcPr()
    v_align = OxmlElement('w:vAlign')
    v_align.set(qn('w:val'), 'center')
    tcPr.append(v_align)

    partes = re.split(r'(\*\*.*?\*\*)', str(texto))
    for parte in partes:
        if not parte:
            continue
        if parte.startswith('**') and parte.endswith('**'):
            run = p.add_run(parte[2:-2])
            run.bold = True
        else:
            run = p.add_run(parte)
            if forcar_negrito:
                run.bold = True
        run.font.size = Pt(tamanho_pt)
        run.font.color.rgb = cor_texto

def _inserir_tabela_markdown_word(doc, linhas_brutas):
    linhas_limpas = []
    for ln in linhas_brutas:
        # Ignora a linha separadora do Markdown (|---|---|---|)
        if re.match(r'^[\|\s\-:]+$', ln):
            continue
        celulas = [c.strip() for c in ln.strip('|').split('|')]
        if any(celulas):
            linhas_limpas.append(celulas)

    if not linhas_limpas:
        return

    n_cols = max(len(r) for r in linhas_limpas)
    linhas_norm = [r + [""] * (n_cols - len(r)) for r in linhas_limpas]

    table = doc.add_table(rows=len(linhas_norm), cols=n_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    _estilizar_bordas_tabela(table)

    # Calcula largura proporcional somando 17.0 cm (largura útil da folha A4)
    pesos = []
    for c_idx in range(n_cols):
        tam_med = max(len(re.sub(r'\*\*', '', r[c_idx])) for r in linhas_norm)
        pesos.append(max(tam_med, 8))
    soma_pesos = sum(pesos)
    larguras_cm = [max(2.2, (p / soma_pesos) * 17.0) for p in pesos]
    fator = 17.0 / sum(larguras_cm)
    larguras_cm = [w * fator for w in larguras_cm]

    COR_BRANCO = RGBColor(255, 255, 255)
    COR_TEXTO_TAB = RGBColor(30, 41, 59)

    for r_idx, linha_dados in enumerate(linhas_norm):
        row = table.rows[r_idx]
        trPr = row._tr.get_or_add_trPr()
        trPr.append(OxmlElement('w:cantSplit'))

        e_cabecalho = (r_idx == 0)
        if e_cabecalho:
            trPr.append(OxmlElement('w:tblHeader'))

        bg_hex = "1B365D" if e_cabecalho else ("F0F4F8" if r_idx % 2 == 0 else "FFFFFF")

        for c_idx, valor in enumerate(linha_dados):
            cell = row.cells[c_idx]
            cell.width = Cm(larguras_cm[c_idx])
            set_cell_background(cell, bg_hex)
            set_cell_margins(cell, top=75, bottom=75, left=100, right=100)

            alinh = WD_ALIGN_PARAGRAPH.LEFT if (c_idx == 0 and not e_cabecalho) else WD_ALIGN_PARAGRAPH.CENTER
            _preencher_celula_formatada(
                cell,
                valor,
                alinhamento=alinh,
                cor_texto=COR_BRANCO if e_cabecalho else COR_TEXTO_TAB,
                tamanho_pt=9.5 if e_cabecalho else 9.0,
                forcar_negrito=e_cabecalho
            )

    p_esp = doc.add_paragraph()
    p_esp.paragraph_format.space_before = Pt(0)
    p_esp.paragraph_format.space_after = Pt(6)

def renderizar_markdown_no_word(doc, texto, cor_titulo=RGBColor(27, 54, 93)):
    """Converte títulos (#, ##, ###), tabelas Markdown (|...|), bullets (*, -) e negrito (**) para formatação nativa do Word."""
    if not str(texto).strip():
        return

    linhas = str(texto).split('\n')
    i = 0
    while i < len(linhas):
        l_limpa = linhas[i].strip()
        if not l_limpa:
            i += 1
            continue

        # Detecta bloco de tabela Markdown (linhas começando com |)
        if l_limpa.startswith("|") and l_limpa.count("|") >= 2:
            bloco_tab = []
            while i < len(linhas):
                ln_tab = linhas[i].strip()
                if ln_tab.startswith("|") and ln_tab.count("|") >= 2:
                    bloco_tab.append(ln_tab)
                    i += 1
                else:
                    break
            _inserir_tabela_markdown_word(doc, bloco_tab)
            continue

        if l_limpa.startswith("# "):
            tit = l_limpa[2:].strip()
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.keep_with_next = True
            r = p.add_run(tit)
            r.bold = True
            r.font.size = Pt(12)
            r.font.color.rgb = cor_titulo

        elif l_limpa.startswith("## "):
            subtit = l_limpa[3:].strip()
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.keep_with_next = True
            r = p.add_run(subtit)
            r.bold = True
            r.font.size = Pt(10.5)
            r.font.color.rgb = cor_titulo

        elif l_limpa.startswith("### "):
            subtit3 = l_limpa[4:].strip()
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.keep_with_next = True
            r = p.add_run(subtit3)
            r.bold = True
            r.font.size = Pt(10)
            r.font.color.rgb = cor_titulo

        elif l_limpa.startswith("---"):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(4)
            r = p.add_run("_________________________________________________________________________________")
            r.font.color.rgb = RGBColor(200, 200, 200)

        elif l_limpa.startswith("* ") or l_limpa.startswith("- "):
            conteudo_item = l_limpa[2:].strip()
            adicionar_paragrafo_com_markdown(doc, conteudo_item, bullet=True)

        else:
            adicionar_paragrafo_com_markdown(doc, l_limpa)

        i += 1

def gerar_docx_solo_editado(tempos_dict, alertas_texto, obs_dict, graficos_dict):
    doc = Document()

    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)

    COR_AZUL_FAB = RGBColor(27, 54, 93)
    COR_CINZA_TEXTO = RGBColor(51, 51, 51)
    HEX_CINZA_CLARO = "F0F4F8"

    p_topo = doc.add_paragraph()
    p_topo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_topo1 = p_topo.add_run("COMANDO DA AERONÁUTICA\n")
    r_topo1.bold = True
    r_topo1.font.size = Pt(11)
    r_topo1.font.color.rgb = COR_AZUL_FAB

    r_topo2 = p_topo.add_run("DEPARTAMENTO DE DELEGAÇÃO E INSPEÇÃO TÉCNICA DE MANUTENÇÃO\nESQUADRÃO DE TELEMETRIA E ENSAIOS DE VOO\n")
    r_topo2.font.size = Pt(9)
    r_topo2.font.color.rgb = RGBColor(100, 100, 100)

    p_titulo_doc = doc.add_paragraph()
    p_titulo_doc.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_titulo_doc.paragraph_format.space_before = Pt(8)
    p_titulo_doc.paragraph_format.space_after = Pt(14)
    r_tit = p_titulo_doc.add_run("RELATÓRIO TÉCNICO DE TELEMETRIA E INSPETORIA (RTT)")
    r_tit.bold = True
    r_tit.font.size = Pt(14)
    r_tit.font.color.rgb = COR_AZUL_FAB

    tabela_meta = doc.add_table(rows=3, cols=2)
    tabela_meta.alignment = WD_TABLE_ALIGNMENT.CENTER
    tabela_meta.autofit = False

    larguras = [Cm(8.5), Cm(8.5)]
    for row in tabela_meta.rows:
        for i, cell in enumerate(row.cells):
            cell.width = larguras[i]
            set_cell_margins(cell, top=80, bottom=80, left=120, right=120)

    dados_matriz = [
        [f"**Aeronave/Perfil:** {projeto_atual}", f"**Data do Voo:** {datetime.now().strftime('%d/%m/%Y')}"],
        [f"**Inspetor / Analista:** {st.session_state.nome_logado}", f"**Tempo de Voo:** {tempos_dict.get('t_voo', 'N/A')}"],
        [f"**Acionamento:** {tempos_dict.get('acionamento', 'N/A')}", f"**Corte do Motor:** {tempos_dict.get('corte', 'N/A')}"]
    ]

    for r_idx, linha in enumerate(dados_matriz):
        for c_idx, valor in enumerate(linha):
            cell = tabela_meta.rows[r_idx].cells[c_idx]
            set_cell_background(cell, HEX_CINZA_CLARO)
            adicionar_paragrafo_com_markdown(cell, valor)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    p_s1 = doc.add_paragraph()
    p_s1.paragraph_format.space_before = Pt(12)
    p_s1.paragraph_format.space_after = Pt(6)
    r_s1 = p_s1.add_run("1. AUDITORIA DE ALERTAS E TRANSPARÊNCIA OPERACIONAL")
    r_s1.bold = True
    r_s1.font.size = Pt(11)
    r_s1.font.color.rgb = COR_AZUL_FAB

    linhas_alerta = [linha for linha in alertas_texto.split('\n') if linha.strip()]
    if linhas_alerta:
        tab_alertas = doc.add_table(rows=1 + len(linhas_alerta), cols=1)
        tab_alertas.alignment = WD_TABLE_ALIGNMENT.CENTER
        hdr_cell = tab_alertas.rows[0].cells[0]
        hdr_cell.width = Cm(17.0)
        set_cell_background(hdr_cell, "1B365D")
        set_cell_margins(hdr_cell, top=80, bottom=80, left=120, right=120)
        p_hdr = hdr_cell.paragraphs[0]
        r_hdr = p_hdr.add_run(" REGISTRO DE EVENTOS E ULTRAPASSAGENS DE LIMITES DA AERONAVE")
        r_hdr.bold = True
        r_hdr.font.color.rgb = RGBColor(255, 255, 255)
        r_hdr.font.size = Pt(9.5)

        for idx, al in enumerate(linhas_alerta):
            cell_al = tab_alertas.rows[idx + 1].cells[0]
            cell_al.width = Cm(17.0)
            set_cell_margins(cell_al, top=60, bottom=60, left=120, right=120)
            bg_cor = "FFF5F5" if ("[CRÍTICO]" in al or "LINHA VERMELHA" in al) else ("FFFBEB" if "[ALERTA]" in al else "FFFFFF")
            set_cell_background(cell_al, bg_cor)
            adicionar_paragrafo_com_markdown(cell_al, al)
    else:
        adicionar_paragrafo_com_markdown(doc, "**Operação Normal:** Nenhum alerta ou ultrapassagem de parâmetro registrado nesta missão.")

    p_s2 = doc.add_paragraph()
    p_s2.paragraph_format.space_before = Pt(14)
    p_s2.paragraph_format.space_after = Pt(6)
    r_s2 = p_s2.add_run("2. ANÁLISE DE TELEMETRIA E PERFORMANCE DO VOO")
    r_s2.bold = True
    r_s2.font.size = Pt(11)
    r_s2.font.color.rgb = COR_AZUL_FAB

    texto_ia = obs_dict.get("Inspetoria e Graficos", "")
    renderizar_markdown_no_word(doc, texto_ia, COR_AZUL_FAB)

    for chave_extra in ["Resumo de Painel e Rota", "Avaliador de Manobras e Pouso"]:
        val_extra = str(obs_dict.get(chave_extra, "")).strip()
        if val_extra:
            p_ex = doc.add_paragraph()
            p_ex.paragraph_format.space_before = Pt(8)
            p_ex.paragraph_format.space_after = Pt(4)
            r_ex = p_ex.add_run(f"Apontamentos Complementares — {chave_extra}")
            r_ex.bold = True
            r_ex.font.size = Pt(10.5)
            r_ex.font.color.rgb = COR_AZUL_FAB
            renderizar_markdown_no_word(doc, val_extra, COR_AZUL_FAB)

    for titulo_grafico, fig in graficos_dict.items():
        if fig is not None:
            p_s3 = doc.add_paragraph()
            p_s3.paragraph_format.space_before = Pt(14)
            p_s3.paragraph_format.space_after = Pt(6)
            r_s3 = p_s3.add_run(f"3. ANEXO DE TELEMETRIA VISUAL ({titulo_grafico.upper()})")
            r_s3.bold = True
            r_s3.font.size = Pt(11)
            r_s3.font.color.rgb = COR_AZUL_FAB

            try:
                img_bytes = fig.to_image(format="png", width=900, height=500, scale=2)
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.paragraph_format.space_before = Pt(6)
                p_img.paragraph_format.space_after = Pt(4)
                p_img.add_run().add_picture(io.BytesIO(img_bytes), width=Cm(16.5))
            except Exception:
                doc.add_paragraph(f"[Aviso: Não foi possível exportar a imagem '{titulo_grafico}'].")

    p_esp = doc.add_paragraph()
    p_esp.paragraph_format.space_before = Pt(20)

    p_ass = doc.add_paragraph()
    p_ass.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_linha = p_ass.add_run("_____________________________________________________\n")
    r_linha.bold = True
    r_linha.font.color.rgb = COR_AZUL_FAB

    r_nome = p_ass.add_run(f"{st.session_state.nome_logado}\n")
    r_nome.bold = True
    r_nome.font.size = Pt(10)
    r_nome.font.color.rgb = COR_CINZA_TEXTO

    r_cargo = p_ass.add_run("Inspetor de Telemetria e Manutenção Aeronáutica\nEsquadrão de Ensaio de Voo / FAB")
    r_cargo.font.size = Pt(8.5)
    r_cargo.font.color.rgb = RGBColor(100, 100, 100)

    io_bytes = io.BytesIO()
    doc.save(io_bytes)
    return io_bytes.getvalue()

def gerar_docx_formatura_editado(arq_lider, arq_ala, obs_dict, graficos_dict):
    doc = Document()
    COR_AZUL_FAB = RGBColor(27, 54, 93)

    titulo = doc.add_heading("RELATÓRIO TÁTICO DE FORMATURA", 0)
    titulo.alignment = 1
    data_hoje = datetime.now().strftime('%d/%m/%Y')
    doc.add_paragraph(f"Perfil: {projeto_atual} | Analista: {st.session_state.nome_logado} | Data: {data_hoje}").alignment = 1

    doc.add_heading("1. AERONAVES ENVOLVIDAS", level=1)
    adicionar_paragrafo_com_markdown(doc, f"**Aeronave Líder (Nº 1):** {arq_lider}")
    adicionar_paragrafo_com_markdown(doc, f"**Aeronave Ala (Nº 2):** {arq_ala}")

    doc.add_heading("2. ANÁLISE TÉCNICA E PERFORMANCE DA FORMATURA", level=1)
    tem_obs = False
    for titulo_obs, texto in obs_dict.items():
        if str(texto).strip() != "":
            tem_obs = True
            doc.add_heading(f"{titulo_obs}", level=2)
            renderizar_markdown_no_word(doc, texto, COR_AZUL_FAB)
    if not tem_obs:
        doc.add_paragraph("Nenhuma observação inserida no dossiê.")

    injetar_graficos_word(doc, graficos_dict)

    doc.add_paragraph("\n\n____________________________________________").alignment = 1
    doc.add_paragraph(f"{st.session_state.nome_logado} - Especialista em Telemetria").alignment = 1

    io_bytes = io.BytesIO()
    doc.save(io_bytes)
    return io_bytes.getvalue()

def gerar_docx_comparativo_editado(arq_a, arq_b, df_metricas, obs_dict, graficos_dict):
    doc = Document()
    COR_AZUL_FAB = RGBColor(27, 54, 93)

    titulo = doc.add_heading("RELATÓRIO DE TREND MONITORING (PERFORMANCE E DEGRADAÇÃO)", 0)
    titulo.alignment = 1

    data_hoje = datetime.now().strftime('%d/%m/%Y')
    doc.add_paragraph(f"Perfil: {projeto_atual} | Analista: {st.session_state.nome_logado} | Data: {data_hoje}").alignment = 1

    doc.add_heading("1. VOOS AVALIADOS", level=1)
    adicionar_paragrafo_com_markdown(doc, f"**Voo A (Referência / Baseline):** {arq_a}")
    adicionar_paragrafo_com_markdown(doc, f"**Voo B (Recente / Avaliado):** {arq_b}")

    doc.add_heading("2. MATRIZ COMPARATIVA DE PICOS (DELTAS)", level=1)
    table = doc.add_table(rows=1, cols=4)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Parâmetro (Picos)'
    hdr_cells[1].text = 'Voo A'
    hdr_cells[2].text = 'Voo B'
    hdr_cells[3].text = 'Variação (Delta)'

    for _, row in df_metricas.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text = str(row['Parâmetro'])
        row_cells[1].text = str(row['Voo A'])
        row_cells[2].text = str(row['Voo B'])
        row_cells[3].text = str(row['Variação (Delta)'])

        texto_delta = str(row['Variação (Delta)'])
        if "+" in texto_delta and ("Temp" in str(row['Parâmetro']) or "Torque" in str(row['Parâmetro'])):
            row_cells[3].paragraphs[0].runs[0].font.color.rgb = RGBColor(220, 0, 0)
        elif "-" in texto_delta and ("Temp" in str(row['Parâmetro']) or "Torque" in str(row['Parâmetro'])):
            row_cells[3].paragraphs[0].runs[0].font.color.rgb = RGBColor(0, 150, 0)

    doc.add_heading("3. ESTUDO DETALHADO DE PERFORMANCE E DEGRADAÇÃO", level=1)
    tem_obs = False
    for titulo_obs, texto in obs_dict.items():
        if str(texto).strip() != "":
            tem_obs = True
            doc.add_heading(f"{titulo_obs}", level=2)
            renderizar_markdown_no_word(doc, texto, COR_AZUL_FAB)

    if not tem_obs:
        doc.add_paragraph("Nenhuma observação inserida.")

    injetar_graficos_word(doc, graficos_dict)

    doc.add_paragraph("\n\n____________________________________________").alignment = 1
    doc.add_paragraph(f"{st.session_state.nome_logado} - Especialista em Telemetria").alignment = 1

    io_bytes = io.BytesIO()
    doc.save(io_bytes)
    return io_bytes.getvalue()

# ==========================================
# 7. FLUXO PRINCIPAL
# ==========================================
if not arquivos_up:
    st.markdown(f"""
    <div style="background: {T['bg_card_grad']}; border: 1px solid {T['border']}; border-left: 4px solid {T['accent']}; border-radius: 12px; padding: 42px 28px; text-align: center; margin-top: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.06);">
        <div style="font-size: 11px; font-weight: 700; color: {T['accent']}; letter-spacing: 1.5px; text-transform: uppercase; margin-bottom: 8px;">ESTAÇÃO PRONTA PARA RECEBIMENTO</div>
        <div style="margin: 0; color: {T['txt_main']}; font-size: 20px; font-weight: 700;">Aguardando Ingestão de Telemetria (.CSV)</div>
        <p style="color: {T['txt_sec']}; font-size: 13.5px; max-width: 560px; margin: 10px auto 0 auto;">
            Utilize o painel à esquerda em <b>INGESTÃO DE TELEMETRIA (.CSV)</b> para carregar <b>1 arquivo</b> (Auditoria de Voo Solo) ou <b>2 arquivos</b> (Formatura Tática / Trend Monitoring de Motores).
        </p>
    </div>
    """, unsafe_allow_html=True)

if arquivos_up:

    # ---------------------------------------------------------
    # FLUXO 1: VOO EM FORMATURA
    # ---------------------------------------------------------
    if modo_analise == "Voo em Formatura (Líder e Ala)":
        if len(arquivos_up) != 2:
            st.warning("Para o Modo Formatura, carregue exatamente 2 arquivos simultaneamente (Líder e Ala) na barra lateral.")
        else:
            st.markdown("### Designação de Elemento / Esquadrilha")

            nome_arq1 = arquivos_up[0].name
            nome_arq2 = arquivos_up[1].name

            col_lider, col_ala, col_sync = st.columns([1.5, 1.5, 1])
            with col_lider:
                arq_lider = st.selectbox("Aeronave LÍDER (Nº 1):", [nome_arq1, nome_arq2], index=0)
            with col_ala:
                idx_padrao_ala = 1 if arq_lider == nome_arq1 else 0
                arq_ala = st.selectbox("Aeronave ALA (Nº 2):", [nome_arq1, nome_arq2], index=idx_padrao_ala)
            with col_sync:
                ajuste_horas = st.number_input("Offset Fuso Horário Ala (h):", min_value=-12, max_value=12, value=0, step=1, help="Use -3 ou +3 caso uma aeronave esteja em UTC e outra em Hora Local.")

            if arq_lider == arq_ala:
                st.error("Você selecionou o mesmo log para Líder e Ala. Altere um dos seletores acima.")
            else:
                idx_lider = 0 if arq_lider == nome_arq1 else 1
                idx_ala = 1 if idx_lider == 0 else 0

                df1 = processar_dados(arquivos_up[idx_lider].getvalue())
                df2 = processar_dados(arquivos_up[idx_ala].getvalue())

                if ajuste_horas != 0:
                    df2['Datetime'] = df2['Datetime'] + pd.Timedelta(hours=ajuste_horas)

                df_merge = pd.merge_asof(
                    df1, df2,
                    on='Datetime',
                    suffixes=('_lider', '_ala'),
                    tolerance=pd.Timedelta(seconds=10),
                    direction='nearest'
                )

                df_merge = df_merge.dropna(subset=['AltMSL_lider', 'AltMSL_ala'])

                if df_merge.empty:
                    st.error("Os relógios não coincidem. Ajuste o 'Offset Fuso Horário Ala' acima para -3 ou +3.")
                else:
                    if all(c in df_merge.columns for c in ['Latitude_lider', 'Longitude_lider', 'Latitude_ala', 'Longitude_ala']):
                        df_merge['Distancia_m'] = formula_haversine(
                            df_merge['Latitude_lider'], df_merge['Longitude_lider'],
                            df_merge['Latitude_ala'], df_merge['Longitude_ala']
                        )
                    else:
                        df_merge['Distancia_m'] = np.nan

                    aba_dist, aba_comp, aba_3d, aba_rel_form = st.tabs([
                        "1. SEPARAÇÃO TÁTICA",
                        "2. COMPARATIVO DINÂMICO",
                        "3. ENVELOPE 3D",
                        "4. DOSSIÊ FINAL"
                    ])

                    fig_dist = None
                    fig_comp = None
                    fig_3d = None

                    with aba_dist:
                        fig_dist = make_subplots(rows=3, cols=1, shared_xaxes=True, vertical_spacing=0.07, subplot_titles=("1. Separação Horizontal Líder-Ala (Metros)", "2. Perfil Vertical de Altitude MSL (ft)", "3. Atitude Longitudinal / Pitch (º)"))
                        fig_dist.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['Distancia_m'], name="Separação (m)", fill='tozeroy', fillcolor='rgba(56, 189, 248, 0.12)', line=dict(color='#0284c7', width=2)), row=1, col=1)
                        if 'AltMSL_lider' in df_merge.columns: fig_dist.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['AltMSL_lider'], name="Alt Líder", line=dict(color='#ea580c', width=2)), row=2, col=1)
                        if 'AltMSL_ala' in df_merge.columns: fig_dist.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['AltMSL_ala'], name="Alt Ala", line=dict(color='#2563eb', width=2, dash='dot')), row=2, col=1)
                        if 'Pitch_lider' in df_merge.columns: fig_dist.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['Pitch_lider'], name="Pitch Líder", line=dict(color='#ea580c', width=2)), row=3, col=1)
                        if 'Pitch_ala' in df_merge.columns: fig_dist.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['Pitch_ala'], name="Pitch Ala", line=dict(color='#2563eb', width=2, dash='dot')), row=3, col=1)
                        fig_dist.update_layout(template=T['plotly_template'], paper_bgcolor=T['plotly_bg'], plot_bgcolor=T['plotly_bg'], height=820, hovermode="x unified", margin=dict(l=15, r=15, t=35, b=15))
                        st.plotly_chart(fig_dist, use_container_width=True)

                        st.subheader("Top 15 Momentos de Maior Afastamento na Formatura")
                        df_top_dist = df_merge.sort_values(by='Distancia_m', ascending=False).head(15).copy()
                        df_formatada = pd.DataFrame()
                        df_formatada['Hora (Líder)'] = df_top_dist['Datetime'].dt.strftime('%H:%M:%S')
                        df_formatada['Distância (m)'] = df_top_dist['Distancia_m'].apply(lambda x: f"{x:.1f} m" if pd.notnull(x) else "-")
                        if 'AltMSL_lider' in df_top_dist.columns: df_formatada['Alt. Líder (ft)'] = df_top_dist['AltMSL_lider'].apply(lambda x: f"{x:.0f}" if pd.notnull(x) else "-")
                        if 'Pitch_lider' in df_top_dist.columns: df_formatada['Arfagem Líder (º)'] = df_top_dist['Pitch_lider'].apply(lambda x: f"{x:.1f}" if pd.notnull(x) else "-")
                        render_html_table(df_formatada)

                        st.divider()
                        st.markdown("**Observações Táticas - Separação e Posicionamento**")
                        obs_dist_tmp = st.text_area("Registar apontamentos sobre manutenção de posição:", value=st.session_state.obs_distancia, key='k_dist')
                        if st.button("SALVAR APONTAMENTO", key="b_dist"):
                            st.session_state.obs_distancia = obs_dist_tmp
                            st.success("Apontamento gravado.")

                    with aba_comp:
                        fig_comp = make_subplots(rows=6, cols=1, shared_xaxes=True, vertical_spacing=0.03, subplot_titles=("1. Altitude MSL (ft)", "2. Fator de Carga (Força G)", "3. Torque (lb.ft)", "4. Temperatura T5 / ITT (ºC)", "5. Rotação Ng (%)", "6. Rotação Hélice Nh (%)"))
                        if 'AltMSL_lider' in df_merge.columns: fig_comp.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['AltMSL_lider'], name="Alt Líder", line=dict(color='#ea580c', width=2)), row=1, col=1)
                        if 'AltMSL_ala' in df_merge.columns: fig_comp.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['AltMSL_ala'], name="Alt Ala", line=dict(color='#2563eb', width=2, dash='dot')), row=1, col=1)
                        if 'NormAc_lider' in df_merge.columns: fig_comp.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['NormAc_lider'], name="G Líder", line=dict(color='#ea580c', width=2)), row=2, col=1)
                        if 'NormAc_ala' in df_merge.columns: fig_comp.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['NormAc_ala'], name="G Ala", line=dict(color='#2563eb', width=2, dash='dot')), row=2, col=1)
                        if 'E1 Torq ft lbs_lider' in df_merge.columns: fig_comp.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['E1 Torq ft lbs_lider'], name="Tq Líder", line=dict(color='#ea580c', width=2)), row=3, col=1)
                        if 'E1 Torq ft lbs_ala' in df_merge.columns: fig_comp.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['E1 Torq ft lbs_ala'], name="Tq Ala", line=dict(color='#2563eb', width=2, dash='dot')), row=3, col=1)
                        if 'E1 ITT_lider' in df_merge.columns: fig_comp.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['E1 ITT_lider'], name="T5 Líder", line=dict(color='#ea580c', width=2)), row=4, col=1)
                        if 'E1 ITT_ala' in df_merge.columns: fig_comp.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['E1 ITT_ala'], name="T5 Ala", line=dict(color='#2563eb', width=2, dash='dot')), row=4, col=1)
                        if 'E1 NG_lider' in df_merge.columns: fig_comp.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['E1 NG_lider'], name="Ng Líder", line=dict(color='#ea580c', width=2)), row=5, col=1)
                        if 'E1 NG_ala' in df_merge.columns: fig_comp.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['E1 NG_ala'], name="Ng Ala", line=dict(color='#2563eb', width=2, dash='dot')), row=5, col=1)
                        if 'E1 NP %_lider' in df_merge.columns: fig_comp.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['E1 NP %_lider'], name="Nh Líder", line=dict(color='#ea580c', width=2)), row=6, col=1)
                        if 'E1 NP %_ala' in df_merge.columns: fig_comp.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['E1 NP %_ala'], name="Nh Ala", line=dict(color='#2563eb', width=2, dash='dot')), row=6, col=1)
                        fig_comp.update_layout(template=T['plotly_template'], paper_bgcolor=T['plotly_bg'], plot_bgcolor=T['plotly_bg'], height=1150, hovermode="x unified", margin=dict(l=15, r=15, t=35, b=15))
                        st.plotly_chart(fig_comp, use_container_width=True)

                        st.divider()
                        st.markdown("**Observações - Sincronismo de Potência e Motores**")
                        obs_comp_tmp = st.text_area("Escreva aqui:", key='k_comp')
                        if st.button("SALVAR COMENTÁRIO", key="b_comp"):
                            st.session_state.obs_comparativo = obs_comp_tmp
                            st.session_state["_sync_form_ia"] = True
                            st.rerun()

                    with aba_3d:
                        fig_3d = go.Figure()
                        if 'Longitude_lider' in df_merge.columns and 'Latitude_lider' in df_merge.columns:
                            fig_3d.add_trace(go.Scatter3d(x=df_merge['Longitude_lider'], y=df_merge['Latitude_lider'], z=df_merge['AltMSL_lider'], mode='lines', line=dict(color='#ea580c', width=4), name="Líder"))
                        if 'Longitude_ala' in df_merge.columns and 'Latitude_ala' in df_merge.columns:
                            fig_3d.add_trace(go.Scatter3d(x=df_merge['Longitude_ala'], y=df_merge['Latitude_ala'], z=df_merge['AltMSL_ala'], mode='lines', line=dict(color='#2563eb', width=4), name="Ala"))
                        fig_3d.update_layout(template=T['plotly_template'], paper_bgcolor=T['plotly_bg'], height=700, margin=dict(l=0, r=0, b=0, t=20), scene=dict(aspectmode='auto'))
                        st.plotly_chart(fig_3d, use_container_width=True)
                        st.divider()
                        st.markdown("**Observações - Trajetória Espacial 3D**")
                        obs_3d_tmp = st.text_area("Escreva aqui:", value=st.session_state.obs_3d, key='k_3d')
                        if st.button("SALVAR COMENTÁRIO", key="b_3d"):
                            st.session_state.obs_3d = obs_3d_tmp
                            st.success("Salvo com sucesso.")

                    with aba_rel_form:
                        st.subheader("Emissão de Dossiê Tático de Formatura")
                        c1, c2 = st.columns(2)
                        with c1: lider_editado = st.text_input("Designação Líder:", arq_lider)
                        with c2: ala_editado = st.text_input("Designação Ala:", arq_ala)

                        st.divider()
                        st.markdown("**Análise Técnica Detalhada de Formatura por IA**")
                        if st.button("GERAR ANÁLISE COMPLETA DE FORMATURA COM IA", key="btn_ia_form", use_container_width=True):
                            def _sf(col, fn="mean", fmt="{:.1f}"):
                                if col not in df_merge.columns:
                                    return "N/D"
                                s = df_merge[col].dropna()
                                if s.empty:
                                    return "N/D"
                                if fn == "max": val = s.max()
                                elif fn == "min": val = s.min()
                                elif fn == "mean": val = s.mean()
                                elif fn == "median": val = s.median()
                                elif fn == "p90": val = s.quantile(0.90)
                                elif fn == "p95": val = s.quantile(0.95)
                                elif fn == "std": val = s.std()
                                else: val = s.mean()
                                return fmt.format(val)

                            dur_sync = formatar_duracao(df_merge['Datetime'].iloc[-1] - df_merge['Datetime'].iloc[0]) if len(df_merge) > 1 else "N/D"

                            if 'Distancia_m' in df_merge.columns and not df_merge['Distancia_m'].dropna().empty:
                                d_s = df_merge['Distancia_m'].dropna()
                                pct_cerrada = (d_s <= 50).mean() * 100
                                pct_media = ((d_s > 50) & (d_s <= 150)).mean() * 100
                                pct_afastada = (d_s > 150).mean() * 100
                            else:
                                pct_cerrada = pct_media = pct_afastada = 0.0

                            diff_alt_med, diff_alt_max = "N/D", "N/D"
                            if 'AltMSL_lider' in df_merge.columns and 'AltMSL_ala' in df_merge.columns:
                                diff_alt = (df_merge['AltMSL_ala'] - df_merge['AltMSL_lider']).abs().dropna()
                                if not diff_alt.empty:
                                    diff_alt_med = f"{diff_alt.mean():.0f} ft"
                                    diff_alt_max = f"{diff_alt.max():.0f} ft"

                            with st.spinner("Processando geometria tática, dinâmica comparada e carga de motores..."):
                                try:
                                    prompt_form = f"""
Produza uma ANÁLISE TÉCNICA COMPLETA E DETALHADA da missão em formatura ({projeto_atual}), comparando exclusivamente os sensores monitorados pelo painel entre o LÍDER e o ALA.

=== 1. IDENTIFICAÇÃO E GEOMETRIA TÁTICA ===
- Aeronave Líder (Nº 1): {arq_lider}
- Aeronave Ala (Nº 2): {arq_ala}
- Tempo Total Sincronizado: {dur_sync}
- Separação Horizontal (m): Mínima {_sf('Distancia_m', 'min', '{:.1f}')} m | Média {_sf('Distancia_m', 'mean', '{:.1f}')} m | Mediana (P50) {_sf('Distancia_m', 'median', '{:.1f}')} m | P90 {_sf('Distancia_m', 'p90', '{:.1f}')} m | Máxima {_sf('Distancia_m', 'max', '{:.1f}')} m
- Distribuição de Proximidade: {pct_cerrada:.1f}% do tempo <= 50 m | {pct_media:.1f}% entre 50 m e 150 m | {pct_afastada:.1f}% > 150 m
- Diferencial Vertical Absoluto (|Alt Ala - Alt Líder|): Médio {diff_alt_med} | Máximo {diff_alt_max}

=== 2. DINÂMICA DE VOO E MANOBRAS (LÍDER VS ALA) ===
- Altitude MSL (ft): Líder Média {_sf('AltMSL_lider', 'mean', '{:.0f}')} ft / Máx {_sf('AltMSL_lider', 'max', '{:.0f}')} ft | Ala Média {_sf('AltMSL_ala', 'mean', '{:.0f}')} ft / Máx {_sf('AltMSL_ala', 'max', '{:.0f}')} ft
- Velocidade IAS (kt): Líder Média {_sf('IAS_lider', 'mean', '{:.1f}')} kt / Máx {_sf('IAS_lider', 'max', '{:.1f}')} kt | Ala Média {_sf('IAS_ala', 'mean', '{:.1f}')} kt / Máx {_sf('IAS_ala', 'max', '{:.1f}')} kt
- Força G (NormAc): Líder [{_sf('NormAc_lider', 'min', '{:.2f}')}G a +{_sf('NormAc_lider', 'max', '{:.2f}')}G] | Ala [{_sf('NormAc_ala', 'min', '{:.2f}')}G a +{_sf('NormAc_ala', 'max', '{:.2f}')}G]
- Pitch (º): Líder [{_sf('Pitch_lider', 'min', '{:.1f}')}º a {_sf('Pitch_lider', 'max', '{:.1f}')}º] | Ala [{_sf('Pitch_ala', 'min', '{:.1f}')}º a {_sf('Pitch_ala', 'max', '{:.1f}')}º]
- Roll (º): Líder [{_sf('Roll_lider', 'min', '{:.1f}')}º a {_sf('Roll_lider', 'max', '{:.1f}')}º] | Ala [{_sf('Roll_ala', 'min', '{:.1f}')}º a {_sf('Roll_ala', 'max', '{:.1f}')}º]

=== 3. PERFORMANCE E CARGA DE TRABALHO DOS MOTORES (LÍDER VS ALA) ===
- Torque (lb.ft):
  * Líder: Média {_sf('E1 Torq ft lbs_lider', 'mean', '{:.0f}')} lb.ft | P95 {_sf('E1 Torq ft lbs_lider', 'p95', '{:.0f}')} lb.ft | Máx {_sf('E1 Torq ft lbs_lider', 'max', '{:.0f}')} lb.ft | Oscilação (Std): {_sf('E1 Torq ft lbs_lider', 'std', '{:.0f}')} lb.ft
  * Ala:   Média {_sf('E1 Torq ft lbs_ala', 'mean', '{:.0f}')} lb.ft | P95 {_sf('E1 Torq ft lbs_ala', 'p95', '{:.0f}')} lb.ft | Máx {_sf('E1 Torq ft lbs_ala', 'max', '{:.0f}')} lb.ft | Oscilação (Std): {_sf('E1 Torq ft lbs_ala', 'std', '{:.0f}')} lb.ft
- Temperatura T5 / ITT (ºC):
  * Líder: Média {_sf('E1 ITT_lider', 'mean', '{:.0f}')} ºC | P95 {_sf('E1 ITT_lider', 'p95', '{:.0f}')} ºC | Máx {_sf('E1 ITT_lider', 'max', '{:.0f}')} ºC | Oscilação (Std): {_sf('E1 ITT_lider', 'std', '{:.1f}')} ºC
  * Ala:   Média {_sf('E1 ITT_ala', 'mean', '{:.0f}')} ºC | P95 {_sf('E1 ITT_ala', 'p95', '{:.0f}')} ºC | Máx {_sf('E1 ITT_ala', 'max', '{:.0f}')} ºC | Oscilação (Std): {_sf('E1 ITT_ala', 'std', '{:.1f}')} ºC
- Gerador de Gás Ng (%):
  * Líder: Média {_sf('E1 NG_lider', 'mean', '{:.1f}')}% | P95 {_sf('E1 NG_lider', 'p95', '{:.1f}')}% | Máx {_sf('E1 NG_lider', 'max', '{:.1f}')}% | Oscilação (Std): {_sf('E1 NG_lider', 'std', '{:.2f}')}%
  * Ala:   Média {_sf('E1 NG_ala', 'mean', '{:.1f}')}% | P95 {_sf('E1 NG_ala', 'p95', '{:.1f}')}% | Máx {_sf('E1 NG_ala', 'max', '{:.1f}')}% | Oscilação (Std): {_sf('E1 NG_ala', 'std', '{:.2f}')}%
- Rotação Hélice Nh (%):
  * Líder: Média {_sf('E1 NP %_lider', 'mean', '{:.1f}')}% | Máx {_sf('E1 NP %_lider', 'max', '{:.1f}')}%
  * Ala:   Média {_sf('E1 NP %_ala', 'mean', '{:.1f}')}% | Máx {_sf('E1 NP %_ala', 'max', '{:.1f}')}%

=== FORMATO OBRIGATÓRIO ===
- NÃO mencione kW, HP, kg/s, toneladas, clima ou resolução de panes.
- Use exatamente `#`, `##`, `*` e **negrito** para sensores, deltas e valores numéricos.

# ANÁLISE TÉCNICA DETALHADA DE VOO EM FORMATURA
## 1. GEOMETRIA TÁTICA, SEPARAÇÃO E ESTABILIDADE DE POSIÇÃO
## 2. DINÂMICA DE VOO E SINCRONISMO DE MANOBRAS (LÍDER VS ALA)
## 3. PERFORMANCE DOS MOTORES E MODULAÇÃO DE POTÊNCIA (TORQUE, T5/ITT, NG E NH)
## 4. SÍNTESE COMPARATIVA DE DESEMPENHO DA ESQUADRILHA
"""
                                    texto_ia_form = gerar_texto_ia(prompt_form)
                                    st.session_state.obs_comparativo = texto_ia_form
                                    st.session_state["_sync_form_ia"] = True
                                    mod_usado = st.session_state.get("ultimo_modelo_ia", "IA")
                                    st.session_state.status_ia_form = ("ok", f"Análise de formatura gerada com sucesso (modelo: {mod_usado}).")
                                    st.rerun()
                                except Exception as e:
                                    dist_med = df_merge['Distancia_m'].mean() if 'Distancia_m' in df_merge.columns else 0
                                    dist_max = df_merge['Distancia_m'].max() if 'Distancia_m' in df_merge.columns else 0
                                    st.session_state.obs_comparativo = (
                                        f"# ANÁLISE TÉCNICA DETALHADA DE VOO EM FORMATURA\n"
                                        f"## 1. GEOMETRIA TÁTICA E SEPARAÇÃO\n"
                                        f"* **Esquadrilha:** Líder **{arq_lider}** e Ala **{arq_ala}** (tempo sincronizado: **{dur_sync}**).\n"
                                        f"* **Separação Horizontal:** Média de **{dist_med:.1f} m** e afastamento máximo de **{dist_max:.1f} m**.\n"
                                        f"* **Diferencial Vertical:** Médio de **{diff_alt_med}** e máximo de **{diff_alt_max}**.\n"
                                        f"## 2. DESEMPENHO DOS MOTORES\n"
                                        f"* **Torque Médio (Líder vs Ala):** **{_sf('E1 Torq ft lbs_lider', 'mean', '{:.0f}')} lb.ft** vs **{_sf('E1 Torq ft lbs_ala', 'mean', '{:.0f}')} lb.ft**.\n"
                                        f"* **T5/ITT Média (Líder vs Ala):** **{_sf('E1 ITT_lider', 'mean', '{:.0f}')} ºC** vs **{_sf('E1 ITT_ala', 'mean', '{:.0f}')} ºC**."
                                    )
                                    st.session_state["_sync_form_ia"] = True
                                    st.session_state.status_ia_form = ("warn", f"API de IA indisponível ({e}). Foi gerado o resumo automático local.")
                                    st.rerun()

                        if st.session_state.status_ia_form:
                            _tp, _tx = st.session_state.status_ia_form
                            if _tp == "ok":
                                st.success(_tx)
                            else:
                                st.warning(_tx)

                        obs_form_edit = st.text_area(
                            "Análise de Formatura (Editável para o Dossiê):",
                            key="k_form_rel_ia",
                            height=280
                        )
                        st.session_state.obs_comparativo = obs_form_edit
                        if st.button("CONFIRMAR TEXTO DO DOSSIÊ", key="b_form_rel_conf"):
                            st.session_state.obs_comparativo = obs_form_edit
                            st.session_state["_sync_form_ia"] = True
                            st.success("Texto salvo para exportação.")

                        obs_formatura_atual = {
                            "Posicionamento e Distância": st.session_state.obs_distancia,
                            "Análise Técnica Comparativa (Voo e Motores)": st.session_state.obs_comparativo,
                            "Trajetória Espacial 3D": st.session_state.obs_3d
                        }

                        graficos_form = {
                            "Análise de Distância Horizontal e Vertical": fig_dist,
                            "Comparativo Dinâmico de Motores": fig_comp
                        }

                        st.divider()
                        docx_bytes_form_edit = gerar_docx_formatura_editado(lider_editado, ala_editado, obs_formatura_atual, graficos_form)
                        st.download_button(label="EXPORTAR DOSSIÊ DE FORMATURA (.DOCX)", data=docx_bytes_form_edit, file_name="Dossie_Formatura.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

    # ---------------------------------------------------------
    # FLUXO 2 E 3: VOO SOLO & COMPARATIVO DE MÁQUINA (TREND MONITORING)
    # ---------------------------------------------------------
    else:
        if len(arquivos_up) > 2:
            st.warning("Carregue apenas 1 arquivo para Voo Solo, ou exatamente 2 arquivos para Trend Monitoring.")

        # --- FLUXO COMPARATIVO DE MÁQUINA ---
        elif len(arquivos_up) == 2:
            st.markdown("### Auditoria de Trend Monitoring (Comparativo de Performance)")

            nome_arqA = arquivos_up[0].name
            nome_arqB = arquivos_up[1].name

            colA, colB = st.columns(2)
            with colA: arq_A = st.selectbox("Voo A (Baseline / Referência):", [nome_arqA, nome_arqB], index=0)
            with colB:
                idx_padrao_b = 1 if arq_A == nome_arqA else 0
                arq_B = st.selectbox("Voo B (Recente / Inspecionado):", [nome_arqA, nome_arqB], index=idx_padrao_b)

            if arq_A == arq_B:
                st.error("Você selecionou o mesmo arquivo para Voo A e Voo B.")
            else:
                idx_A = 0 if arq_A == nome_arqA else 1
                idx_B = 1 if idx_A == 0 else 0

                dfA = processar_dados(arquivos_up[idx_A].getvalue())
                dfB = processar_dados(arquivos_up[idx_B].getvalue())

                dfA['Minutos'] = calcular_minutos_voo(dfA)
                dfB['Minutos'] = calcular_minutos_voo(dfB)

                aba_resumo_comp, aba_graf_comp, aba_rel_comp = st.tabs([
                    "1. MATRIZ DE VARIAÇÃO (DELTAS)",
                    "2. CURVAS SOBREPOSTAS",
                    "3. RELATÓRIO DE TENDÊNCIA"
                ])

                df_metricas = pd.DataFrame(columns=['Parâmetro', 'Voo A', 'Voo B', 'Variação (Delta)'])
                def add_metric(param_name, col_name, is_min=False, e_forca_g=False):
                    if col_name in dfA.columns and col_name in dfB.columns:
                        valA = dfA[col_name].min() if is_min else dfA[col_name].max()
                        valB = dfB[col_name].min() if is_min else dfB[col_name].max()
                        diff = valB - valA
                        sinal = "+" if diff > 0 else ""
                        fmt = "{:.2f}" if e_forca_g else "{:.0f}"
                        if param_name.startswith("Ng") or param_name.startswith("Nh"): fmt = "{:.1f}"
                        df_metricas.loc[len(df_metricas)] = [param_name, fmt.format(valA), fmt.format(valB), f"{sinal}{fmt.format(diff)}"]

                add_metric("Força G Máxima (G)", "NormAc", False, True)
                add_metric("Força G Mínima (G)", "NormAc", True, True)
                for i in range(1, qtd_motores_config + 1):
                    add_metric(f"Torque Máx Motor {i} (lb.ft)", f"E{i} Torq ft lbs")
                    add_metric(f"Temp T5/ITT Máx Motor {i} (ºC)", f"E{i} ITT")
                    add_metric(f"Ng Máx Motor {i} (%)", f"E{i} NG")
                    add_metric(f"Nh Máx Motor {i} (%)", f"E{i} NP %")
                    add_metric(f"Temp Óleo Máx Motor {i} (ºC)", f"E{i} OilT_Celsius")

                fig_trend = None

                with aba_resumo_comp:
                    st.markdown("Variações positivas acentuadas (+) em **T5 (ITT)** ou **Ng** para uma mesma faixa de Torque indicam perda de eficiência compressora ou degradação na Seção Quente (*Hot Section*).")
                    render_html_table(df_metricas, height=360)

                    st.divider()
                    st.markdown("**Parecer Técnico - Análise de Deltas**")
                    obs_resumo_tmp = st.text_area("Escreva aqui:", value=st.session_state.obs_resumo, key='k_comp_resumo')
                    if st.button("SALVAR COMENTÁRIO", key="b_comp_res"):
                        st.session_state.obs_resumo = obs_resumo_tmp
                        st.success("Salvo.")

                with aba_graf_comp:
                    fig_trend = make_subplots(rows=4, cols=1, shared_xaxes=True, vertical_spacing=0.04, subplot_titles=("1. Perfil de Altitude MSL (ft)", "2. Torque do Motor 1 (lb.ft)", "3. Temperatura T5 / ITT Motor 1 (ºC)", "4. Rotação Ng Motor 1 (%)"))

                    corA = '#ea580c'
                    corB = '#0284c7'

                    if 'AltMSL' in dfA.columns: fig_trend.add_trace(go.Scattergl(x=dfA['Minutos'], y=dfA['AltMSL'], name="Alt Voo A", line=dict(color=corA, width=2, dash='dot')), row=1, col=1)
                    if 'AltMSL' in dfB.columns: fig_trend.add_trace(go.Scattergl(x=dfB['Minutos'], y=dfB['AltMSL'], name="Alt Voo B", line=dict(color=corB, width=2)), row=1, col=1)
                    if 'E1 Torq ft lbs' in dfA.columns: fig_trend.add_trace(go.Scattergl(x=dfA['Minutos'], y=dfA['E1 Torq ft lbs'], name="Tq Voo A", line=dict(color=corA, width=2)), row=2, col=1)
                    if 'E1 Torq ft lbs' in dfB.columns: fig_trend.add_trace(go.Scattergl(x=dfB['Minutos'], y=dfB['E1 Torq ft lbs'], name="Tq Voo B", line=dict(color=corB, width=2)), row=2, col=1)
                    if 'E1 ITT' in dfA.columns: fig_trend.add_trace(go.Scattergl(x=dfA['Minutos'], y=dfA['E1 ITT'], name="T5 Voo A", line=dict(color=corA, width=2)), row=3, col=1)
                    if 'E1 ITT' in dfB.columns: fig_trend.add_trace(go.Scattergl(x=dfB['Minutos'], y=dfB['E1 ITT'], name="T5 Voo B", line=dict(color=corB, width=2)), row=3, col=1)
                    if 'E1 NG' in dfA.columns: fig_trend.add_trace(go.Scattergl(x=dfA['Minutos'], y=dfA['E1 NG'], name="Ng Voo A", line=dict(color=corA, width=2)), row=4, col=1)
                    if 'E1 NG' in dfB.columns: fig_trend.add_trace(go.Scattergl(x=dfB['Minutos'], y=dfB['E1 NG'], name="Ng Voo B", line=dict(color=corB, width=2)), row=4, col=1)

                    fig_trend.update_layout(template=T['plotly_template'], paper_bgcolor=T['plotly_bg'], plot_bgcolor=T['plotly_bg'], height=900, hovermode="x unified", xaxis4_title="Tempo de Voo Sincronizado (Minutos)", margin=dict(l=15, r=15, t=35, b=15))
                    st.plotly_chart(fig_trend, use_container_width=True)

                    st.divider()
                    st.markdown("**Parecer Técnico - Comportamento Dinâmico**")
                    obs_graficos_tmp = st.text_area("Escreva aqui:", key='k_comp_graficos')
                    if st.button("SALVAR COMENTÁRIO", key="b_comp_graf"):
                        st.session_state.obs_graficos = obs_graficos_tmp
                        st.session_state["_sync_trend_ia"] = True
                        st.rerun()

                with aba_rel_comp:
                    st.markdown("#### Estudo Detalhado de Performance e Degradação de Motor (Trend Monitoring)")

                    if st.button("GERAR ANÁLISE PROFUNDA DE PERFORMANCE E DEGRADAÇÃO COM IA", key="btn_ia_trend", use_container_width=True):
                        dfA_v = dfA[dfA['IAS'] > 50] if 'IAS' in dfA.columns and not dfA[dfA['IAS'] > 50].empty else dfA
                        dfB_v = dfB[dfB['IAS'] > 50] if 'IAS' in dfB.columns and not dfB[dfB['IAS'] > 50].empty else dfB
                        dfA_m = dfA[dfA['E1 NG'] > 50] if 'E1 NG' in dfA.columns and not dfA[dfA['E1 NG'] > 50].empty else dfA_v
                        dfB_m = dfB[dfB['E1 NG'] > 50] if 'E1 NG' in dfB.columns and not dfB[dfB['E1 NG'] > 50].empty else dfB_v

                        def _m(df_x, col, fn="mean"):
                            if col not in df_x.columns:
                                return np.nan
                            s = df_x[col].dropna()
                            if s.empty:
                                return np.nan
                            if fn == "mean": return float(s.mean())
                            if fn == "max": return float(s.max())
                            if fn == "min": return float(s.min())
                            if fn == "p95": return float(s.quantile(0.95))
                            if fn == "std": return float(s.std())
                            return float(s.mean())

                        def _cmp(label, col, fn="mean", fmt="{:.1f}", unit="", df1_ref=dfA_v, df2_ref=dfB_v):
                            vA = _m(df1_ref, col, fn)
                            vB = _m(df2_ref, col, fn)
                            if np.isnan(vA) or np.isnan(vB):
                                return f"- {label}: N/D"
                            d = vB - vA
                            sinal = "+" if d > 0 else ""
                            pct = f" ({sinal}{(d / vA) * 100:.2f}%)" if vA != 0 else ""
                            return f"- {label}: Voo A = {fmt.format(vA)}{unit} | Voo B = {fmt.format(vB)}{unit} | Delta (B-A) = {sinal}{fmt.format(d)}{unit}{pct}"

                        bloco_isotorque = "- Dados de Torque insuficientes para comparação em banda de Iso-Torque."
                        if 'E1 Torq ft lbs' in dfA_v.columns and 'E1 Torq ft lbs' in dfB_v.columns:
                            tq_a_s = dfA_v['E1 Torq ft lbs'].dropna()
                            tq_b_s = dfB_v['E1 Torq ft lbs'].dropna()
                            if not tq_a_s.empty and not tq_b_s.empty:
                                low_tq = float(tq_a_s.quantile(0.60))
                                high_tq = float(tq_a_s.quantile(0.92))
                                isoA = dfA_v[(dfA_v['E1 Torq ft lbs'] >= low_tq) & (dfA_v['E1 Torq ft lbs'] <= high_tq)]
                                isoB = dfB_v[(dfB_v['E1 Torq ft lbs'] >= low_tq) & (dfB_v['E1 Torq ft lbs'] <= high_tq)]

                                if len(isoA) >= 10 and len(isoB) >= 10:
                                    tqA_iso, tqB_iso = _m(isoA, 'E1 Torq ft lbs', 'mean'), _m(isoB, 'E1 Torq ft lbs', 'mean')
                                    ittA_iso, ittB_iso = _m(isoA, 'E1 ITT', 'mean'), _m(isoB, 'E1 ITT', 'mean')
                                    ngA_iso, ngB_iso = _m(isoA, 'E1 NG', 'mean'), _m(isoB, 'E1 NG', 'mean')

                                    raz_itt_A = (ittA_iso / tqA_iso) * 1000 if (not np.isnan(ittA_iso) and not np.isnan(tqA_iso) and tqA_iso > 0) else np.nan
                                    raz_itt_B = (ittB_iso / tqB_iso) * 1000 if (not np.isnan(ittB_iso) and not np.isnan(tqB_iso) and tqB_iso > 0) else np.nan
                                    raz_ng_A = (ngA_iso / tqA_iso) * 1000 if (not np.isnan(ngA_iso) and not np.isnan(tqA_iso) and tqA_iso > 0) else np.nan
                                    raz_ng_B = (ngB_iso / tqB_iso) * 1000 if (not np.isnan(ngB_iso) and not np.isnan(tqB_iso) and tqB_iso > 0) else np.nan

                                    linha_raz_itt = (
                                        f"- Índice Térmico Específico (ºC de ITT por 1000 lb.ft): Voo A = {raz_itt_A:.1f} | Voo B = {raz_itt_B:.1f} | Delta = {(raz_itt_B - raz_itt_A):+.1f}"
                                        if not (np.isnan(raz_itt_A) or np.isnan(raz_itt_B)) else "- Índice Térmico Específico: N/D"
                                    )
                                    linha_raz_ng = (
                                        f"- Índice de Demanda do Compressor (% de Ng por 1000 lb.ft): Voo A = {raz_ng_A:.2f} | Voo B = {raz_ng_B:.2f} | Delta = {(raz_ng_B - raz_ng_A):+.2f}"
                                        if not (np.isnan(raz_ng_A) or np.isnan(raz_ng_B)) else "- Índice de Demanda do Compressor: N/D"
                                    )

                                    bloco_isotorque = f"""- Banda de Referência de Iso-Torque (Regime de Alta Potência): {low_tq:.0f} a {high_tq:.0f} lb.ft (Amostras: Voo A = {len(isoA)} | Voo B = {len(isoB)})
{_cmp('Torque Médio dentro da Banda', 'E1 Torq ft lbs', 'mean', '{:.0f}', ' lb.ft', isoA, isoB)}
{_cmp('T5 / ITT Média para o Mesmo Torque', 'E1 ITT', 'mean', '{:.1f}', ' ºC', isoA, isoB)}
{_cmp('Ng Média para o Mesmo Torque', 'E1 NG', 'mean', '{:.2f}', '%', isoA, isoB)}
{_cmp('Altitude Média durante a Banda de Iso-Torque', 'AltMSL', 'mean', '{:.0f}', ' ft', isoA, isoB)}
{_cmp('Velocidade Média (IAS) durante a Banda', 'IAS', 'mean', '{:.1f}', ' kt', isoA, isoB)}
{linha_raz_itt}
{linha_raz_ng}"""

                        itt_max_A, itt_max_B = _m(dfA, 'E1 ITT', 'max'), _m(dfB, 'E1 ITT', 'max')
                        ng_max_A, ng_max_B = _m(dfA, 'E1 NG', 'max'), _m(dfB, 'E1 NG', 'max')

                        lim_itt = l_itt_red if l_itt_red is not None else 790.0
                        lim_ng = l_ng_red if l_ng_red is not None else 101.5
                        margem_itt_A = f"{lim_itt - itt_max_A:.0f} ºC" if not np.isnan(itt_max_A) else "N/D"
                        margem_itt_B = f"{lim_itt - itt_max_B:.0f} ºC" if not np.isnan(itt_max_B) else "N/D"
                        margem_ng_A = f"{lim_ng - ng_max_A:.1f}%" if not np.isnan(ng_max_A) else "N/D"
                        margem_ng_B = f"{lim_ng - ng_max_B:.1f}%" if not np.isnan(ng_max_B) else "N/D"

                        durA = formatar_duracao(dfA_v['Datetime'].iloc[-1] - dfA_v['Datetime'].iloc[0]) if len(dfA_v) > 1 else "N/D"
                        durB = formatar_duracao(dfB_v['Datetime'].iloc[-1] - dfB_v['Datetime'].iloc[0]) if len(dfB_v) > 1 else "N/D"

                        with st.spinner("Executando análise termodinâmica e de tendência de degradação (Voo A vs Voo B)..."):
                            try:
                                prompt_trend = f"""
Redija um ESTUDO TÉCNICO EXAUSTIVO comparando o Voo A (Baseline / Referência) com o Voo B (Voo Recente / Avaliado) da MESMA aeronave ({projeto_atual}), usando EXCLUSIVAMENTE os sensores monitorados pelo painel para quantificar desvios de performance, deriva térmica da Seção Quente (T5/ITT) e alteração de demanda do gerador de gás (Ng).

=== 1. CONTEXTO E EQUIVALÊNCIA DOS PERFIS DE VOO ===
- Aeronave / Perfil: {projeto_atual}
- Voo A (Referência): {arq_A} | Duração em voo: {durA}
- Voo B (Avaliado): {arq_B} | Duração em voo: {durB}
{_cmp('Altitude Média em Voo (AltMSL)', 'AltMSL', 'mean', '{:.0f}', ' ft')}
{_cmp('Altitude Máxima (Teto)', 'AltMSL', 'max', '{:.0f}', ' ft')}
{_cmp('Velocidade Média (IAS)', 'IAS', 'mean', '{:.1f}', ' kt')}
{_cmp('Velocidade Máxima (IAS)', 'IAS', 'max', '{:.1f}', ' kt')}
{_cmp('Força G Máxima (+)', 'NormAc', 'max', '{:.2f}', ' G')}
{_cmp('Força G Mínima (-)', 'NormAc', 'min', '{:.2f}', ' G')}

=== 2. COMPARATIVO EM BANDA DE ISO-TORQUE (MESMA CARGA MECÂNICA) ===
{bloco_isotorque}

=== 3. ESTATÍSTICA COMPLETA DO GRUPO MOTOPROPULSOR (EM VOO) ===
- TORQUE (lb.ft):
{_cmp('Torque Médio', 'E1 Torq ft lbs', 'mean', '{:.0f}', ' lb.ft')}
{_cmp('Torque P95 (Alta Carga Sustentada)', 'E1 Torq ft lbs', 'p95', '{:.0f}', ' lb.ft')}
{_cmp('Torque Pico Máximo', 'E1 Torq ft lbs', 'max', '{:.0f}', ' lb.ft', dfA, dfB)}
{_cmp('Torque Desvio-Padrão (Variabilidade)', 'E1 Torq ft lbs', 'std', '{:.0f}', ' lb.ft')}

- TEMPERATURA INTER-TURBINAS T5 / ITT (ºC):
{_cmp('T5/ITT Média', 'E1 ITT', 'mean', '{:.1f}', ' ºC')}
{_cmp('T5/ITT P95 (Regime Sustentado)', 'E1 ITT', 'p95', '{:.1f}', ' ºC')}
{_cmp('T5/ITT Pico Máximo', 'E1 ITT', 'max', '{:.0f}', ' ºC', dfA, dfB)}
{_cmp('T5/ITT Oscilação (Desvio-Padrão)', 'E1 ITT', 'std', '{:.1f}', ' ºC')}
- Margem Térmica até Red Line ({lim_itt:.0f} ºC): Voo A = {margem_itt_A} | Voo B = {margem_itt_B}

- GERADOR DE GÁS Ng (%):
{_cmp('Ng Média', 'E1 NG', 'mean', '{:.2f}', '%')}
{_cmp('Ng P95 (Alta Rotação Sustentada)', 'E1 NG', 'p95', '{:.2f}', '%')}
{_cmp('Ng Pico Máximo', 'E1 NG', 'max', '{:.1f}', '%', dfA, dfB)}
{_cmp('Ng Oscilação (Desvio-Padrão)', 'E1 NG', 'std', '{:.2f}', '%')}
- Margem de Rotação até Red Line ({lim_ng:.1f}%): Voo A = {margem_ng_A} | Voo B = {margem_ng_B}

- HÉLICE Nh (%):
{_cmp('Nh Média', 'E1 NP %', 'mean', '{:.2f}', '%')}
{_cmp('Nh Máxima', 'E1 NP %', 'max', '{:.1f}', '%', dfA, dfB)}
{_cmp('Nh Estabilidade (Desvio-Padrão)', 'E1 NP %', 'std', '{:.2f}', '%')}

- SISTEMA DE LUBRIFICAÇÃO (MOTOR EM REGIME Ng > 50%):
{_cmp('Temp. Óleo Média', 'E1 OilT_Celsius', 'mean', '{:.1f}', ' ºC', dfA_m, dfB_m)}
{_cmp('Temp. Óleo Máxima', 'E1 OilT_Celsius', 'max', '{:.0f}', ' ºC', dfA_m, dfB_m)}
{_cmp('Pressão Óleo Mínima em Regime', 'E1 OilP', 'min', '{:.0f}', ' PSI', dfA_m, dfB_m)}
{_cmp('Pressão Óleo Média em Regime', 'E1 OilP', 'mean', '{:.1f}', ' PSI', dfA_m, dfB_m)}
{_cmp('Pressão Óleo Máxima em Regime', 'E1 OilP', 'max', '{:.0f}', ' PSI', dfA_m, dfB_m)}

=== FORMATO OBRIGATÓRIO ===
- NÃO mencione kW, HP, kg/s, fluxo de combustível, toneladas, clima ou resolução de panes/manutenção corretiva.
- Separe claramente o que decorre de diferença de altitude/perfil de voo do que representa variação real de rendimento do motor.
- Use exatamente `#`, `##`, `*` e **negrito** para todos os parâmetros, deltas e conclusões quantitativas.

# ESTUDO DETALHADO DE PERFORMANCE E DEGRADAÇÃO DE MOTOR (TREND MONITORING)
## 1. EQUIVALÊNCIA DE MISSÃO E IMPACTO DO PERFIL DE VOO (ALTITUDE, VELOCIDADE E CARGAS)
## 2. ANÁLISE DE EFICIÊNCIA EM BANDA DE ISO-TORQUE (DESEMPENHO REAL DO MOTOR)
## 3. COMPORTAMENTO TÉRMICO (T5/ITT), DERIVA DA SEÇÃO QUENTE E MARGEM TÉRMICA
## 4. COMPORTAMENTO DO GERADOR DE GÁS (NG), ENTREGA DE TORQUE E GOVERNANÇA DE HÉLICE (NH)
## 5. REGIME TERMOMECÂNICO DO SISTEMA DE LUBRIFICAÇÃO (OILT E OILP)
## 6. CLASSIFICAÇÃO FINAL DA TENDÊNCIA DE PERFORMANCE (ESTÁVEL / DESVIO LEVE / DEGRADAÇÃO ACENTUADA)
"""
                                texto_ia_trend = gerar_texto_ia(prompt_trend)
                                st.session_state.obs_graficos = texto_ia_trend
                                st.session_state["_sync_trend_ia"] = True
                                mod_usado = st.session_state.get("ultimo_modelo_ia", "IA")
                                st.session_state.status_ia_trend = ("ok", f"Estudo de degradação gerado com sucesso (modelo: {mod_usado}).")
                                st.rerun()
                            except Exception as e:
                                st.session_state.obs_graficos = (
                                    f"# ESTUDO DETALHADO DE PERFORMANCE E DEGRADAÇÃO (RESUMO LOCAL)\n"
                                    f"## 1. COMPARATIVO GERAL\n"
                                    f"* **Voo A (Baseline):** {arq_A} ({durA})\n"
                                    f"* **Voo B (Avaliado):** {arq_B} ({durB})\n"
                                    f"## 2. INDICADORES EM BANDA DE ISO-TORQUE\n"
                                    f"{bloco_isotorque}"
                                )
                                st.session_state["_sync_trend_ia"] = True
                                st.session_state.status_ia_trend = ("warn", f"API de IA indisponível ({e}). Foi gerado o resumo automático local.")
                                st.rerun()

                    if st.session_state.status_ia_trend:
                        _tp, _tx = st.session_state.status_ia_trend
                        if _tp == "ok":
                            st.success(_tx)
                        else:
                            st.warning(_tx)

                    obs_trend_edit = st.text_area(
                        "Estudo de Performance e Degradação (Editável para o Relatório Word):",
                        key="k_trend_rel_ia",
                        height=300
                    )
                    st.session_state.obs_graficos = obs_trend_edit
                    if st.button("CONFIRMAR ESTUDO PARA EXPORTAÇÃO", key="b_trend_conf"):
                        st.session_state.obs_graficos = obs_trend_edit
                        st.session_state["_sync_trend_ia"] = True
                        st.success("Estudo salvo para exportação.")

                    obs_comp_atual = {
                        "Resumo de Degradação e Picos (Deltas)": st.session_state.obs_resumo,
                        "Estudo Detalhado de Performance e Degradação": st.session_state.obs_graficos
                    }

                    graficos_trend = {
                        "Comparativo de Performance (Voo A vs Voo B)": fig_trend
                    }

                    st.divider()
                    docx_bytes_comp = gerar_docx_comparativo_editado(arq_A, arq_B, df_metricas, obs_comp_atual, graficos_trend)
                    st.download_button(
                        label="EXPORTAR LAUDO DE TREND MONITORING (.DOCX)",
                        data=docx_bytes_comp,
                        file_name="Comparativo_Degradacao.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )


        # --- FLUXO VOO SOLO ORIGINAL (COM REGRAS OFICIAIS DO MANUAL T-27M) ---
        else:
            df = processar_dados(arquivos_up[0].getvalue())

            max_g = df['NormAc'].max() if 'NormAc' in df.columns else 0
            min_g = df['NormAc'].min() if 'NormAc' in df.columns else 0
            alertas = []

            def add_alerta_contexto(coluna, valor, mensagem):
                try:
                    row = df[df[coluna] == valor].iloc[0]
                    tempo = row['Datetime'].strftime('%H:%M:%S') if pd.notna(row['Datetime']) else 'N/A'
                    alt_val = row.get('AltMSL', np.nan)
                    vel_val = row.get('IAS', np.nan)
                    alt = f"{alt_val:.0f} ft" if pd.notna(alt_val) else 'N/A'
                    vel = f"{vel_val:.0f} kt" if pd.notna(vel_val) else 'N/A'
                    alertas.append(f"{mensagem} -> Ocorrência às {tempo} | Alt: {alt} | Vel: {vel}")
                except:
                    alertas.append(mensagem)

            if l_g_max is not None and max_g > l_g_max:
                add_alerta_contexto('NormAc', max_g, f"[CRÍTICO] LINHA VERMELHA: Força G Positiva ({max_g:.2f}G | Limite: +{l_g_max:.1f}G)")
            if l_g_min is not None and min_g < l_g_min:
                add_alerta_contexto('NormAc', min_g, f"[CRÍTICO] LINHA VERMELHA: Força G Negativa ({min_g:.2f}G | Limite: {l_g_min:.1f}G)")

            if projeto_atual == "T-27M (Oficial)":
                max_torq = df['E1 Torq ft lbs'].max() if 'E1 Torq ft lbs' in df.columns else 0
                max_itt = df['E1 ITT'].max() if 'E1 ITT' in df.columns else 0
                max_ng = df['E1 NG'].max() if 'E1 NG' in df.columns else 0
                max_nh = df['E1 NP %'].max() if 'E1 NP %' in df.columns else 0

                if 'NormAc' in df.columns:
                    em_invertido = df['NormAc'] < -0.15
                    seg_invertido = int(em_invertido.groupby((~em_invertido).cumsum()).sum().max()) if not em_invertido.empty else 0
                    if seg_invertido > 15:
                        add_alerta_contexto('NormAc', min_g, f"[CRÍTICO] EXCEDÊNCIA VOO INVERTIDO (INSTRUÇÃO): {seg_invertido}s contínuos (Limite Manual: 15s)")

                if verificar_temporizador(df, 'E1 Torq ft lbs', 2100, 2):
                    add_alerta_contexto('E1 Torq ft lbs', max_torq, f"[CRÍTICO] OVER-TORQUE CRÍTICO: >2100 lb.ft por 2s ({max_torq:.0f} lb.ft)")
                elif verificar_temporizador(df, 'E1 Torq ft lbs', 1970, 1):
                    add_alerta_contexto('E1 Torq ft lbs', max_torq, f"[CRÍTICO] OVER-TORQUE: Passou de 1970 lb.ft ({max_torq:.0f} lb.ft)")
                elif l_tq_red is not None and max_torq > l_tq_red:
                    add_alerta_contexto('E1 Torq ft lbs', max_torq, f"[CRÍTICO] LINHA VERMELHA Torque Decolagem (>1795 lb.ft): {max_torq:.0f} lb.ft")

                if verificar_temporizador(df, 'E1 NG', 102.6, 6):
                    add_alerta_contexto('E1 NG', max_ng, f"[CRÍTICO] NG CRÍTICO: >102.6% por 6 seg ({max_ng:.1f}%)")
                elif l_ng_red is not None and max_ng > l_ng_red:
                    add_alerta_contexto('E1 NG', max_ng, f"[CRÍTICO] LINHA VERMELHA Ng (>101.5%): {max_ng:.1f}%")
                elif l_ng_yel is not None and max_ng > l_ng_yel:
                    add_alerta_contexto('E1 NG', max_ng, f"[ALERTA] ARCO AMARELO Ng (>100.0%): {max_ng:.1f}%")

                if verificar_temporizador(df, 'E1 ITT', 1090, 2):
                    add_alerta_contexto('E1 ITT', max_itt, f"[CRÍTICO] T5 CRÍTICA NA PARTIDA: >1090ºC por 2 seg ({max_itt:.0f}ºC)")
                elif verificar_temporizador(df, 'E1 ITT', 850, 2):
                    add_alerta_contexto('E1 ITT', max_itt, f"[CRÍTICO] T5 MÁX EM ACELERAÇÃO EXCEDIDA: >850ºC por 2 seg ({max_itt:.0f}ºC)")
                elif verificar_temporizador(df, 'E1 ITT', 790, 1):
                    add_alerta_contexto('E1 ITT', max_itt, f"[CRÍTICO] LINHA VERMELHA T5 DECOLAGEM (>790ºC): {max_itt:.0f}ºC")
                elif verificar_temporizador(df, 'E1 ITT', 765, 5):
                    add_alerta_contexto('E1 ITT', max_itt, f"[ALERTA] T5 MÁX SUBIDA (>765ºC): {max_itt:.0f}ºC")
                elif l_itt_yel is not None and max_itt > l_itt_yel:
                    add_alerta_contexto('E1 ITT', max_itt, f"[ALERTA] ARCO AMARELO T5 CRUZEIRO (>740ºC): {max_itt:.0f}ºC")

                if verificar_temporizador(df, 'E1 NP %', 110, 1):
                    add_alerta_contexto('E1 NP %', max_nh, f"[CRÍTICO] OVERSPEED HÉLICE CRÍTICO: >110% ({max_nh:.1f}%)")
                elif verificar_temporizador(df, 'E1 NP %', 106, 1):
                    add_alerta_contexto('E1 NP %', max_nh, f"[CRÍTICO] OVERSPEED HÉLICE: >106% ({max_nh:.1f}%)")
                elif l_nh_red is not None and max_nh > l_nh_red:
                    add_alerta_contexto('E1 NP %', max_nh, f"[CRÍTICO] LINHA VERMELHA Nh (>100.0%): {max_nh:.1f}%")

                if 'E1 OilP' in df.columns and 'E1 NG' in df.columns:
                    df_mot_ligado = df[df['E1 NG'] > 50]
                    if not df_mot_ligado.empty and (df_mot_ligado['E1 OilP'] < 40).any():
                        min_oilp_voo = df_mot_ligado['E1 OilP'].min()
                        add_alerta_contexto('E1 OilP', min_oilp_voo, f"[CRÍTICO] LUZ PRESS ÓLEO ATIVADA: Pressão caiu para {min_oilp_voo:.0f} PSI (Limite mín: 40 PSI)")
            else:
                for i_mot in range(1, qtd_motores_config + 1):
                    m_torq = df[f'E{i_mot} Torq ft lbs'].max() if f'E{i_mot} Torq ft lbs' in df.columns else 0
                    m_itt = df[f'E{i_mot} ITT'].max() if f'E{i_mot} ITT' in df.columns else 0
                    m_ng = df[f'E{i_mot} NG'].max() if f'E{i_mot} NG' in df.columns else 0
                    m_nh = df[f'E{i_mot} NP %'].max() if f'E{i_mot} NP %' in df.columns else 0

                    if l_tq_red is not None and m_torq >= l_tq_red: add_alerta_contexto(f'E{i_mot} Torq ft lbs', m_torq, f"[CRÍTICO] Motor {i_mot} LINHA VERMELHA Torque: {m_torq:.0f} lb.ft")
                    elif l_tq_yel is not None and m_torq >= l_tq_yel: add_alerta_contexto(f'E{i_mot} Torq ft lbs', m_torq, f"[ALERTA] Motor {i_mot} ARCO AMARELO Torque: {m_torq:.0f} lb.ft")
                    if l_ng_red is not None and m_ng >= l_ng_red: add_alerta_contexto(f'E{i_mot} NG', m_ng, f"[CRÍTICO] Motor {i_mot} LINHA VERMELHA Ng: {m_ng:.1f}%")
                    elif l_ng_yel is not None and m_ng >= l_ng_yel: add_alerta_contexto(f'E{i_mot} NG', m_ng, f"[ALERTA] Motor {i_mot} ARCO AMARELO Ng: {m_ng:.1f}%")
                    if l_itt_red is not None and m_itt >= l_itt_red: add_alerta_contexto(f'E{i_mot} ITT', m_itt, f"[CRÍTICO] Motor {i_mot} LINHA VERMELHA ITT: {m_itt:.0f} ºC")
                    elif l_itt_yel is not None and m_itt >= l_itt_yel: add_alerta_contexto(f'E{i_mot} ITT', m_itt, f"[ALERTA] Motor {i_mot} ARCO AMARELO ITT: {m_itt:.0f} ºC")
                    if l_nh_red is not None and m_nh >= l_nh_red: add_alerta_contexto(f'E{i_mot} NP %', m_nh, f"[CRÍTICO] Motor {i_mot} LINHA VERMELHA Nh: {m_nh:.1f}%")
                    elif l_nh_yel is not None and m_nh >= l_nh_yel: add_alerta_contexto(f'E{i_mot} NP %', m_nh, f"[ALERTA] Motor {i_mot} ARCO AMARELO Nh: {m_nh:.1f}%")

            for i_mot in range(1, qtd_motores_config + 1):
                m_oilt = df[f'E{i_mot} OilT_Celsius'].max() if f'E{i_mot} OilT_Celsius' in df.columns else 0
                m_oilp = df[f'E{i_mot} OilP'].max() if f'E{i_mot} OilP' in df.columns else 0
                if l_oilt_red is not None and m_oilt >= l_oilt_red: add_alerta_contexto(f'E{i_mot} OilT_Celsius', m_oilt, f"[CRÍTICO] Motor {i_mot} LINHA VERMELHA Temp Óleo: {m_oilt:.0f} ºC")
                if l_oilp_red is not None and m_oilp >= l_oilp_red: add_alerta_contexto(f'E{i_mot} OilP', m_oilp, f"[CRÍTICO] Motor {i_mot} LINHA VERMELHA Pressão Óleo: {m_oilp:.0f} PSI")

            tempos = {'acionamento': '-', 'decolagem': '-', 'pouso': '-', 'corte': '-', 't_voo': '-'}
            if tipo_projeto == "Asa Rotativa":
                motor_on = df[(df.get('E1 NG', pd.Series([0])) > 15) & (df['Datetime'].notna())]
                if not motor_on.empty:
                    tempos['acionamento'] = motor_on.iloc[0]['Datetime'].strftime('%H:%M:%S')
                    tempos['corte'] = motor_on.iloc[-1]['Datetime'].strftime('%H:%M:%S')
                    tempos['decolagem'] = tempos['acionamento']
                    tempos['pouso'] = tempos['corte']
                    tempos['t_voo'] = formatar_duracao(motor_on.iloc[-1]['Datetime'] - motor_on.iloc[0]['Datetime'])
                voo_ativo = motor_on
            else:
                if 'E1 NG' in df.columns:
                    motor_on = df[(df['E1 NG'] > 10) & (df['Datetime'].notna())]
                    if not motor_on.empty:
                        tempos['acionamento'] = motor_on.iloc[0]['Datetime'].strftime('%H:%M:%S')
                        tempos['corte'] = motor_on.iloc[-1]['Datetime'].strftime('%H:%M:%S')
                if 'IAS' in df.columns:
                    voo_ativo = df[(df['IAS'] > 50) & (df['Datetime'].notna())]
                    if not voo_ativo.empty:
                        tempos['decolagem'] = voo_ativo.iloc[0]['Datetime'].strftime('%H:%M:%S')
                        tempos['pouso'] = voo_ativo.iloc[-1]['Datetime'].strftime('%H:%M:%S')
                        tempos['t_voo'] = formatar_duracao(voo_ativo.iloc[-1]['Datetime'] - voo_ativo.iloc[0]['Datetime'])
                else:
                    voo_ativo = pd.DataFrame()

            cor_status_missao = "#ef4444" if any("[CRÍTICO]" in a for a in alertas) else ("#f59e0b" if len(alertas) > 0 else "#10b981")
            txt_status_missao = f"{len(alertas)} EXCEDÊNCIA(S) DETECTADA(S)" if len(alertas) > 0 else "OPERAÇÃO NORMAL (ARCO VERDE)"

            st.markdown(f"""
            <div style="background: {T['bg_card']}; border: 1px solid {T['border']}; border-left: 5px solid {cor_status_missao}; border-radius: 10px; padding: 14px 20px; margin-bottom: 18px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 15px; box-shadow: 0 2px 8px rgba(0,0,0,0.05);">
                <div style="display: flex; gap: 28px; flex-wrap: wrap;">
                    <div>
                        <div style="font-size: 10px; color: {T['txt_sec']}; font-weight: 700; text-transform: uppercase;">Acionamento (Start)</div>
                        <div style="font-size: 15px; font-weight: 700; color: {T['txt_main']}; font-family: 'JetBrains Mono', Consolas, monospace;">{tempos['acionamento']}</div>
                    </div>
                    <div>
                        <div style="font-size: 10px; color: {T['txt_sec']}; font-weight: 700; text-transform: uppercase;">Decolagem (T/O)</div>
                        <div style="font-size: 15px; font-weight: 700; color: {T['accent']}; font-family: 'JetBrains Mono', Consolas, monospace;">{tempos['decolagem']}</div>
                    </div>
                    <div>
                        <div style="font-size: 10px; color: {T['txt_sec']}; font-weight: 700; text-transform: uppercase;">Pouso (LDG)</div>
                        <div style="font-size: 15px; font-weight: 700; color: {T['accent']}; font-family: 'JetBrains Mono', Consolas, monospace;">{tempos['pouso']}</div>
                    </div>
                    <div>
                        <div style="font-size: 10px; color: {T['txt_sec']}; font-weight: 700; text-transform: uppercase;">Corte (Shutdown)</div>
                        <div style="font-size: 15px; font-weight: 700; color: {T['txt_main']}; font-family: 'JetBrains Mono', Consolas, monospace;">{tempos['corte']}</div>
                    </div>
                    <div>
                        <div style="font-size: 10px; color: {T['txt_sec']}; font-weight: 700; text-transform: uppercase;">Tempo de Voo</div>
                        <div style="font-size: 15px; font-weight: 700; color: #10b981; font-family: 'JetBrains Mono', Consolas, monospace;">{tempos['t_voo']}</div>
                    </div>
                </div>
                <div style="background: {cor_status_missao}22; border: 1px solid {cor_status_missao}66; color: {cor_status_missao}; padding: 6px 14px; border-radius: 20px; font-size: 11.5px; font-weight: 700; letter-spacing: 0.6px;">
                    ● {txt_status_missao}
                </div>
            </div>
            """, unsafe_allow_html=True)

            aba_didatica, aba_inspetoria, aba_manobras, aba_relatorio = st.tabs([
                "1. PAINEL DE INSTRUMENTOS (EIS)",
                "2. INSPETORIA E TELEMETRIA",
                "3. ENVELOPE DE MANOBRAS E POUSO",
                "4. LAUDO TÉCNICO E IA"
            ])

            fig_dash = None

            with aba_didatica:
                st.markdown("#### Leituras de Pico da Missão (Engine & Airframe Limits)")
                c1, c2, c3, c4 = st.columns(4)
                with c1:
                    exibir_card_arcos("Força G Máxima (+)", max_g, l_g_max, None, None, "G", ":.2f")
                    max_ng = df['E1 NG'].max() if 'E1 NG' in df.columns else 0
                    exibir_card_arcos("Gerador de Gás (Ng)", max_ng, l_ng_red, l_ng_yel, None, "%", ":.1f")
                with c2:
                    exibir_card_arcos("Força G Mínima (-)", min_g, l_g_min, None, None, "G", ":.2f", e_minimo=True)
                    max_nh = df['E1 NP %'].max() if 'E1 NP %' in df.columns else 0
                    exibir_card_arcos("Rotação Hélice (Nh)", max_nh, l_nh_red, l_nh_yel, None, "%", ":.1f")
                with c3:
                    max_torq = df['E1 Torq ft lbs'].max() if 'E1 Torq ft lbs' in df.columns else 0
                    exibir_card_arcos("Torque Máximo", max_torq, l_tq_red, l_tq_yel, None, "lb.ft", ":.0f")
                    max_oilt = df['E1 OilT_Celsius'].max() if 'E1 OilT_Celsius' in df.columns else 0
                    exibir_card_arcos("Temp. do Óleo (OilT)", max_oilt, l_oilt_red, None, l_oilt_min, "ºC", ":.0f")
                with c4:
                    max_itt = df['E1 ITT'].max() if 'E1 ITT' in df.columns else 0
                    exibir_card_arcos("Temperatura T5 / ITT", max_itt, l_itt_red, l_itt_yel, None, "ºC", ":.0f")
                    max_oilp = df['E1 OilP'].max() if 'E1 OilP' in df.columns else 0
                    exibir_card_arcos("Pressão do Óleo (OilP)", max_oilp, l_oilp_red, None, l_oilp_min, "PSI", ":.0f")

                if projeto_atual == "T-27M (Oficial)":
                    with st.expander("Consultar Gabarito Oficial de Limites Operacionais T-27M"):
                        st.markdown("""
                        | Parâmetro Operacional | Limite Oficial T-27M | Parâmetro Operacional | Limite Oficial T-27M |
                        | :--- | :--- | :--- | :--- |
                        | **Torque Máx para Decolagem** | `1795 lb.ft` | **Vel. Máx Flape Decolagem** | `176 kt` |
                        | **T5 Máx para Decolagem** | `790 °C` | **Vel. Máx Flape Pouso** | `148 kt` |
                        | **T5 Máx de Subida** | `765 °C` | **Vel. Máx Baixamento Trem (Normal / Emerg)** | `145 kt / 110 kt` |
                        | **T5 Máx em Aceleração (Atitude Normal)** | `850 °C em 2s` | **Vel. Máx Recolhimento do Trem** | `120 kt` |
                        | **T5 Máx de Cruzeiro** | `740 °C` | **Vel. Máx com Trem Baixado** | `146 kt` |
                        | **T5 Máx para a Partida** | `1090 °C em 2s` | **Luz Pressão do Óleo** | `Abaixo de 40 psi` |
                        | **Ng Máx (Decolagem / Subida / Cruzeiro)** | `101,5 %` | **Pressão de Oxigênio (Mín / Máx)** | `50 psi / 450 psi` |
                        | **Nh Máx (Decolagem / Subida / Cruzeiro)** | `100 %` | **Corrente DC Gerador (Contínua / 5 min)** | `200 A / 250 A` |
                        | **Voo Invertido (Na Instrução)** | `15 segundos` | **Limites G (Flapes Recolhidos / Baixados)** | `-3G a +6G` / `0 a +2G` |
                        """)

                st.divider()
                st.markdown("#### Navegação e Reconstituição de Rota (GPS / Google Earth)")
                df_mapa = df[(df['IAS'] > 30)].dropna(subset=['Latitude', 'Longitude'])
                if not df_mapa.empty:
                    fig_map_temp = go.Figure()
                    fig_map_temp.add_trace(go.Scattermap(lat=df_mapa['Latitude'], lon=df_mapa['Longitude'], mode='lines', line=dict(width=4, color='#0284c7'), name="Rota"))
                    fig_map_temp.add_trace(go.Scattermap(lat=[df_mapa['Latitude'].iloc[0]], lon=[df_mapa['Longitude'].iloc[0]], mode='markers', marker=dict(size=13, color='#10b981'), name="Decolagem"))
                    fig_map_temp.add_trace(go.Scattermap(lat=[df_mapa['Latitude'].iloc[-1]], lon=[df_mapa['Longitude'].iloc[-1]], mode='markers', marker=dict(size=13, color='#ef4444'), name="Corte"))
                    fig_map_temp.update_layout(map_style=T['map_style'], map_zoom=8, map_center={"lat": df_mapa['Latitude'].mean(), "lon": df_mapa['Longitude'].mean()}, margin={"r":0,"t":0,"l":0,"b":0}, height=420, showlegend=False)
                    st.plotly_chart(fig_map_temp, use_container_width=True)

                    kml = simplekml.Kml()
                    linestring = kml.newlinestring(name="Trajetoria de Voo 3D", coords=[(r['Longitude'], r['Latitude'], r['AltMSL'] * 0.3048) for i, r in df_mapa.iterrows()])
                    linestring.altitudemode = simplekml.AltitudeMode.absolute
                    linestring.style.linestyle.color = simplekml.Color.cyan
                    linestring.style.linestyle.width = 4
                    st.download_button("EXPORTAR TELEMETRIA 3D PARA GOOGLE EARTH (.KML)", data=kml.kml(), file_name="rota_voo.kml", mime="application/vnd.google-earth.kml+xml")

                st.divider()
                st.markdown("**Apontamentos Técnicos - Painel e Navegação**")
                obs_res_tmp = st.text_area("Escreva aqui:", value=st.session_state.obs_resumo, key='k_res')
                if st.button("SALVAR COMENTÁRIO", key="b_res"):
                    st.session_state.obs_resumo = obs_res_tmp
                    st.success("Salvo com sucesso.")

            with aba_inspetoria:
                st.markdown("#### Auditoria Automática de Excedências (Arcos Amarelo e Vermelho)")
                if len(alertas) > 0:
                    for a in alertas:
                        if "[CRÍTICO]" in a or "LINHA VERMELHA" in a:
                            st.error(a)
                        else:
                            st.warning(a)
                else:
                    st.success("Nenhuma violação de limite operacional registrada durante a missão.")

                st.divider()
                st.markdown("#### Filtro de Investigação de Eventos (Data Query)")
                c_busca1, c_busca2, c_busca3 = st.columns([2, 1, 1])
                opcoes_busca = {'Força G': 'NormAc', 'Torque M1': 'E1 Torq ft lbs', 'T5 / ITT M1': 'E1 ITT', 'Ng M1 (%)': 'E1 NG', 'Nh M1 (%)': 'E1 NP %', 'Temp. Óleo M1': 'E1 OilT_Celsius', 'Pressão Óleo M1': 'E1 OilP', 'Velocidade': 'IAS', 'Altitude': 'AltMSL', 'Arfagem (Pitch)': 'Pitch', 'Rolagem (Roll)': 'Roll'}
                if qtd_motores_config > 1: opcoes_busca.update({'Torque M2': 'E2 Torq ft lbs', 'T5 / ITT M2': 'E2 ITT', 'Ng M2 (%)': 'E2 NG', 'Nh M2 (%)': 'E2 NP %', 'Temp. Óleo M2': 'E2 OilT_Celsius', 'Pressão Óleo M2': 'E2 OilP'})

                with c_busca1: escolha = st.selectbox("Sensor / Parâmetro:", list(opcoes_busca.keys()))
                with c_busca2: operador = st.selectbox("Operador Lógico:", ["Maior que >", "Menor que <"])
                with c_busca3: valor = st.number_input("Limiar de Disparo:", value=3.0)

                coluna_filtro = opcoes_busca[escolha]
                if coluna_filtro in df.columns:
                    resultado_busca = df[df[coluna_filtro] > valor] if operador == "Maior que >" else df[df[coluna_filtro] < valor]
                    if not resultado_busca.empty:
                        colunas_ver = ['Lcl Time', 'AltMSL', 'IAS', 'NormAc', 'Pitch', 'Roll']
                        for j in range(1, qtd_motores_config + 1): colunas_ver.extend([f'E{j} Torq ft lbs', f'E{j} ITT', f'E{j} NG', f'E{j} NP %', f'E{j} OilT_Celsius', f'E{j} OilP'])
                        colunas_ver = [c for c in colunas_ver if c in df.columns]
                        exibir_tabela_grande(resultado_busca[colunas_ver].head(20).copy())
                    else:
                        st.info("Nenhum registro encontrado para o filtro selecionado.")

                st.divider()
                st.markdown("#### Gráficos Sincronizados de Dinâmica e Propulsão")
                fig_dash = make_subplots(rows=6, cols=1, shared_xaxes=True, vertical_spacing=0.03, specs=[[{"secondary_y": True}], [{"secondary_y": False}], [{"secondary_y": False}], [{"secondary_y": False}], [{"secondary_y": False}], [{"secondary_y": False}]], subplot_titles=("1. Altitude MSL (ft) e Velocidade IAS (kt)", "2. Fator de Carga Normal (Força G)", "3. Torque do Motor (lb.ft)", "4. Temperatura Inter-Turbinas - T5 / ITT (ºC)", "5. Rotação do Gerador de Gás - Ng (%)", "6. Rotação da Hélice - Nh (%)"))
                cores_motores = ['#10b981', '#f59e0b', '#f97316', '#8b5cf6']
                cores_itt = ['#d97706', '#ef4444', '#ea580c', '#dc2626']
                cores_ng = ['#9333ea', '#0284c7', '#0d9488', '#ea580c']
                cores_nh = ['#ca8a04', '#16a34a', '#475569', '#64748b']

                if 'AltMSL' in df.columns: fig_dash.add_trace(go.Scattergl(x=df['Datetime'], y=df['AltMSL'], name="Altitude (ft)", fill='tozeroy', fillcolor='rgba(37, 99, 235, 0.12)', line=dict(color='#2563eb', width=2)), row=1, col=1, secondary_y=False)
                if 'IAS' in df.columns: fig_dash.add_trace(go.Scattergl(x=df['Datetime'], y=df['IAS'], name="Velocidade (kt)", line=dict(color='#0891b2', width=2)), row=1, col=1, secondary_y=True)
                if 'NormAc' in df.columns: fig_dash.add_trace(go.Scattergl(x=df['Datetime'], y=df['NormAc'], name="Força G", fill='tozeroy', fillcolor='rgba(239, 68, 68, 0.12)', line=dict(color='#ef4444', width=2)), row=2, col=1)
                if l_g_max is not None: fig_dash.add_hline(y=l_g_max, line_dash="dash", line_color="#ef4444", row=2, col=1)

                for i_m in range(1, qtd_motores_config + 1):
                    idx_c = (i_m - 1) % 4
                    if f'E{i_m} Torq ft lbs' in df.columns: fig_dash.add_trace(go.Scattergl(x=df['Datetime'], y=df[f'E{i_m} Torq ft lbs'], name=f"Tq M{i_m}", line=dict(color=cores_motores[idx_c], width=2)), row=3, col=1)
                    if f'E{i_m} ITT' in df.columns: fig_dash.add_trace(go.Scattergl(x=df['Datetime'], y=df[f'E{i_m} ITT'], name=f"T5 M{i_m}", line=dict(color=cores_itt[idx_c], width=2)), row=4, col=1)
                    if f'E{i_m} NG' in df.columns: fig_dash.add_trace(go.Scattergl(x=df['Datetime'], y=df[f'E{i_m} NG'], name=f"Ng M{i_m} (%)", line=dict(color=cores_ng[idx_c], width=2)), row=5, col=1)
                    if f'E{i_m} NP %' in df.columns: fig_dash.add_trace(go.Scattergl(x=df['Datetime'], y=df[f'E{i_m} NP %'], name=f"Nh M{i_m} (%)", line=dict(color=cores_nh[idx_c], width=2)), row=6, col=1)

                if l_tq_red is not None: fig_dash.add_hline(y=l_tq_red, line_dash="dash", line_color="#ef4444", row=3, col=1)
                if l_itt_red is not None: fig_dash.add_hline(y=l_itt_red, line_dash="dash", line_color="#ef4444", row=4, col=1)
                if l_ng_red is not None: fig_dash.add_hline(y=l_ng_red, line_dash="dash", line_color="#ef4444", row=5, col=1)
                if l_nh_red is not None: fig_dash.add_hline(y=l_nh_red, line_dash="dash", line_color="#ef4444", row=6, col=1)

                fig_dash.update_layout(template=T['plotly_template'], paper_bgcolor=T['plotly_bg'], plot_bgcolor=T['plotly_bg'], height=1200, hovermode="x unified", margin=dict(l=15, r=15, t=35, b=15), legend=dict(title="Sensores", orientation="v", yanchor="top", y=1, xanchor="left", x=1.02))
                st.plotly_chart(fig_dash, use_container_width=True)

                st.divider()
                st.markdown("**Observações Manuais de Inspetoria**")
                obs_graf_tmp_manual = st.text_area("Escreva o seu parecer manual aqui:", key='k_graf_manual')
                if st.button("SALVAR COMENTÁRIO MANUAL", key="b_graf_manual"):
                    st.session_state.obs_graficos = obs_graf_tmp_manual
                    st.session_state["_sync_solo_ia"] = True
                    st.rerun()

            with aba_manobras:
                c_m1, c_m2 = st.columns(2)
                with c_m1:
                    st.markdown("#### Diagnóstico de Aproximação e Toque (Landing)")
                    if 'IAS' in df.columns and 'AltMSL' in df.columns and not voo_ativo.empty:
                        fase_final = voo_ativo.tail(300)
                        pico_g_pouso = fase_final['NormAc'].max()
                        vel_aprox = fase_final['IAS'].mean()
                        st.info(f"**Velocidade Média na Final:** `{vel_aprox:.1f} kt` (Limite Flape Pouso T-27M: `148 kt`)")
                        if pico_g_pouso > 2.0: st.error(f"**Pouso Brusco / Excedência Flape Baixado (>2.0G):** Pico de `{pico_g_pouso:.2f} G` registrado.")
                        elif pico_g_pouso > 1.5: st.warning(f"**Pouso Firme:** Pico de `{pico_g_pouso:.2f} G` na pista.")
                        else: st.success(f"**Pouso Suave:** Toque estabilizado em `{pico_g_pouso:.2f} G`.")
                with c_m2:
                    st.markdown("#### Detector de Manobras Acrobáticas / Atitudes Extremas")
                    if 'Pitch' in df.columns and 'Roll' in df.columns:
                        df_acro = df[(df['Pitch'] > 45) | (df['Pitch'] < -45) | (df['Roll'] > 60) | (df['Roll'] < -60)]
                        if not df_acro.empty:
                            st.markdown(f"**Registros de elevada atitude (Pitch > 45º ou Roll > 60º):** `{len(df_acro)}` amostras.")
                            colunas_acro = ['Lcl Time', 'AltMSL', 'IAS', 'Pitch', 'Roll', 'NormAc']
                            for j in range(1, qtd_motores_config + 1): colunas_acro.extend([f'E{j} Torq ft lbs', f'E{j} ITT', f'E{j} NG', f'E{j} NP %'])
                            colunas_acro = [c for c in colunas_acro if c in df.columns]
                            exibir_tabela_grande(df_acro.sort_values(by='NormAc', ascending=False).head(20)[colunas_acro].copy())
                        else:
                            st.info("Voo nivelado / sem manobras de grande inclinação detectadas.")

                st.divider()
                st.markdown("#### Perfil Espacial 3D Colorido por Velocidade (IAS)")
                if not voo_ativo.empty and 'Latitude' in voo_ativo.columns and 'Longitude' in voo_ativo.columns:
                    df_3d = voo_ativo.dropna(subset=['Latitude', 'Longitude', 'AltMSL'])
                    fig_3d_temp = go.Figure(data=[go.Scatter3d(x=df_3d['Longitude'], y=df_3d['Latitude'], z=df_3d['AltMSL'], mode='lines', line=dict(color=df_3d['IAS'], colorscale='Turbo', width=4, colorbar=dict(title="IAS (kt)", len=0.5, x=1.05)))])
                    fig_3d_temp.update_layout(template=T['plotly_template'], paper_bgcolor=T['plotly_bg'], height=600, margin=dict(l=0, r=0, b=0, t=20), scene=dict(aspectmode='auto'))
                    st.plotly_chart(fig_3d_temp, use_container_width=True)

                st.divider()
                st.markdown("**Observações - Envelope de Manobras e Pouso**")
                obs_manob_tmp = st.text_area("Escreva aqui:", value=st.session_state.obs_manobras, key='k_manob')
                if st.button("SALVAR COMENTÁRIO", key="b_manob"):
                    st.session_state.obs_manobras = obs_manob_tmp
                    st.success("Salvo com sucesso.")

            with aba_relatorio:
                st.markdown("#### Consolidação do Relatório Técnico de Telemetria (RTT)")

                col_t1, col_t2, col_t3, col_t4, col_t5 = st.columns(5)
                with col_t1: temp_ac = st.text_input("Acionamento:", tempos.get('acionamento', '-'))
                with col_t2: temp_dec = st.text_input("Decolagem:", tempos.get('decolagem', '-'))
                with col_t3: temp_pou = st.text_input("Pouso:", tempos.get('pouso', '-'))
                with col_t4: temp_cor = st.text_input("Corte:", tempos.get('corte', '-'))
                with col_t5: temp_voo = st.text_input("Tempo de Voo:", tempos.get('t_voo', '-'))
                tempos_editados = {'acionamento': temp_ac, 'decolagem': temp_dec, 'pouso': temp_pou, 'corte': temp_cor, 't_voo': temp_voo}

                alertas_txt = st.text_area("Registro de Alertas e Excedências (Injetado na Tabela 1 do Word):", value="\n".join(alertas), height=110)

                st.divider()
                st.markdown(f"""
                <div style="background: {T['bg_card_grad']}; border: 1px solid {T['accent']}; border-left: 5px solid {T['accent']}; border-radius: 10px; padding: 16px 20px; margin-bottom: 12px;">
                    <div style="font-size: 14px; font-weight: 700; color: {T['accent']};">Inteligência Artificial de Análise de Voo e Performance ({projeto_atual})</div>
                    <div style="font-size: 12.5px; color: {T['txt_sec']}; margin-top: 4px;">Gera automaticamente a análise técnica detalhada do perfil de voo, comportamento aerodinâmico/estrutural da aeronave e performance termodinâmica e mecânica do motor.</div>
                </div>
                """, unsafe_allow_html=True)

                if st.button("GERAR ANÁLISE TÉCNICA COMPLETA DO VOO COM IA", key="btn_ia_solo_rel", use_container_width=True):
                    def _safe(df_ref, col, fn="max", fmt="{:.1f}"):
                        if col not in df_ref.columns:
                            return "N/D"
                        s = df_ref[col].dropna()
                        if s.empty:
                            return "N/D"
                        if fn == "max": val = s.max()
                        elif fn == "min": val = s.min()
                        elif fn == "mean": val = s.mean()
                        elif fn == "p95": val = s.quantile(0.95)
                        elif fn == "std": val = s.std()
                        else: val = s.max()
                        return fmt.format(val)

                    df_voo = voo_ativo if not voo_ativo.empty else df
                    df_mot = df[df['E1 NG'] > 50] if 'E1 NG' in df.columns else df_voo

                    seg_g3 = int((df_voo['NormAc'] >= 3.0).sum()) if 'NormAc' in df_voo.columns else 0
                    seg_g4 = int((df_voo['NormAc'] >= 4.0).sum()) if 'NormAc' in df_voo.columns else 0
                    if 'NormAc' in df_voo.columns and not df_voo.empty:
                        inv = df_voo['NormAc'] < -0.15
                        seg_inv_max = int(inv.groupby((~inv).cumsum()).sum().max())
                    else:
                        seg_inv_max = 0

                    amostras_acro = 0
                    if 'Pitch' in df_voo.columns and 'Roll' in df_voo.columns:
                        amostras_acro = int(((df_voo['Pitch'].abs() > 45) | (df_voo['Roll'].abs() > 60)).sum())

                    fase_final = df_voo.tail(300) if not df_voo.empty else df
                    ias_final = _safe(fase_final, 'IAS', 'mean', '{:.1f}')
                    g_pouso = _safe(fase_final, 'NormAc', 'max', '{:.2f}')

                    itt_no_tq_max, ng_no_tq_max, alt_no_tq_max = "N/D", "N/D", "N/D"
                    if 'E1 Torq ft lbs' in df_voo.columns and not df_voo['E1 Torq ft lbs'].dropna().empty:
                        idx_tq = df_voo['E1 Torq ft lbs'].idxmax()
                        row_tq = df_voo.loc[idx_tq]
                        if 'E1 ITT' in df_voo.columns and pd.notna(row_tq['E1 ITT']): itt_no_tq_max = f"{row_tq['E1 ITT']:.0f} ºC"
                        if 'E1 NG' in df_voo.columns and pd.notna(row_tq['E1 NG']): ng_no_tq_max = f"{row_tq['E1 NG']:.1f}%"
                        if 'AltMSL' in df_voo.columns and pd.notna(row_tq['AltMSL']): alt_no_tq_max = f"{row_tq['AltMSL']:.0f} ft"

                    pico_tq = df['E1 Torq ft lbs'].max() if 'E1 Torq ft lbs' in df.columns else 0
                    pico_itt = df['E1 ITT'].max() if 'E1 ITT' in df.columns else 0
                    pico_ng = df['E1 NG'].max() if 'E1 NG' in df.columns else 0
                    pico_nh = df['E1 NP %'].max() if 'E1 NP %' in df.columns else 0

                    with st.spinner("Processando análise completa do perfil de voo, célula e motor..."):
                        try:
                            texto_alertas = "\n".join([f"- {a}" for a in alertas]) if alertas else "- Nenhuma excedência de arco amarelo ou vermelho registrada."

                            prompt = f"""
Elabore uma ANÁLISE TÉCNICA DETALHADA do voo abaixo, analisando EXCLUSIVAMENTE os parâmetros monitorados pela telemetria da aeronave {projeto_atual}.

=== 1. CRONOLOGIA REAL DO VOO ===
- Aeronave / Perfil: {projeto_atual} ({tipo_projeto} | {qtd_motores_config} motor(es))
- Acionamento: {tempos.get('acionamento', 'N/D')} | Decolagem: {tempos.get('decolagem', 'N/D')}
- Pouso: {tempos.get('pouso', 'N/D')} | Corte: {tempos.get('corte', 'N/D')}
- Tempo Efetivo de Voo: {tempos.get('t_voo', 'N/D')}

=== 2. PARÂMETROS DE VOO, ATITUDE E CÉLULA (EM VOO) ===
- Altitude MSL (ft): Média {_safe(df_voo, 'AltMSL', 'mean', '{:.0f}')} ft | Máxima {_safe(df_voo, 'AltMSL', 'max', '{:.0f}')} ft
- Velocidade Indicada IAS (kt): Média {_safe(df_voo, 'IAS', 'mean', '{:.1f}')} kt | P95 {_safe(df_voo, 'IAS', 'p95', '{:.1f}')} kt | Máxima {_safe(df_voo, 'IAS', 'max', '{:.1f}')} kt
- Arfagem Pitch (º): Mínima {_safe(df_voo, 'Pitch', 'min', '{:.1f}')}º | Máxima {_safe(df_voo, 'Pitch', 'max', '{:.1f}')}º
- Rolagem Roll (º): Mínima {_safe(df_voo, 'Roll', 'min', '{:.1f}')}º | Máxima {_safe(df_voo, 'Roll', 'max', '{:.1f}')}º
- Registros em Atitude Elevada (|Pitch| > 45º ou |Roll| > 60º): {amostras_acro} amostras
- Força G (NormAc): Mínima {min_g:.2f}G | Média {_safe(df_voo, 'NormAc', 'mean', '{:.2f}')}G | Máxima +{max_g:.2f}G (Limites: {l_g_min}G a +{l_g_max}G)
- Exposição a Cargas G: {seg_g3} amostras >= +3.0G | {seg_g4} amostras >= +4.0G | Maior sequência em voo invertido (<-0.15G): {seg_inv_max}s
- Aproximação Final e Pouso (últimas 300 amostras em voo): IAS média {ias_final} kt | Pico de Força G no toque {g_pouso}G

=== 3. PARÂMETROS DO MOTOR 1 ===
- Torque (lb.ft): Média em voo {_safe(df_voo, 'E1 Torq ft lbs', 'mean', '{:.0f}')} lb.ft | P95 {_safe(df_voo, 'E1 Torq ft lbs', 'p95', '{:.0f}')} lb.ft | Pico Máximo {_safe(df, 'E1 Torq ft lbs', 'max', '{:.0f}')} lb.ft | Oscilação (Std) {_safe(df_voo, 'E1 Torq ft lbs', 'std', '{:.0f}')} lb.ft (Limite Red Line: {l_tq_red} lb.ft)
- Temperatura T5 / ITT (ºC): Média em voo {_safe(df_voo, 'E1 ITT', 'mean', '{:.0f}')} ºC | P95 {_safe(df_voo, 'E1 ITT', 'p95', '{:.0f}')} ºC | Pico Máximo {_safe(df, 'E1 ITT', 'max', '{:.0f}')} ºC (Amarelo: {l_itt_yel} ºC | Red Line: {l_itt_red} ºC)
- Gerador de Gás Ng (%): Média em voo {_safe(df_voo, 'E1 NG', 'mean', '{:.1f}')}% | P95 {_safe(df_voo, 'E1 NG', 'p95', '{:.1f}')}% | Pico Máximo {_safe(df, 'E1 NG', 'max', '{:.1f}')}% (Amarelo: {l_ng_yel}% | Red Line: {l_ng_red}%)
- Rotação da Hélice Nh (%): Mínima em voo {_safe(df_voo, 'E1 NP %', 'min', '{:.1f}')}% | Média {_safe(df_voo, 'E1 NP %', 'mean', '{:.1f}')}% | Pico Máximo {_safe(df, 'E1 NP %', 'max', '{:.1f}')}% (Red Line: {l_nh_red}%)
- Relação no Instante do Torque Máximo: T5/ITT = {itt_no_tq_max} | Ng = {ng_no_tq_max} | Altitude = {alt_no_tq_max}

=== 4. SISTEMA DE LUBRIFICAÇÃO (MOTOR EM REGIME Ng > 50%) ===
- Temperatura do Óleo (OilT): Média {_safe(df_mot, 'E1 OilT_Celsius', 'mean', '{:.0f}')} ºC | Máxima {_safe(df_mot, 'E1 OilT_Celsius', 'max', '{:.0f}')} ºC (Limite Red Line: {l_oilt_red} ºC)
- Pressão do Óleo (OilP): Mínima {_safe(df_mot, 'E1 OilP', 'min', '{:.0f}')} PSI | Média {_safe(df_mot, 'E1 OilP', 'mean', '{:.0f}')} PSI | Máxima {_safe(df_mot, 'E1 OilP', 'max', '{:.0f}')} PSI (Faixa: {l_oilp_min} a {l_oilp_red} PSI)

=== 5. EXCEDÊNCIAS DETECTADAS PELO SISTEMA ===
{texto_alertas}

=== INSTRUÇÕES DE ESTRUTURA (OBRIGATÓRIO) ===
- NÃO mencione kW, HP, kg/s, toneladas, pressão atmosférica, clima ou resolução de panes/manutenção corretiva.
- Na seção 5, analise cada excedência listada acima informando horário, altitude e velocidade — destacando explicitamente se ocorreu no solo (Vel: 0 kt) ou em voo.
- Use exatamente `#`, `##`, `*` e **negrito** para os valores e sensores.

# ANÁLISE TÉCNICA DETALHADA DE VOO E PERFORMANCE (SOLO)
## 1. CRONOLOGIA E PERFIL DE VOO (ALTITUDE E VELOCIDADE)
## 2. ATITUDES, ENVELOPE ESTRUTURAL (FORÇA G) E POUSO
## 3. PERFORMANCE DO MOTOR (TORQUE, T5/ITT, NG E NH)
## 4. SISTEMA DE LUBRIFICAÇÃO (OILT E OILP)
## 5. DETALHAMENTO DAS EXCEDÊNCIAS E MARGENS OPERACIONAIS
"""

                            texto_ia_solo = gerar_texto_ia(prompt)
                            st.session_state.obs_graficos = texto_ia_solo
                            st.session_state["_sync_solo_ia"] = True
                            mod_usado = st.session_state.get("ultimo_modelo_ia", "IA")
                            st.session_state.status_ia_solo = ("ok", f"Análise técnica gerada com sucesso (modelo: {mod_usado}).")
                            st.rerun()
                        except Exception as e:
                            st.session_state.obs_graficos = gerar_parecer_contingencia(
                                max_g, min_g, pico_tq, pico_itt, pico_ng, pico_nh, alertas
                            )
                            st.session_state["_sync_solo_ia"] = True
                            st.session_state.status_ia_solo = ("warn", f"API de IA indisponível ({e}). Foi gerado o resumo automático local.")
                            st.rerun()

                if st.session_state.status_ia_solo:
                    _tp, _tx = st.session_state.status_ia_solo
                    if _tp == "ok":
                        st.success(_tx)
                    else:
                        st.warning(_tx)

                obs_graf_tmp_rel = st.text_area(
                    "Análise Técnica do Voo (Editável para o Documento Oficial):",
                    key="k_graf_rel",
                    height=280
                )
                st.session_state.obs_graficos = obs_graf_tmp_rel
                if st.button("CONFIRMAR ALTERAÇÕES NO PARECER", key="b_graf_rel"):
                    st.session_state.obs_graficos = obs_graf_tmp_rel
                    st.session_state["_sync_solo_ia"] = True
                    st.success("Parecer atualizado e pronto para exportação.")

                obs_solo_atual = {
                    "Resumo de Painel e Rota": st.session_state.obs_resumo,
                    "Inspetoria e Graficos": st.session_state.obs_graficos,
                    "Avaliador de Manobras e Pouso": st.session_state.obs_manobras
                }

                graficos_solo = {
                    "Dashboard de Inspeção de Motores e Dinâmica": fig_dash
                }

                st.divider()
                docx_bytes_solo_editado = gerar_docx_solo_editado(tempos_editados, alertas_txt, obs_solo_atual, graficos_solo)
                st.download_button(
                    label="BAIXAR RELATÓRIO TÉCNICO OFICIAL (.DOCX)",
                    data=docx_bytes_solo_editado,
                    file_name=f"RTT_{projeto_atual.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
