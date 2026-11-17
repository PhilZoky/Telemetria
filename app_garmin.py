import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import simplekml
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm 
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
import re
import os   
import io
import json
import math
import requests
from datetime import datetime
import streamlit.components.v1 as components

try:
    import google.generativeai as genai
except ImportError:
    genai = None

# ==========================================
# CONFIGURAÇÃO DA PÁGINA E ESTILO AEROESPACIAL (UI/UX PRO)
# ==========================================
st.set_page_config(
    page_title="AeroTelemetry | Flight Data Monitoring",
    page_icon="✈️",
    layout="wide",
    initial_sidebar_state="expanded"
)

def injetar_estilo_profissional():
    st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;600;700&display=swap');

        html, body, [class*="css"] {
            font-family: 'Inter', sans-serif;
        }
        .stApp {
            background: radial-gradient(circle at 50% 0%, #0f172a 0%, #070b14 100%);
            color: #f8fafc;
        }
        .block-container {
            padding-top: 2rem !important;
            padding-bottom: 3rem !important;
            max-width: 96% !important;
        }

        [data-testid="stSidebar"] {
            background-color: #090d16 !important;
            border-right: 1px solid #1e293b !important;
        }
        [data-testid="stSidebar"] .block-container {
            padding-top: 1.5rem !important;
        }

        .stTabs [data-baseweb="tab-list"] {
            gap: 8px;
            background-color: #0f172a;
            padding: 6px;
            border-radius: 10px;
            border: 1px solid #1e293b;
        }
        .stTabs [data-baseweb="tab"] {
            height: 44px;
            border-radius: 7px;
            padding: 0 20px;
            background-color: transparent;
            color: #94a3b8;
            font-weight: 600;
            font-size: 13px;
            letter-spacing: 0.4px;
            border: none !important;
            transition: all 0.2s ease;
        }
        .stTabs [data-baseweb="tab"]:hover {
            color: #f8fafc;
            background-color: rgba(255, 255, 255, 0.04);
        }
        .stTabs [aria-selected="true"] {
            background: linear-gradient(135deg, #1d4ed8 0%, #1e40af 100%) !important;
            color: #ffffff !important;
            box-shadow: 0 4px 12px rgba(29, 78, 216, 0.35);
        }

        .stButton > button, .stDownloadButton > button {
            background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
            color: #f8fafc;
            border: 1px solid #334155;
            border-radius: 8px;
            padding: 0.55rem 1.2rem;
            font-weight: 600;
            font-size: 13.5px;
            letter-spacing: 0.3px;
            transition: all 0.2s ease;
            box-shadow: 0 2px 6px rgba(0, 0, 0, 0.25);
        }
        .stButton > button:hover, .stDownloadButton > button:hover {
            border-color: #3b82f6;
            background: linear-gradient(135deg, #1d4ed8 0%, #1e3a8a 100%);
            color: #ffffff;
            transform: translateY(-1px);
            box-shadow: 0 4px 14px rgba(59, 130, 246, 0.35);
        }

        [data-testid="stFileUploader"] {
            background-color: rgba(15, 23, 42, 0.65);
            border: 1px dashed #334155;
            border-radius: 12px;
            padding: 14px;
            transition: border 0.2s ease;
        }
        [data-testid="stFileUploader"]:hover {
            border-color: #3b82f6;
        }

        .stTextInput > div > div > input,
        .stNumberInput > div > div > input,
        .stTextArea > div > div > textarea,
        .stSelectbox > div > div {
            background-color: #0f172a !important;
            border: 1px solid #1e293b !important;
            color: #f8fafc !important;
            border-radius: 8px !important;
        }
        .stTextInput > div > div > input:focus,
        .stTextArea > div > div > textarea:focus {
            border-color: #3b82f6 !important;
            box-shadow: 0 0 0 1px #3b82f6 !important;
        }

        .streamlit-expanderHeader {
            background-color: #0f172a !important;
            border: 1px solid #1e293b !important;
            border-radius: 8px !important;
            font-size: 13px !important;
            font-weight: 600 !important;
        }

        h1, h2, h3 {
            letter-spacing: -0.3px;
            font-weight: 700 !important;
        }
        hr {
            border-color: #1e293b !important;
            margin: 1.8rem 0 !important;
        }
    </style>
    """, unsafe_allow_html=True)

injetar_estilo_profissional()


# ==========================================
# CONFIGURAÇÃO DA INTELIGÊNCIA ARTIFICIAL
# ==========================================
# Lembrete: No VS Code pode manter a chave gsk_... aqui. Ao subir para o GitHub, deixe PADRAO_API_KEY = ""
PADRAO_API_KEY = "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"

def obter_credencial_ia():
    try:
        if "GEMINI_API_KEY" in st.secrets:
            return str(st.secrets["GEMINI_API_KEY"]).strip()
    except Exception:
        pass
    return PADRAO_API_KEY.strip()

def gerar_texto_ia(prompt_texto):
    credencial = obter_credencial_ia()
    
    if credencial.startswith("gsk_"):
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {credencial}"
        }
        
        modelos_groq = [
            "llama-3.3-70b-specdec",
            "llama-3.1-8b-instant",
            "llama3-70b-8192",
            "llama3-8b-8192",
            "mixtral-8x7b-32768",
            "gemma2-9b-it"
        ]
        
        try:
            r_mod = requests.get("https://api.groq.com/openai/v1/models", headers=headers, timeout=15)
            if r_mod.status_code == 200:
                ativos = [
                    m["id"] for m in r_mod.json().get("data", [])
                    if not any(x in m["id"].lower() for x in ["whisper", "guard", "tts", "vision", "embed"])
                ]
                preferidos = [m for m in modelos_groq if m in ativos]
                outros = [m for m in ativos if m not in preferidos]
                if preferidos or outros:
                    modelos_groq = preferidos + outros
        except Exception:
            pass

        ultimo_erro_groq = ""
        url = "https://api.groq.com/openai/v1/chat/completions"
        
        for mod in modelos_groq:
            payload = {
                "model": mod,
                "messages": [{"role": "user", "content": prompt_texto}],
                "temperature": 0.3
            }
            resp = requests.post(url, headers=headers, json=payload, timeout=60)
            if resp.status_code == 200:
                return resp.json()["choices"][0]["message"]["content"]
            ultimo_erro_groq = f"Erro na Groq ({resp.status_code}) [{mod}]: {resp.text}"
            
        raise RuntimeError(ultimo_erro_groq)

    modelos_tentativa = ["gemini-1.5-flash", "gemini-2.0-flash", "gemini-1.5-pro", "gemini-pro"]
    payload = {"contents": [{"parts": [{"text": prompt_texto}]}]}
    ultimo_erro = ""
    
    for mod in modelos_tentativa:
        url_key = f"https://generativelanguage.googleapis.com/v1beta/models/{mod}:generateContent?key={credencial}"
        resp = requests.post(url_key, headers={"Content-Type": "application/json"}, json=payload, timeout=60)
        if resp.status_code == 200:
            return resp.json()["candidates"][0]["content"]["parts"][0]["text"]
            
        if credencial.startswith("AQ."):
            url_bearer = f"https://generativelanguage.googleapis.com/v1beta/models/{mod}:generateContent"
            headers_b = {"Content-Type": "application/json", "Authorization": f"Bearer {credencial}"}
            resp_b = requests.post(url_bearer, headers=headers_b, json=payload, timeout=60)
            if resp_b.status_code == 200:
                return resp_b.json()["candidates"][0]["content"]["parts"][0]["text"]
                
        ultimo_erro = f"Status {resp.status_code}: {resp.text}"

    raise RuntimeError(f"Não foi possível autenticar na API. Detalhe: {ultimo_erro}")


# ==========================================
# 0. SISTEMA DE LOGIN E CADASTRO (SaaS)
# ==========================================
ARQUIVO_USUARIOS = "banco_usuarios.json"

def carregar_usuarios():
    if os.path.exists(ARQUIVO_USUARIOS):
        try:
            with open(ARQUIVO_USUARIOS, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return criar_admin_padrao()
    return criar_admin_padrao()

def criar_admin_padrao():
    return {
        "Filipe_Iack": {
            "senha": "210772",
            "nome": "Filipe Iack",
            "email": "admin@telemetria.com",
            "organizacao": "AFA - Master"
        }
    }

def salvar_usuarios(usuarios):
    with open(ARQUIVO_USUARIOS, "w", encoding="utf-8") as f:
        json.dump(usuarios, f, indent=4)

usuarios_permitidos = carregar_usuarios()

if "usuario_logado" not in st.session_state:
    st.session_state.usuario_logado = None
if "nome_logado" not in st.session_state:
    st.session_state.nome_logado = None

if st.session_state.usuario_logado is None:
    st.markdown("<br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1.2, 1.8, 1.2])
    
    with col2:
        st.markdown("""
        <div style="background: linear-gradient(180deg, #0f172a 0%, #090d16 100%); border: 1px solid #1e293b; border-top: 3px solid #3b82f6; border-radius: 14px; padding: 28px 32px; box-shadow: 0 20px 40px rgba(0,0,0,0.5); text-align: center; margin-bottom: 18px;">
            <div style="display: inline-block; background: rgba(59, 130, 246, 0.12); border: 1px solid rgba(59, 130, 246, 0.3); color: #60a5fa; font-size: 11px; font-weight: 700; letter-spacing: 1.5px; padding: 4px 12px; border-radius: 20px; margin-bottom: 12px;">
                FLIGHT DATA MONITORING SYSTEM
            </div>
            <h2 style="margin: 0; color: #f8fafc; font-size: 26px; font-weight: 700;">✈️ AeroTelemetry Tactical</h2>
            <p style="margin: 6px 0 0 0; color: #94a3b8; font-size: 13.5px;">Plataforma Integrada de Análise de Voo, Motores e Inspetoria Técnica</p>
        </div>
        """, unsafe_allow_html=True)
        
        aba_login, aba_cadastro = st.tabs(["🔐 ACESSAR CREDENCIAIS", "📋 CADASTRAR OPERADOR"])
        
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
                        st.error("❌ Credenciais inválidas. Verifique seu ID e Senha.")
                else:
                    st.error("❌ Operador não localizado na base de dados.")
                    
        with aba_cadastro:
            st.markdown("<div style='height: 8px'></div>", unsafe_allow_html=True)
            novo_nome = st.text_input("Nome e Posto / Graduação", key="cad_nome")
            novo_email = st.text_input("E-mail Institucional / Contato", key="cad_email")
            nova_org = st.text_input("Esquadrão / Unidade Aérea / Empresa", key="cad_org")
            
            c_cad1, c_cad2 = st.columns(2)
            with c_cad1:
                novo_usuario = st.text_input("ID de Login", key="cad_user")
            with c_cad2:
                nova_senha = st.text_input("Definir Senha", type="password", key="cad_pass")
            
            if st.button("REGISTRAR OPERADOR", use_container_width=True):
                if novo_usuario in usuarios_permitidos:
                    st.warning("⚠️ Este ID de usuário já está em uso.")
                elif not novo_usuario or not nova_senha or not novo_nome or not novo_email:
                    st.warning("⚠️ Preencha todos os campos obrigatórios.")
                else:
                    usuarios_permitidos[novo_usuario] = {
                        "senha": nova_senha,
                        "nome": novo_nome,
                        "email": novo_email,
                        "organizacao": nova_org,
                        "data_cadastro": datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    }
                    salvar_usuarios(usuarios_permitidos)
                    st.success(f"✅ Credencial de {novo_nome} habilitada! Acesse pela aba de Login.")
                    
    st.stop()


# ==========================================
# 1. GERENCIADOR DE PERFIS INDIVIDUAIS DO CLIENTE (ATUALIZADO "BOA CEPA!" T-27M)
# ==========================================
ARQUIVO_PERFIS = f"perfis_{st.session_state.usuario_logado}.json"

PERFIS_PADRAO = {
    "T-27M (Oficial)": {
        "tipo_aeronave": "Asa Fixa",
        "qtd_motores": 1,
        "g_max": 6.0, "g_min": -3.0,          # Flapes Recolhidos: -3G a +6G
        "tq_red": 1795.0, "tq_yel": 1795.0,   # Torque Máx Decolagem: 1795 lb.ft
        "itt_red": 790.0, "itt_yel": 740.0,   # T5 Máx Decolagem: 790°C | T5 Máx Cruzeiro: 740°C
        "ng_red": 101.5, "ng_yel": 100.0,     # Ng Máx Dec/Sub/Cruz: 101,5%
        "nh_red": 100.0, "nh_yel": 100.0,     # Nh Máx Dec/Sub/Cruz: 100%
        "oilt_red": 99.0, "oilt_min_green": 10.0,
        "oilp_red": 95.0, "oilp_min_green": 40.0 # Luz Press Óleo abaixo de 40 psi
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
    if os.path.exists(ARQUIVO_PERFIS):
        try:
            with open(ARQUIVO_PERFIS, "r", encoding="utf-8") as f: 
                return json.load(f)
        except: 
            return PERFIS_PADRAO
    return PERFIS_PADRAO

def salvar_perfis(perfis):
    with open(ARQUIVO_PERFIS, "w", encoding="utf-8") as f: 
        json.dump(perfis, f, indent=4)

perfis_salvos = carregar_perfis()


# ==========================================
# 2. INICIALIZAÇÃO DE VARIÁVEIS DE SESSÃO
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


# ==========================================
# 3. BARRA LATERAL (PAINEL DE COMANDO TÁTICO)
# ==========================================
st.sidebar.markdown(f"""
<div style="background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%); border: 1px solid #334155; border-left: 4px solid #3b82f6; border-radius: 10px; padding: 14px 16px; margin-bottom: 12px;">
    <div style="font-size: 10px; font-weight: 700; color: #60a5fa; letter-spacing: 1.2px; text-transform: uppercase;">INSPETOR CONECTADO</div>
    <div style="font-size: 15px; font-weight: 700; color: #f8fafc; margin-top: 3px;">👨‍✈️ {st.session_state.nome_logado}</div>
    <div style="font-size: 11px; color: #94a3b8; margin-top: 2px;">Status: <span style="color: #10b981; font-weight: 600;">● Online / Operacional</span></div>
</div>
""", unsafe_allow_html=True)

if st.sidebar.button("🚪 Encerrar Sessão (Logout)", use_container_width=True):
    st.session_state.usuario_logado = None
    st.session_state.nome_logado = None
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.markdown("<div style='font-size: 11px; font-weight: 700; color: #64748b; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 8px;'>✈️ CONFIGURAÇÃO DA AERONAVE</div>", unsafe_allow_html=True)

projeto_atual = st.sidebar.selectbox("Perfil de Envelope:", list(perfis_salvos.keys()))
p_dados = perfis_salvos[projeto_atual]

tipo_projeto = p_dados.get("tipo_aeronave", "Asa Fixa")
qtd_motores_config = int(p_dados.get("qtd_motores", 1))

st.sidebar.markdown(f"""
<div style="display: flex; gap: 8px; margin-top: 6px; margin-bottom: 10px;">
    <div style="flex: 1; background: #0f172a; border: 1px solid #1e293b; border-radius: 8px; padding: 8px 10px; text-align: center;">
        <div style="font-size: 10px; color: #64748b; text-transform: uppercase;">Categoria</div>
        <div style="font-size: 13px; font-weight: 700; color: #e2e8f0; margin-top: 2px;">{tipo_projeto}</div>
    </div>
    <div style="flex: 1; background: #0f172a; border: 1px solid #1e293b; border-radius: 8px; padding: 8px 10px; text-align: center;">
        <div style="font-size: 10px; color: #64748b; text-transform: uppercase;">Propulsão</div>
        <div style="font-size: 13px; font-weight: 700; color: #38bdf8; margin-top: 2px;">{qtd_motores_config} Motor(es)</div>
    </div>
</div>
""", unsafe_allow_html=True)

if projeto_atual == "T-27M (Oficial)":
    if st.sidebar.button("🔄 Restaurar Padrão 'BOA CEPA!' T-27M", use_container_width=True):
        perfis_salvos["T-27M (Oficial)"] = PERFIS_PADRAO["T-27M (Oficial)"].copy()
        salvar_perfis(perfis_salvos)
        st.sidebar.success("Limites oficiais do T-27M aplicados!")
        st.rerun()

st.sidebar.markdown("---")
st.sidebar.markdown("<div style='font-size: 11px; font-weight: 700; color: #64748b; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 8px;'>🕹️ MODO DE MISSÃO</div>", unsafe_allow_html=True)
modo_analise = st.sidebar.radio(
    "Selecione o perfil de análise:", 
    ["Voo Solo / Comparativo de Máquina", "Voo em Formatura (Líder e Ala)"],
    label_visibility="collapsed"
)
st.sidebar.markdown("---")

with st.sidebar.expander("➕ Cadastrar Nova Aeronave", expanded=False):
    novo_projeto = st.text_input("Designação da Aeronave:")
    nova_classe = st.selectbox("Categoria:", ["Asa Fixa", "Asa Rotativa"])
    qtd_mot_novo = st.number_input("Nº de Motores:", min_value=1, max_value=6, value=1)
    if st.button("Salvar Novo Perfil", use_container_width=True):
        if novo_projeto and novo_projeto not in perfis_salvos:
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
            st.success(f"Aeronave '{novo_projeto}' cadastrada!")
            st.rerun()

with st.sidebar.expander("⚙️ Limites de Envelope (Arcos)", expanded=False):
    st.caption("Parametrização dos limites estruturais e de motor.")
    
    novo_tipo_aeronave = st.selectbox("Categoria Atual:", ["Asa Fixa", "Asa Rotativa"], index=0 if tipo_projeto == "Asa Fixa" else 1)
    nova_qtd_mot = st.number_input("Quantidade de Motores Atual:", min_value=1, max_value=6, value=qtd_motores_config)
    st.markdown("---")
    
    def ler_limite(label, chave):
        val = st.text_input(label, value=str(p_dados.get(chave, "")))
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

    if st.button("💾 Atualizar Envelope", use_container_width=True):
        perfis_salvos[projeto_atual].update({
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
        })
        perfis_salvos[projeto_atual] = {k: (v if v is not None else "") for k, v in perfis_salvos[projeto_atual].items()}
        salvar_perfis(perfis_salvos)
        st.success(f"Envelope do {projeto_atual} atualizado!")
        st.rerun()


# ==========================================
# 4. CABEÇALHO EXECUTIVO E FUNÇÕES MATEMÁTICAS
# ==========================================
st.markdown(f"""
<div style="background: linear-gradient(90deg, #0f172a 0%, #1e293b 50%, #0f172a 100%); border: 1px solid #1e293b; border-bottom: 2px solid #3b82f6; border-radius: 12px; padding: 18px 24px; margin-bottom: 20px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 15px;">
    <div>
        <div style="display: flex; align-items: center; gap: 10px;">
            <span style="background: #1d4ed8; color: #fff; font-size: 10px; font-weight: 800; padding: 3px 8px; border-radius: 4px; letter-spacing: 1px;">GARMIN G1000 / G3000</span>
            <span style="color: #64748b; font-size: 12px; font-family: 'JetBrains Mono', monospace;">FDM ENGINE v3.6 • BOA CEPA EDITION</span>
        </div>
        <h1 style="margin: 6px 0 0 0; font-size: 26px; color: #f8fafc;">Centro de Telemetria e Inspetoria de Voo</h1>
    </div>
    <div style="display: flex; gap: 12px; align-items: center;">
        <div style="background: #090d16; border: 1px solid #1e293b; border-radius: 8px; padding: 8px 16px; text-align: right;">
            <div style="font-size: 10px; color: #64748b; text-transform: uppercase; letter-spacing: 0.8px;">Aeronave Ativa</div>
            <div style="font-size: 15px; font-weight: 700; color: #38bdf8; font-family: 'JetBrains Mono', monospace;">{projeto_atual}</div>
        </div>
        <div style="background: #090d16; border: 1px solid #1e293b; border-radius: 8px; padding: 8px 16px; text-align: right;">
            <div style="font-size: 10px; color: #64748b; text-transform: uppercase; letter-spacing: 0.8px;">Modo Operacional</div>
            <div style="font-size: 14px; font-weight: 600; color: #e2e8f0;">{"Formatura Tática" if "Formatura" in modo_analise else "Solo / Trend Monitoring"}</div>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

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
    
    R = 6371000.0  # raio da Terra em metros
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
    <div style="background: linear-gradient(160deg, #0f172a 0%, #090d16 100%); padding: 16px 18px; border-radius: 12px; border: 1px solid #1e293b; border-top: 3px solid {cor}; margin-bottom: 16px; box-shadow: 0 8px 20px rgba(0,0,0,0.35); position: relative; overflow: hidden;">
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
            <span style="font-size: 11.5px; color: #94a3b8; font-weight: 700; text-transform: uppercase; letter-spacing: 0.6px;">{titulo}</span>
            <span style="font-size: 10px; color: {cor}; background: {bg_badge}; border: 1px solid {cor}44; padding: 2px 8px; border-radius: 12px; font-weight: 700; letter-spacing: 0.5px;">{status}</span>
        </div>
        <div style="display: flex; align-items: baseline; gap: 6px; margin-bottom: 12px;">
            <span style="color: #f8fafc; font-size: 30px; font-weight: 700; font-family: 'JetBrains Mono', monospace; letter-spacing: -0.5px;">{pico_str}</span>
            <span style="color: #64748b; font-size: 13px; font-weight: 600;">{unidade}</span>
        </div>
        <div style="width: 100%; background: #1e293b; height: 5px; border-radius: 4px; overflow: hidden; margin-bottom: 8px;">
            <div style="width: {pct_barra}%; background: {cor}; height: 100%; border-radius: 4px;"></div>
        </div>
        <div style="display: flex; justify-content: space-between; align-items: center;">
            <span style="font-size: 11px; color: #64748b;">Red Line Limit</span>
            <span style="font-size: 11.5px; color: #cbd5e1; font-family: 'JetBrains Mono', monospace; font-weight: 600;">{limite_str} {unidade}</span>
        </div>
    </div>
    """
    st.markdown(html, unsafe_allow_html=True)

def render_html_table(df_formatted, height=420):
    html_table = df_formatted.to_html(index=False, justify='center', border=0, classes='tabela-grande')
    html_completo = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&family=JetBrains+Mono:wght@400;600&display=swap" rel="stylesheet">
    <style>
    body {{ background-color: transparent; color: #f8fafc; font-family: 'Inter', sans-serif; margin: 0; }}
    .tabela-container {{ max-height: {height - 20}px; overflow-y: auto; overflow-x: auto; border: 1px solid #1e293b; border-radius: 10px; background: #090d16; box-shadow: inset 0 2px 8px rgba(0,0,0,0.4); }}
    .tabela-container::-webkit-scrollbar {{ width: 8px; height: 8px; }}
    .tabela-container::-webkit-scrollbar-track {{ background: #090d16; }}
    .tabela-container::-webkit-scrollbar-thumb {{ background: #334155; border-radius: 4px; }}
    .tabela-container::-webkit-scrollbar-thumb:hover {{ background: #3b82f6; }}
    .tabela-grande {{ width: 100%; font-size: 13.5px; text-align: center; border-collapse: collapse; white-space: nowrap; }}
    .tabela-grande th {{ background: #0f172a; color: #94a3b8; font-size: 11.5px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.7px; padding: 13px 14px; position: sticky; top: 0; z-index: 1; border-bottom: 2px solid #1e293b; }}
    .tabela-grande td {{ padding: 10px 14px; border-bottom: 1px solid #131c2e; font-family: 'JetBrains Mono', monospace; color: #e2e8f0; }}
    .tabela-grande tr:nth-child(even) {{ background-color: rgba(15, 23, 42, 0.4); }}
    .tabela-grande tr:hover {{ background-color: rgba(59, 130, 246, 0.12); }}
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
                except Exception as e:
                    doc.add_paragraph(f"[Aviso: Não foi possível exportar a imagem '{titulo_grafico}'. O servidor pode estar sem a biblioteca 'kaleido'].")

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
        set_cell_background(hdr_cell, "333333")
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
            bg_cor = "FFF5F5" if "🔴" in al else ("FFFBEB" if "🟡" in al else "FFFFFF")
            set_cell_background(cell_al, bg_cor)
            adicionar_paragrafo_com_markdown(cell_al, al)
    else:
        adicionar_paragrafo_com_markdown(doc, "🟢 **Operação Normal:** Nenhum alerta ou ultrapassagem de parâmetro registrado nesta missão.")

    p_s2 = doc.add_paragraph()
    p_s2.paragraph_format.space_before = Pt(14)
    p_s2.paragraph_format.space_after = Pt(6)
    r_s2 = p_s2.add_run("2. ANÁLISE DE TELEMETRIA E PARECER TÉCNICO DE INSPETORIA")
    r_s2.bold = True
    r_s2.font.size = Pt(11)
    r_s2.font.color.rgb = COR_AZUL_FAB

    texto_ia = obs_dict.get("Inspetoria e Graficos", "")
    if str(texto_ia).strip():
        linhas = str(texto_ia).split('\n')
        for linha in linhas:
            l_limpa = linha.strip()
            if not l_limpa:
                continue

            if l_limpa.startswith("# "):
                tit = l_limpa.replace("# ", "").strip()
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(10)
                p.paragraph_format.space_after = Pt(4)
                r = p.add_run(tit)
                r.bold = True
                r.font.size = Pt(12)
                r.font.color.rgb = COR_AZUL_FAB

            elif l_limpa.startswith("## "):
                subtit = l_limpa.replace("## ", "").strip()
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(8)
                p.paragraph_format.space_after = Pt(4)
                r = p.add_run(subtit)
                r.bold = True
                r.font.size = Pt(10.5)
                r.font.color.rgb = COR_AZUL_FAB

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
            except Exception as e:
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
    titulo = doc.add_heading("RELATÓRIO TÁTICO DE FORMATURA", 0)
    titulo.alignment = 1
    data_hoje = datetime.now().strftime('%d/%m/%Y')
    doc.add_paragraph(f"Analista: {st.session_state.nome_logado} | Data: {data_hoje}").alignment = 1
    
    doc.add_heading("1. AERONAVES ENVOLVIDAS", level=1)
    doc.add_paragraph(f"Ficheiro Líder: {arq_lider}")
    doc.add_paragraph(f"Ficheiro Ala: {arq_ala}")
    
    doc.add_heading("2. DEBRIEFING E AVALIAÇÃO TÁTICA", level=1)
    tem_obs = False
    for titulo_obs, texto in obs_dict.items():
        if str(texto).strip() != "":
            tem_obs = True
            doc.add_heading(f"[{titulo_obs}]", level=2)
            doc.add_paragraph(str(texto))
    if not tem_obs: doc.add_paragraph("Nenhuma observação inserida no dossiê.")
        
    injetar_graficos_word(doc, graficos_dict)
        
    doc.add_paragraph("\n\n____________________________________________").alignment = 1
    doc.add_paragraph(f"{st.session_state.nome_logado} - Especialista em Telemetria").alignment = 1
    
    io_bytes = io.BytesIO()
    doc.save(io_bytes)
    return io_bytes.getvalue()

def gerar_docx_comparativo_editado(arq_a, arq_b, df_metricas, obs_dict, graficos_dict):
    doc = Document()
    titulo = doc.add_heading("RELATÓRIO DE TREND MONITORING (COMPARATIVO)", 0)
    titulo.alignment = 1

    data_hoje = datetime.now().strftime('%d/%m/%Y')
    doc.add_paragraph(f"Perfil: {projeto_atual} | Analista: {st.session_state.nome_logado} | Data: {data_hoje}").alignment = 1

    doc.add_heading("1. VOOS AVALIADOS", level=1)
    doc.add_paragraph(f"Voo A (Referência / Antigo): {arq_a}")
    doc.add_paragraph(f"Voo B (Avaliado / Recente): {arq_b}")

    doc.add_heading("2. ANÁLISE DE DESEMPENHO E DEGRADAÇÃO", level=1)
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

    doc.add_heading("3. PARECER TÉCNICO DA INSPETORIA", level=1)
    tem_obs = False
    for titulo_obs, texto in obs_dict.items():
        if str(texto).strip() != "":
            tem_obs = True
            doc.add_heading(f"[{titulo_obs}]", level=2)
            doc.add_paragraph(str(texto))

    if not tem_obs:
        doc.add_paragraph("Nenhuma observação inserida.")

    injetar_graficos_word(doc, graficos_dict)

    doc.add_paragraph("\n\n____________________________________________").alignment = 1
    doc.add_paragraph(f"{st.session_state.nome_logado} - Especialista em Telemetria").alignment = 1

    io_bytes = io.BytesIO()
    doc.save(io_bytes)
    return io_bytes.getvalue()


# ==========================================
# 7. UPLOADER E FLUXO PRINCIPAL
# ==========================================
arquivos_up = st.file_uploader(
    f"📂 Importar Log(s) de Dados do Cartão SD Garmin (.CSV) — Perfil Ativo: {projeto_atual}", 
    type=['csv'], 
    accept_multiple_files=True
)

if not arquivos_up:
    st.markdown("""
    <div style="background: rgba(15, 23, 42, 0.45); border: 1px solid #1e293b; border-radius: 12px; padding: 36px 24px; text-align: center; margin-top: 15px;">
        <div style="font-size: 38px; margin-bottom: 10px;">🛰️</div>
        <h3 style="margin: 0; color: #e2e8f0; font-size: 18px;">Aguardando Ingestão de Telemetria</h3>
        <p style="color: #64748b; font-size: 13.5px; max-width: 560px; margin: 8px auto 0 auto;">
            Arraste <b>1 arquivo CSV</b> acima para auditoria completa de Voo Solo, ou selecione <b>2 arquivos CSV</b> simultaneamente para Análise de Formatura (Líder/Ala) ou Trend Monitoring de Motores.
        </p>
    </div>
    """, unsafe_allow_html=True)

if arquivos_up:
    
    # ---------------------------------------------------------
    # FLUXO 1: VOO EM FORMATURA
    # ---------------------------------------------------------
    if modo_analise == "Voo em Formatura (Líder e Ala)":
        if len(arquivos_up) != 2:
            st.warning("⚠️ Para o Modo Formatura, carregue **exatamente 2 arquivos** simultaneamente (Líder e Ala).")
        else:
            st.markdown("### 🔀 Designação de Elemento / Esquadrilha")
            
            nome_arq1 = arquivos_up[0].name
            nome_arq2 = arquivos_up[1].name
            
            col_lider, col_ala, col_sync = st.columns([1.5, 1.5, 1])
            with col_lider:
                arq_lider = st.selectbox("🔵 Aeronave LÍDER (Nº 1):", [nome_arq1, nome_arq2], index=0)
            with col_ala:
                idx_padrao_ala = 1 if arq_lider == nome_arq1 else 0
                arq_ala = st.selectbox("🟠 Aeronave ALA (Nº 2):", [nome_arq1, nome_arq2], index=idx_padrao_ala)
            with col_sync:
                ajuste_horas = st.number_input("⏳ Offset Fuso Horário Ala (h):", min_value=-12, max_value=12, value=0, step=1, help="Use -3 ou +3 caso uma aeronave esteja em UTC e outra em Hora Local.")
                
            st.sidebar.markdown("---")
            st.sidebar.success("✅ **Telemetria Sincronizada**")

            if arq_lider == arq_ala:
                st.error("⚠️ Você selecionou o mesmo log para Líder e Ala. Altere um dos seletores acima.")
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
                    st.error("❌ Os relógios não coincidem. Ajuste o 'Offset Fuso Horário Ala' acima para -3 ou +3.")
                else:
                    aba_dist, aba_comp, aba_3d, aba_rel_form = st.tabs(["📏 1. SEPARAÇÃO TÁTICA", "⚙️ 2. COMPARATIVO DINÂMICO", "✈️ 3. ENVELOPE 3D", "📄 4. DOSSIÊ FINAL"])
                    
                    fig_dist = None
                    fig_comp = None
                    fig_3d = None

                    with aba_dist:
                        df_merge['Distancia_m'] = formula_haversine(
                            df_merge['Latitude_lider'], df_merge['Longitude_lider'], 
                            df_merge['Latitude_ala'], df_merge['Longitude_ala']
                        )
                        
                        fig_dist = make_subplots(rows=3, cols=1, shared_xaxes=True, vertical_spacing=0.07, subplot_titles=("1. Separação Horizontal Líder-Ala (Metros)", "2. Perfil Vertical de Altitude MSL (ft)", "3. Atitude Longitudinal / Pitch (º)"))
                        fig_dist.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['Distancia_m'], name="Separação (m)", fill='tozeroy', fillcolor='rgba(56, 189, 248, 0.12)', line=dict(color='#38bdf8', width=2)), row=1, col=1)
                        if 'AltMSL_lider' in df_merge.columns: fig_dist.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['AltMSL_lider'], name="Alt Líder", line=dict(color='#f97316', width=2)), row=2, col=1)
                        if 'AltMSL_ala' in df_merge.columns: fig_dist.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['AltMSL_ala'], name="Alt Ala", line=dict(color='#3b82f6', width=2, dash='dot')), row=2, col=1)
                        if 'Pitch_lider' in df_merge.columns: fig_dist.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['Pitch_lider'], name="Pitch Líder", line=dict(color='#f97316', width=2)), row=3, col=1)
                        if 'Pitch_ala' in df_merge.columns: fig_dist.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['Pitch_ala'], name="Pitch Ala", line=dict(color='#3b82f6', width=2, dash='dot')), row=3, col=1)
                        fig_dist.update_layout(template="plotly_dark", paper_bgcolor="#090d16", plot_bgcolor="#090d16", height=820, hovermode="x unified", margin=dict(l=15, r=15, t=35, b=15))
                        st.plotly_chart(fig_dist, use_container_width=True)
                        
                        st.subheader("🚨 Top 15 Momentos de Maior Afastamento na Formatura")
                        df_top_dist = df_merge.sort_values(by='Distancia_m', ascending=False).head(15).copy()
                        df_formatada = pd.DataFrame()
                        df_formatada['Hora (Líder)'] = df_top_dist['Datetime'].dt.strftime('%H:%M:%S')
                        df_formatada['Distância (m)'] = df_top_dist['Distancia_m'].apply(lambda x: f"{x:.1f} m")
                        if 'AltMSL_lider' in df_top_dist.columns: df_formatada['Alt. Líder (ft)'] = df_top_dist['AltMSL_lider'].apply(lambda x: f"{x:.0f}")
                        if 'Pitch_lider' in df_top_dist.columns: df_formatada['Arfagem Líder (º)'] = df_top_dist['Pitch_lider'].apply(lambda x: f"{x:.1f}")
                        render_html_table(df_formatada)

                        st.divider()
                        st.markdown("📝 **Observações Táticas - Separação e Posicionamento**")
                        obs_dist_tmp = st.text_area("Registar apontamentos sobre manutenção de posição:", value=st.session_state.obs_distancia, key='k_dist')
                        if st.button("💾 Salvar Apontamento", key="b_dist"): 
                            st.session_state.obs_distancia = obs_dist_tmp
                            st.success("Apontamento gravado!")

                    with aba_comp:
                        fig_comp = make_subplots(rows=6, cols=1, shared_xaxes=True, vertical_spacing=0.03, subplot_titles=("1. Altitude MSL (ft)", "2. Fator de Carga (Força G)", "3. Torque (Lbs)", "4. Temperatura T5 / ITT (ºC)", "5. Rotação Ng (%)", "6. Rotação Hélice Nh (%)"))
                        if 'AltMSL_lider' in df_merge.columns: fig_comp.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['AltMSL_lider'], name="Alt Líder", line=dict(color='#f97316', width=2)), row=1, col=1)
                        if 'AltMSL_ala' in df_merge.columns: fig_comp.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['AltMSL_ala'], name="Alt Ala", line=dict(color='#3b82f6', width=2, dash='dot')), row=1, col=1)
                        if 'NormAc_lider' in df_merge.columns: fig_comp.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['NormAc_lider'], name="G Líder", line=dict(color='#f97316', width=2)), row=2, col=1)
                        if 'NormAc_ala' in df_merge.columns: fig_comp.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['NormAc_ala'], name="G Ala", line=dict(color='#3b82f6', width=2)), row=2, col=1)
                        if 'E1 Torq ft lbs_lider' in df_merge.columns: fig_comp.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['E1 Torq ft lbs_lider'], name="Tq Líder", line=dict(color='#f97316', width=2)), row=3, col=1)
                        if 'E1 Torq ft lbs_ala' in df_merge.columns: fig_comp.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['E1 Torq ft lbs_ala'], name="Tq Ala", line=dict(color='#3b82f6', width=2)), row=3, col=1)
                        if 'E1 ITT_lider' in df_merge.columns: fig_comp.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['E1 ITT_lider'], name="T5 Líder", line=dict(color='#f97316', width=2)), row=4, col=1)
                        if 'E1 ITT_ala' in df_merge.columns: fig_comp.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['E1 ITT_ala'], name="T5 Ala", line=dict(color='#3b82f6', width=2)), row=4, col=1)
                        if 'E1 NG_lider' in df_merge.columns: fig_comp.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['E1 NG_lider'], name="Ng Líder", line=dict(color='#f97316', width=2)), row=5, col=1)
                        if 'E1 NG_ala' in df_merge.columns: fig_comp.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['E1 NG_ala'], name="Ng Ala", line=dict(color='#3b82f6', width=2)), row=5, col=1)
                        if 'E1 NP %_lider' in df_merge.columns: fig_comp.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['E1 NP %_lider'], name="Nh Líder", line=dict(color='#f97316', width=2)), row=6, col=1)
                        if 'E1 NP %_ala' in df_merge.columns: fig_comp.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['E1 NP %_ala'], name="Nh Ala", line=dict(color='#3b82f6', width=2)), row=6, col=1)
                        fig_comp.update_layout(template="plotly_dark", paper_bgcolor="#090d16", plot_bgcolor="#090d16", height=1150, hovermode="x unified", margin=dict(l=15, r=15, t=35, b=15))
                        st.plotly_chart(fig_comp, use_container_width=True)

                        st.divider()
                        st.markdown("📝 **Observações - Sincronismo de Potência e Motores**")
                        obs_comp_tmp = st.text_area("Escreva aqui:", value=st.session_state.obs_comparativo, key='k_comp')
                        if st.button("💾 Salvar Comentário", key="b_comp"): 
                            st.session_state.obs_comparativo = obs_comp_tmp
                            st.success("Salvo com sucesso!")

                    with aba_3d:
                        fig_3d = go.Figure()
                        fig_3d.add_trace(go.Scatter3d(x=df_merge['Longitude_lider'], y=df_merge['Latitude_lider'], z=df_merge['AltMSL_lider'], mode='lines', line=dict(color='#f97316', width=4), name="Líder"))
                        fig_3d.add_trace(go.Scatter3d(x=df_merge['Longitude_ala'], y=df_merge['Latitude_ala'], z=df_merge['AltMSL_ala'], mode='lines', line=dict(color='#3b82f6', width=4), name="Ala"))
                        fig_3d.update_layout(template="plotly_dark", paper_bgcolor="#090d16", height=700, margin=dict(l=0, r=0, b=0, t=20), scene=dict(aspectmode='auto'))
                        st.plotly_chart(fig_3d, use_container_width=True)
                        st.divider()
                        st.markdown("📝 **Observações - Trajetória Espacial 3D**")
                        obs_3d_tmp = st.text_area("Escreva aqui:", value=st.session_state.obs_3d, key='k_3d')
                        if st.button("💾 Salvar Comentário", key="b_3d"): 
                            st.session_state.obs_3d = obs_3d_tmp
                            st.success("Salvo com sucesso!")

                    with aba_rel_form:
                        st.subheader("📝 Emissão de Dossiê Tático de Formatura")
                        c1, c2 = st.columns(2)
                        with c1: lider_editado = st.text_input("Designação Líder:", arq_lider)
                        with c2: ala_editado = st.text_input("Designação Ala:", arq_ala)
                        
                        st.divider()
                        st.markdown("🪄 **Síntese Tática Automatizada por IA**")
                        if st.button("🤖 Gerar Debriefing de Formatura com IA", key="btn_ia_form"):
                            with st.spinner("A IA está processando a telemetria da esquadrilha..."):
                                try:
                                    prompt_form = f"""
                                    Atue como um Piloto Inspetor da Força Aérea Brasileira.
                                    Faça um breve debriefing técnico (máximo 2 parágrafos) de um voo em formatura.
                                    Aeronave Líder: {arq_lider}
                                    Aeronave Ala: {arq_ala}
                                    Escreva com jargão militar de aviação que a missão foi avaliada quanto à dispersão geométrica e aos parâmetros dinâmicos de motor (Ng, T5/ITT, Torque). Seja profissional e direto.
                                    """
                                    texto_ia_form = gerar_texto_ia(prompt_form)
                                    st.session_state.obs_comparativo = texto_ia_form
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Erro na IA: {e}")
                        
                        obs_formatura_atual = {
                            "Posicionamento e Distancia": st.session_state.obs_distancia,
                            "Comparativo Dinamico (Motores)": st.session_state.obs_comparativo,
                            "Trajetoria Espacial 3D": st.session_state.obs_3d
                        }
                        
                        graficos_form = {
                            "Análise de Distância Horizontal e Vertical": fig_dist,
                            "Comparativo Dinâmico de Motores": fig_comp
                        }
                        
                        st.divider()
                        docx_bytes_form_edit = gerar_docx_formatura_editado(lider_editado, ala_editado, obs_formatura_atual, graficos_form)
                        st.download_button(label="📥 EXPORTAR DOSSIÊ DE FORMATURA (.DOCX)", data=docx_bytes_form_edit, file_name="Dossie_Formatura.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

    # ---------------------------------------------------------
    # FLUXO 2 E 3: VOO SOLO & COMPARATIVO DE MÁQUINA (TREND MONITORING)
    # ---------------------------------------------------------
    else:
        if len(arquivos_up) > 2:
            st.warning("⚠️ Carregue apenas 1 ficheiro para Voo Solo, ou exatamente 2 ficheiros para Comparativo de Degradação (Trend Monitoring).")
        
        # --- FLUXO COMPARATIVO DE MÁQUINA ---
        elif len(arquivos_up) == 2:
            st.markdown("### 🔄 Auditoria de Trend Monitoring (Comparativo de Performance)")
            
            nome_arqA = arquivos_up[0].name
            nome_arqB = arquivos_up[1].name
            
            colA, colB = st.columns(2)
            with colA: arq_A = st.selectbox("🟠 Voo A (Baseline / Referência):", [nome_arqA, nome_arqB], index=0)
            with colB: 
                idx_padrao_b = 1 if arq_A == nome_arqA else 0
                arq_B = st.selectbox("🔵 Voo B (Recente / Inspecionado):", [nome_arqA, nome_arqB], index=idx_padrao_b)

            if arq_A == arq_B:
                st.error("⚠️ Você selecionou o mesmo ficheiro para Voo A e Voo B.")
            else:
                idx_A = 0 if arq_A == nome_arqA else 1
                idx_B = 1 if idx_A == 0 else 0
                
                dfA = processar_dados(arquivos_up[idx_A].getvalue())
                dfB = processar_dados(arquivos_up[idx_B].getvalue())
                
                dfA['Minutos'] = calcular_minutos_voo(dfA)
                dfB['Minutos'] = calcular_minutos_voo(dfB)

                st.sidebar.markdown("---")
                st.sidebar.success("✅ **Trend Monitoring Ativo**")
                
                aba_resumo_comp, aba_graf_comp, aba_rel_comp = st.tabs(["📊 1. MATRIZ DE VARIAÇÃO (DELTAS)", "📈 2. CURVAS SOBREPOSTAS", "📄 3. RELATÓRIO DE TENDÊNCIA"])

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
                    add_metric(f"Torque Máx Motor {i} (Lbs)", f"E{i} Torq ft lbs")
                    add_metric(f"Temp T5/ITT Máx Motor {i} (ºC)", f"E{i} ITT")
                    add_metric(f"Ng Máx Motor {i} (%)", f"E{i} NG")
                    add_metric(f"Nh Máx Motor {i} (%)", f"E{i} NP %")
                    add_metric(f"Temp Óleo Máx Motor {i} (ºC)", f"E{i} OilT_Celsius")

                fig_trend = None

                with aba_resumo_comp:
                    st.markdown("Variações positivas acentuadas (+) em **T5 (ITT)** ou **Ng** para uma mesma faixa de Torque indicam perda de eficiência compressora ou degradação na Seção Quente (*Hot Section*).")
                    render_html_table(df_metricas, height=360)
                    
                    st.divider()
                    st.markdown("📝 **Parecer Técnico - Análise de Deltas**")
                    obs_resumo_tmp = st.text_area("Escreva aqui:", value=st.session_state.obs_resumo, key='k_comp_resumo')
                    if st.button("💾 Salvar Comentário", key="b_comp_res"): 
                        st.session_state.obs_resumo = obs_resumo_tmp
                        st.success("Salvo!")

                with aba_graf_comp:
                    fig_trend = make_subplots(rows=4, cols=1, shared_xaxes=True, vertical_spacing=0.04, subplot_titles=("1. Perfil de Altitude MSL (ft)", "2. Torque do Motor 1 (Lbs)", "3. Temperatura T5 / ITT Motor 1 (ºC)", "4. Rotação Ng Motor 1 (%)"))
                    
                    corA = '#f97316'
                    corB = '#38bdf8'

                    if 'AltMSL' in dfA.columns: fig_trend.add_trace(go.Scattergl(x=dfA['Minutos'], y=dfA['AltMSL'], name="Alt Voo A", line=dict(color=corA, width=2, dash='dot')), row=1, col=1)
                    if 'AltMSL' in dfB.columns: fig_trend.add_trace(go.Scattergl(x=dfB['Minutos'], y=dfB['AltMSL'], name="Alt Voo B", line=dict(color=corB, width=2)), row=1, col=1)
                    if 'E1 Torq ft lbs' in dfA.columns: fig_trend.add_trace(go.Scattergl(x=dfA['Minutos'], y=dfA['E1 Torq ft lbs'], name="Tq Voo A", line=dict(color=corA, width=2)), row=2, col=1)
                    if 'E1 Torq ft lbs' in dfB.columns: fig_trend.add_trace(go.Scattergl(x=dfB['Minutos'], y=dfB['E1 Torq ft lbs'], name="Tq Voo B", line=dict(color=corB, width=2)), row=2, col=1)
                    if 'E1 ITT' in dfA.columns: fig_trend.add_trace(go.Scattergl(x=dfA['Minutos'], y=dfA['E1 ITT'], name="T5 Voo A", line=dict(color=corA, width=2)), row=3, col=1)
                    if 'E1 ITT' in dfB.columns: fig_trend.add_trace(go.Scattergl(x=dfB['Minutos'], y=dfB['E1 ITT'], name="T5 Voo B", line=dict(color=corB, width=2)), row=3, col=1)
                    if 'E1 NG' in dfA.columns: fig_trend.add_trace(go.Scattergl(x=dfA['Minutos'], y=dfA['E1 NG'], name="Ng Voo A", line=dict(color=corA, width=2)), row=4, col=1)
                    if 'E1 NG' in dfB.columns: fig_trend.add_trace(go.Scattergl(x=dfB['Minutos'], y=dfB['E1 NG'], name="Ng Voo B", line=dict(color=corB, width=2)), row=4, col=1)

                    fig_trend.update_layout(template="plotly_dark", paper_bgcolor="#090d16", plot_bgcolor="#090d16", height=900, hovermode="x unified", xaxis4_title="Tempo de Voo Sincronizado (Minutos)", margin=dict(l=15, r=15, t=35, b=15))
                    st.plotly_chart(fig_trend, use_container_width=True)

                    st.divider()
                    st.markdown("📝 **Parecer Técnico - Comportamento Dinâmico**")
                    obs_graficos_tmp = st.text_area("Escreva aqui:", value=st.session_state.obs_graficos, key='k_comp_graficos')
                    if st.button("💾 Salvar Comentário", key="b_comp_graf"): 
                        st.session_state.obs_graficos = obs_graficos_tmp
                        st.success("Salvo!")

                with aba_rel_comp:
                    obs_comp_atual = {
                        "Resumo de Degradacao e Picos (Deltas)": st.session_state.obs_resumo,
                        "Analise Comportamental Grafica": st.session_state.obs_graficos
                    }
                    
                    graficos_trend = {
                        "Comparativo de Performance (Voo A vs Voo B)": fig_trend
                    }
                    
                    docx_bytes_comp = gerar_docx_comparativo_editado(arq_A, arq_B, df_metricas, obs_comp_atual, graficos_trend)
                    st.download_button(
                        label="📥 EXPORTAR LAUDO DE TREND MONITORING (.DOCX)", 
                        data=docx_bytes_comp, 
                        file_name="Comparativo_Degradacao.docx", 
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
                        use_container_width=True
                    )


        # --- FLUXO VOO SOLO ORIGINAL (COM REGRAS OFICIAIS "BOA CEPA!" T-27M) ---
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
                add_alerta_contexto('NormAc', max_g, f"🔴 LINHA VERMELHA: Força G Positiva ({max_g:.2f}G | Limite: +{l_g_max:.1f}G)")
            if l_g_min is not None and min_g < l_g_min:
                add_alerta_contexto('NormAc', min_g, f"🔴 LINHA VERMELHA: Força G Negativa ({min_g:.2f}G | Limite: {l_g_min:.1f}G)")
            
            if projeto_atual == "T-27M (Oficial)":
                max_torq = df['E1 Torq ft lbs'].max() if 'E1 Torq ft lbs' in df.columns else 0
                max_itt = df['E1 ITT'].max() if 'E1 ITT' in df.columns else 0
                max_ng = df['E1 NG'].max() if 'E1 NG' in df.columns else 0
                max_nh = df['E1 NP %'].max() if 'E1 NP %' in df.columns else 0

                # 1. VOO INVERTIDO NA INSTRUÇÃO (Máx 15 segundos contínuos com G negativo)
                if 'NormAc' in df.columns:
                    em_invertido = df['NormAc'] < -0.15
                    seg_invertido = int(em_invertido.groupby((~em_invertido).cumsum()).sum().max()) if not em_invertido.empty else 0
                    if seg_invertido > 15:
                        add_alerta_contexto('NormAc', min_g, f"🔴 EXCEDÊNCIA VOO INVERTIDO (INSTRUÇÃO): {seg_invertido}s contínuos (Limite BOA CEPA: 15s)")

                # 2. TORQUE (Máx Decolagem: 1795 lb.ft)
                if verificar_temporizador(df, 'E1 Torq ft lbs', 2100, 2):
                    add_alerta_contexto('E1 Torq ft lbs', max_torq, f"🔴 OVER-TORQUE CRÍTICO: >2100 lb.ft por 2s ({max_torq:.0f} lb.ft)")
                elif verificar_temporizador(df, 'E1 Torq ft lbs', 1970, 1):
                    add_alerta_contexto('E1 Torq ft lbs', max_torq, f"🔴 OVER-TORQUE: Passou de 1970 lb.ft ({max_torq:.0f} lb.ft)")
                elif l_tq_red is not None and max_torq > l_tq_red:
                    add_alerta_contexto('E1 Torq ft lbs', max_torq, f"🔴 LINHA VERMELHA Torque Decolagem (>1795 lb.ft): {max_torq:.0f} lb.ft")
                
                # 3. GERADOR DE GÁS - Ng (Máx Decolagem/Subida/Cruzeiro: 101,5%)
                if verificar_temporizador(df, 'E1 NG', 102.6, 6):
                    add_alerta_contexto('E1 NG', max_ng, f"🔴 NG CRÍTICO: >102.6% por 6 seg ({max_ng:.1f}%)")
                elif l_ng_red is not None and max_ng > l_ng_red:
                    add_alerta_contexto('E1 NG', max_ng, f"🔴 LINHA VERMELHA Ng (>101.5%): {max_ng:.1f}%")
                elif l_ng_yel is not None and max_ng > l_ng_yel:
                    add_alerta_contexto('E1 NG', max_ng, f"🟡 ARCO AMARELO Ng (>100.0%): {max_ng:.1f}%")
                
                # 4. TEMPERATURA T5 / ITT (Partida: 1090°C/2s | Aceleração: 850°C/2s | Dec: 790°C | Subida: 765°C | Cruzeiro: 740°C)
                if verificar_temporizador(df, 'E1 ITT', 1090, 2):
                    add_alerta_contexto('E1 ITT', max_itt, f"🔴 T5 CRÍTICA NA PARTIDA: >1090ºC por 2 seg ({max_itt:.0f}ºC)")
                elif verificar_temporizador(df, 'E1 ITT', 850, 2):
                    add_alerta_contexto('E1 ITT', max_itt, f"🔴 T5 MÁX EM ACELERAÇÃO EXCEDIDA: >850ºC por 2 seg ({max_itt:.0f}ºC)")
                elif verificar_temporizador(df, 'E1 ITT', 790, 1):
                    add_alerta_contexto('E1 ITT', max_itt, f"🔴 LINHA VERMELHA T5 DECOLAGEM (>790ºC): {max_itt:.0f}ºC")
                elif verificar_temporizador(df, 'E1 ITT', 765, 5):
                    add_alerta_contexto('E1 ITT', max_itt, f"🟡 ALERTA T5 MÁX SUBIDA (>765ºC): {max_itt:.0f}ºC")
                elif l_itt_yel is not None and max_itt > l_itt_yel:
                    add_alerta_contexto('E1 ITT', max_itt, f"🟡 ARCO AMARELO T5 CRUZEIRO (>740ºC): {max_itt:.0f}ºC")
                
                # 5. ROTAÇÃO DA HÉLICE - Nh (Máx Decolagem/Subida/Cruzeiro: 100%)
                if verificar_temporizador(df, 'E1 NP %', 110, 1):
                    add_alerta_contexto('E1 NP %', max_nh, f"🔴 OVERSPEED HÉLICE CRÍTICO: >110% ({max_nh:.1f}%)")
                elif verificar_temporizador(df, 'E1 NP %', 106, 1):
                    add_alerta_contexto('E1 NP %', max_nh, f"🔴 OVERSPEED HÉLICE: >106% ({max_nh:.1f}%)")
                elif l_nh_red is not None and max_nh > l_nh_red:
                    add_alerta_contexto('E1 NP %', max_nh, f"🔴 LINHA VERMELHA Nh (>100.0%): {max_nh:.1f}%")

                # 6. LUZ PRESSÃO DO ÓLEO (Abaixo de 40 PSI com motor em funcionamento)
                if 'E1 OilP' in df.columns and 'E1 NG' in df.columns:
                    df_mot_ligado = df[df['E1 NG'] > 50]
                    if not df_mot_ligado.empty and (df_mot_ligado['E1 OilP'] < 40).any():
                        min_oilp_voo = df_mot_ligado['E1 OilP'].min()
                        add_alerta_contexto('E1 OilP', min_oilp_voo, f"🔴 LUZ PRESS ÓLEO ATIVADA: Pressão caiu para {min_oilp_voo:.0f} PSI (Limite mín: 40 PSI)")
            else:
                for i_mot in range(1, qtd_motores_config + 1):
                    m_torq = df[f'E{i_mot} Torq ft lbs'].max() if f'E{i_mot} Torq ft lbs' in df.columns else 0
                    m_itt = df[f'E{i_mot} ITT'].max() if f'E{i_mot} ITT' in df.columns else 0
                    m_ng = df[f'E{i_mot} NG'].max() if f'E{i_mot} NG' in df.columns else 0
                    m_nh = df[f'E{i_mot} NP %'].max() if f'E{i_mot} NP %' in df.columns else 0
                    
                    if l_tq_red is not None and m_torq >= l_tq_red: add_alerta_contexto(f'E{i_mot} Torq ft lbs', m_torq, f"🔴 Motor {i_mot} LINHA VERMELHA Torque: {m_torq:.0f} Lbs")
                    elif l_tq_yel is not None and m_torq >= l_tq_yel: add_alerta_contexto(f'E{i_mot} Torq ft lbs', m_torq, f"🟡 Motor {i_mot} ARCO AMARELO Torque: {m_torq:.0f} Lbs")
                    if l_ng_red is not None and m_ng >= l_ng_red: add_alerta_contexto(f'E{i_mot} NG', m_ng, f"🔴 Motor {i_mot} LINHA VERMELHA Ng: {m_ng:.1f}%")
                    elif l_ng_yel is not None and m_ng >= l_ng_yel: add_alerta_contexto(f'E{i_mot} NG', m_ng, f"🟡 Motor {i_mot} ARCO AMARELO Ng: {m_ng:.1f}%")
                    if l_itt_red is not None and m_itt >= l_itt_red: add_alerta_contexto(f'E{i_mot} ITT', m_itt, f"🔴 Motor {i_mot} LINHA VERMELHA ITT: {m_itt:.0f} ºC")
                    elif l_itt_yel is not None and m_itt >= l_itt_yel: add_alerta_contexto(f'E{i_mot} ITT', m_itt, f"🟡 Motor {i_mot} ARCO AMARELO ITT: {m_itt:.0f} ºC")
                    if l_nh_red is not None and m_nh >= l_nh_red: add_alerta_contexto(f'E{i_mot} NP %', m_nh, f"🔴 Motor {i_mot} LINHA VERMELHA Nh: {m_nh:.1f}%")
                    elif l_nh_yel is not None and m_nh >= l_nh_yel: add_alerta_contexto(f'E{i_mot} NP %', m_nh, f"🟡 Motor {i_mot} ARCO AMARELO Nh: {m_nh:.1f}%")

            for i_mot in range(1, qtd_motores_config + 1):
                m_oilt = df[f'E{i_mot} OilT_Celsius'].max() if f'E{i_mot} OilT_Celsius' in df.columns else 0
                m_oilp = df[f'E{i_mot} OilP'].max() if f'E{i_mot} OilP' in df.columns else 0
                if l_oilt_red is not None and m_oilt >= l_oilt_red: add_alerta_contexto(f'E{i_mot} OilT_Celsius', m_oilt, f"🔴 Motor {i_mot} LINHA VERMELHA Temp Óleo: {m_oilt:.0f} ºC")
                if l_oilp_red is not None and m_oilp >= l_oilp_red: add_alerta_contexto(f'E{i_mot} OilP', m_oilp, f"🔴 Motor {i_mot} LINHA VERMELHA Pressão Óleo: {m_oilp:.0f} PSI")

            tempos = {'acionamento': '-', 'decolagem': '-', 'pouso': '-', 'corte': '-', 't_voo': '-'}
            if tipo_projeto == "Asa Rotativa":
                motor_on = df[(df.get('E1 NG', pd.Series([0])) > 15) & (df['Datetime'].notna())]
                if not motor_on.empty:
                    tempos['acionamento'] = motor_on.iloc[0]['Datetime'].strftime('%H:%M:%S')
                    tempos['corte'] = motor_on.iloc[-1]['Datetime'].strftime('%H:%M:%S')
                    tempos['decolagem'] = tempos['acionamento']
                    tempos['pouso'] = tempos['corte']
                    tempos['t_voo'] = str(motor_on.iloc[-1]['Datetime'] - motor_on.iloc[0]['Datetime']).split('.')[0]
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
                        tempos['t_voo'] = str(voo_ativo.iloc[-1]['Datetime'] - voo_ativo.iloc[0]['Datetime']).split('.')[0]
                else:
                    voo_ativo = pd.DataFrame()

            cor_status_missao = "#ef4444" if any("🔴" in a for a in alertas) else ("#f59e0b" if len(alertas) > 0 else "#10b981")
            txt_status_missao = f"{len(alertas)} EXCEDÊNCIA(S) DETECTADA(S)" if len(alertas) > 0 else "OPERAÇÃO NORMAL (ARCO VERDE)"

            st.markdown(f"""
            <div style="background: #090d16; border: 1px solid #1e293b; border-left: 4px solid {cor_status_missao}; border-radius: 10px; padding: 14px 20px; margin-bottom: 20px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 15px;">
                <div style="display: flex; gap: 28px; flex-wrap: wrap;">
                    <div>
                        <div style="font-size: 10px; color: #64748b; text-transform: uppercase;">Acionamento (Start)</div>
                        <div style="font-size: 15px; font-weight: 700; color: #f8fafc; font-family: 'JetBrains Mono', monospace;">{tempos['acionamento']}</div>
                    </div>
                    <div>
                        <div style="font-size: 10px; color: #64748b; text-transform: uppercase;">Decolagem (T/O)</div>
                        <div style="font-size: 15px; font-weight: 700; color: #38bdf8; font-family: 'JetBrains Mono', monospace;">{tempos['decolagem']}</div>
                    </div>
                    <div>
                        <div style="font-size: 10px; color: #64748b; text-transform: uppercase;">Pouso (LDG)</div>
                        <div style="font-size: 15px; font-weight: 700; color: #38bdf8; font-family: 'JetBrains Mono', monospace;">{tempos['pouso']}</div>
                    </div>
                    <div>
                        <div style="font-size: 10px; color: #64748b; text-transform: uppercase;">Corte (Shutdown)</div>
                        <div style="font-size: 15px; font-weight: 700; color: #f8fafc; font-family: 'JetBrains Mono', monospace;">{tempos['corte']}</div>
                    </div>
                    <div>
                        <div style="font-size: 10px; color: #64748b; text-transform: uppercase;">Tempo de Voo</div>
                        <div style="font-size: 15px; font-weight: 700; color: #10b981; font-family: 'JetBrains Mono', monospace;">{tempos['t_voo']}</div>
                    </div>
                </div>
                <div style="background: {cor_status_missao}22; border: 1px solid {cor_status_missao}66; color: {cor_status_missao}; padding: 6px 14px; border-radius: 20px; font-size: 11.5px; font-weight: 700; letter-spacing: 0.6px;">
                    ● {txt_status_missao}
                </div>
            </div>
            """, unsafe_allow_html=True)

            aba_didatica, aba_inspetoria, aba_manobras, aba_relatorio = st.tabs([
                "🎛️ 1. PAINEL DE INSTRUMENTOS (EIS)", 
                "📈 2. INSPETORIA E TELEMETRIA", 
                "🛩️ 3. ENVELOPE DE MANOBRAS E POUSO", 
                "📄 4. LAUDO TÉCNICO E IA"
            ])

            fig_dash = None

            with aba_didatica:
                st.markdown("#### 🎯 Leituras de Pico da Missão (Engine & Airframe Limits)")
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
                    with st.expander("📋 Consultar Gabarito Oficial de Limites Operacionais T-27M ('BOA CEPA!')"):
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
                st.markdown("#### 🗺️ Navegação e Reconstituição de Rota (GPS / Google Earth)")
                df_mapa = df[(df['IAS'] > 30)].dropna(subset=['Latitude', 'Longitude'])
                if not df_mapa.empty:
                    fig_map_temp = go.Figure()
                    fig_map_temp.add_trace(go.Scattermap(lat=df_mapa['Latitude'], lon=df_mapa['Longitude'], mode='lines', line=dict(width=4, color='#38bdf8'), name="Rota"))
                    fig_map_temp.add_trace(go.Scattermap(lat=[df_mapa['Latitude'].iloc[0]], lon=[df_mapa['Longitude'].iloc[0]], mode='markers', marker=dict(size=13, color='#10b981'), name="Decolagem"))
                    fig_map_temp.add_trace(go.Scattermap(lat=[df_mapa['Latitude'].iloc[-1]], lon=[df_mapa['Longitude'].iloc[-1]], mode='markers', marker=dict(size=13, color='#ef4444'), name="Corte"))
                    fig_map_temp.update_layout(map_style="carto-darkmatter", map_zoom=8, map_center={"lat": df_mapa['Latitude'].mean(), "lon": df_mapa['Longitude'].mean()}, margin={"r":0,"t":0,"l":0,"b":0}, height=420, showlegend=False)
                    st.plotly_chart(fig_map_temp, use_container_width=True)
                    
                    kml = simplekml.Kml()
                    linestring = kml.newlinestring(name="Trajetoria de Voo 3D", coords=[(r['Longitude'], r['Latitude'], r['AltMSL'] * 0.3048) for i, r in df_mapa.iterrows()])
                    linestring.altitudemode = simplekml.AltitudeMode.absolute
                    linestring.style.linestyle.color = simplekml.Color.cyan
                    linestring.style.linestyle.width = 4
                    st.download_button("🌍 Exportar Telemetria 3D para Google Earth (.KML)", data=kml.kml(), file_name="rota_voo.kml", mime="application/vnd.google-earth.kml+xml")

                st.divider()
                st.markdown("📝 **Apontamentos Técnicos - Painel e Navegação**")
                obs_res_tmp = st.text_area("Escreva aqui:", value=st.session_state.obs_resumo, key='k_res')
                if st.button("💾 Salvar Comentário", key="b_res"): 
                    st.session_state.obs_resumo = obs_res_tmp
                    st.success("Salvo com sucesso!")

            with aba_inspetoria:
                st.markdown("#### 🚨 Auditoria Automática de Excedências (Arcos Amarelo e Vermelho)")
                if len(alertas) > 0:
                    for a in alertas:
                        if "LINHA VERMELHA" in a or "CRÍTICO" in a or "EXCEDÊNCIA" in a or "ATIVADA" in a: st.error(a)
                        else: st.warning(a)
                else:
                    st.success("✅ Nenhuma violação de limite operacional registrada durante a missão.")
                    
                st.divider()
                st.markdown("#### 🔍 Filtro de Investigação de Eventos (Data Query)")
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
                    else: st.info("Nenhum registo encontrado para o filtro selecionado.")

                st.divider()
                st.markdown("#### 📈 Gráficos Sincronizados de Dinâmica e Propulsão")
                fig_dash = make_subplots(rows=6, cols=1, shared_xaxes=True, vertical_spacing=0.03, specs=[[{"secondary_y": True}], [{"secondary_y": False}], [{"secondary_y": False}], [{"secondary_y": False}], [{"secondary_y": False}], [{"secondary_y": False}]], subplot_titles=("1. Altitude MSL (ft) e Velocidade IAS (kt)", "2. Fator de Carga Normal (Força G)", "3. Torque do Motor (lb.ft)", "4. Temperatura Inter-Turbinas - T5 / ITT (ºC)", "5. Rotação do Gerador de Gás - Ng (%)", "6. Rotação da Hélice - Nh (%)"))
                cores_motores = ['#10b981', '#f59e0b', '#f97316', '#8b5cf6']
                cores_itt = ['#f59e0b', '#ef4444', '#ea580c', '#dc2626']
                cores_ng = ['#a855f7', '#38bdf8', '#14b8a6', '#f97316']
                cores_nh = ['#eab308', '#22c55e', '#64748b', '#cbd5e1']
                
                if 'AltMSL' in df.columns: fig_dash.add_trace(go.Scattergl(x=df['Datetime'], y=df['AltMSL'], name="Altitude (ft)", fill='tozeroy', fillcolor='rgba(59, 130, 246, 0.1)', line=dict(color='#3b82f6', width=2)), row=1, col=1, secondary_y=False)
                if 'IAS' in df.columns: fig_dash.add_trace(go.Scattergl(x=df['Datetime'], y=df['IAS'], name="Velocidade (kt)", line=dict(color='#06b6d4', width=2)), row=1, col=1, secondary_y=True)
                if 'NormAc' in df.columns: fig_dash.add_trace(go.Scattergl(x=df['Datetime'], y=df['NormAc'], name="Força G", fill='tozeroy', fillcolor='rgba(239, 68, 68, 0.1)', line=dict(color='#ef4444', width=2)), row=2, col=1)
                if l_g_max is not None: fig_dash.add_hline(y=l_g_max, line_dash="dash", line_color="#ef4444", row=2, col=1)
                        
                for i_m in range(1, qtd_motores_config + 1):
                    if f'E{i_m} Torq ft lbs' in df.columns: fig_dash.add_trace(go.Scattergl(x=df['Datetime'], y=df[f'E{i_m} Torq ft lbs'], name=f"Tq M{i_m}", line=dict(color=cores_motores[i_m-1], width=2)), row=3, col=1)
                    if f'E{i_m} ITT' in df.columns: fig_dash.add_trace(go.Scattergl(x=df['Datetime'], y=df[f'E{i_m} ITT'], name=f"T5 M{i_m}", line=dict(color=cores_itt[i_m-1], width=2)), row=4, col=1)
                    if f'E{i_m} NG' in df.columns: fig_dash.add_trace(go.Scattergl(x=df['Datetime'], y=df[f'E{i_m} NG'], name=f"Ng M{i_m} (%)", line=dict(color=cores_ng[i_m-1], width=2)), row=5, col=1)
                    if f'E{i_m} NP %' in df.columns: fig_dash.add_trace(go.Scattergl(x=df['Datetime'], y=df[f'E{i_m} NP %'], name=f"Nh M{i_m} (%)", line=dict(color=cores_nh[i_m-1], width=2)), row=6, col=1)
                
                if l_tq_red is not None: fig_dash.add_hline(y=l_tq_red, line_dash="dash", line_color="#ef4444", row=3, col=1)
                if l_itt_red is not None: fig_dash.add_hline(y=l_itt_red, line_dash="dash", line_color="#ef4444", row=4, col=1)
                if l_ng_red is not None: fig_dash.add_hline(y=l_ng_red, line_dash="dash", line_color="#ef4444", row=5, col=1)
                if l_nh_red is not None: fig_dash.add_hline(y=l_nh_red, line_dash="dash", line_color="#ef4444", row=6, col=1)
                
                fig_dash.update_layout(template="plotly_dark", paper_bgcolor="#090d16", plot_bgcolor="#090d16", height=1200, hovermode="x unified", margin=dict(l=15, r=15, t=35, b=15), legend=dict(title="Sensores", orientation="v", yanchor="top", y=1, xanchor="left", x=1.02, bgcolor="rgba(9,13,22,0.8)"))
                st.plotly_chart(fig_dash, use_container_width=True)

                st.divider()
                st.markdown("📝 **Observações Manuais de Inspetoria**")
                obs_graf_tmp_manual = st.text_area("Escreva o seu parecer manual aqui:", value=st.session_state.obs_graficos, key='k_graf_manual')
                if st.button("💾 Salvar Comentário Manual", key="b_graf_manual"): 
                    st.session_state.obs_graficos = obs_graf_tmp_manual
                    st.success("Salvo com sucesso!")

            with aba_manobras:
                c_m1, c_m2 = st.columns(2)
                with c_m1:
                    st.markdown("#### 🛬 Diagnóstico de Aproximação e Toque (Landing)")
                    if 'IAS' in df.columns and 'AltMSL' in df.columns and not voo_ativo.empty:
                        fase_final = voo_ativo.tail(300) 
                        pico_g_pouso = fase_final['NormAc'].max()
                        vel_aprox = fase_final['IAS'].mean()
                        st.info(f"**Velocidade Média na Final:** `{vel_aprox:.1f} kt` (Limite Flape Pouso T-27M: `148 kt`)")
                        if pico_g_pouso > 2.0: st.error(f"🔴 **Pouso Brusco / Excedência Flape Baixado (>2.0G):** Pico de `{pico_g_pouso:.2f} G` registrado.")
                        elif pico_g_pouso > 1.5: st.warning(f"🟡 **Pouso Firme:** Pico de `{pico_g_pouso:.2f} G` na pista.")
                        else: st.success(f"✅ **Pouso Suave:** Toque estabilizado em `{pico_g_pouso:.2f} G`.")
                with c_m2:
                    st.markdown("#### 🤸 Detetor de Manobras Acrobáticas / Atitudes Extremas")
                    if 'Pitch' in df.columns and 'Roll' in df.columns:
                        df_acro = df[(df['Pitch'] > 45) | (df['Pitch'] < -45) | (df['Roll'] > 60) | (df['Roll'] < -60)]
                        if not df_acro.empty:
                            st.markdown(f"**Registos de elevada atitude (Pitch > 45º ou Roll > 60º):** `{len(df_acro)}` amostras.")
                            colunas_acro = ['Lcl Time', 'AltMSL', 'IAS', 'Pitch', 'Roll', 'NormAc']
                            for j in range(1, qtd_motores_config + 1): colunas_acro.extend([f'E{j} Torq ft lbs', f'E{j} ITT', f'E{j} NG', f'E{j} NP %'])
                            colunas_acro = [c for c in colunas_acro if c in df.columns]
                            exibir_tabela_grande(df_acro.sort_values(by='NormAc', ascending=False).head(20)[colunas_acro].copy())
                        else: st.info("Voo nivelado / sem manobras de grande inclinação detectadas.")
                            
                st.divider()
                st.markdown("#### ✈️ Perfil Espacial 3D Colorido por Velocidade (IAS)")
                if not voo_ativo.empty and 'Latitude' in voo_ativo.columns and 'Longitude' in voo_ativo.columns:
                    df_3d = voo_ativo.dropna(subset=['Latitude', 'Longitude', 'AltMSL'])
                    fig_3d_temp = go.Figure(data=[go.Scatter3d(x=df_3d['Longitude'], y=df_3d['Latitude'], z=df_3d['AltMSL'], mode='lines', line=dict(color=df_3d['IAS'], colorscale='Turbo', width=4, colorbar=dict(title="IAS (kt)", len=0.5, x=1.05)))])
                    fig_3d_temp.update_layout(template="plotly_dark", paper_bgcolor="#090d16", height=600, margin=dict(l=0, r=0, b=0, t=20), scene=dict(aspectmode='auto'))
                    st.plotly_chart(fig_3d_temp, use_container_width=True)

                st.divider()
                st.markdown("📝 **Observações - Envelope de Manobras e Pouso**")
                obs_manob_tmp = st.text_area("Escreva aqui:", value=st.session_state.obs_manobras, key='k_manob')
                if st.button("💾 Salvar Comentário", key="b_manob"): 
                    st.session_state.obs_manobras = obs_manob_tmp
                    st.success("Salvo com sucesso!")

            with aba_relatorio:
                st.markdown("#### 📝 Consolidação do Relatório Técnico de Telemetria (RTT)")
        
                col_t1, col_t2, col_t3, col_t4, col_t5 = st.columns(5)
                with col_t1: temp_ac = st.text_input("Acionamento:", tempos.get('acionamento','-'))
                with col_t2: temp_dec = st.text_input("Decolagem:", tempos.get('decolagem','-'))
                with col_t3: temp_pou = st.text_input("Pouso:", tempos.get('pouso','-'))
                with col_t4: temp_cor = st.text_input("Corte:", tempos.get('corte','-'))
                with col_t5: temp_voo = st.text_input("Tempo de Voo:", tempos.get('t_voo','-'))
                tempos_editados = {'acionamento': temp_ac, 'decolagem': temp_dec, 'pouso': temp_pou, 'corte': temp_cor, 't_voo': temp_voo}

                alertas_txt = st.text_area("Registo de Alertas e Excedências (Injetado na Tabela 1 do Word):", value="\n".join(alertas), height=110)
                
                st.divider()
                st.markdown("""
                <div style="background: linear-gradient(135deg, rgba(29, 78, 216, 0.15) 0%, rgba(15, 23, 42, 0.8) 100%); border: 1px solid #1d4ed8; border-radius: 10px; padding: 16px 20px; margin-bottom: 12px;">
                    <div style="font-size: 14px; font-weight: 700; color: #60a5fa;">🤖 Inteligência Artificial de Inspetoria Aeronáutica (Padrão T-27M)</div>
                    <div style="font-size: 12.5px; color: #94a3b8; margin-top: 4px;">Gera automaticamente o Estudo de Caso e o Parecer de Manutenção confrontando os dados do voo com os limites operacionais oficiais do T-27M.</div>
                </div>
                """, unsafe_allow_html=True)

                if st.button("🪄 Gerar Parecer Técnico Oficial com IA", key="btn_ia_solo_rel", use_container_width=True):
                    with st.spinner("Analisando telemetria e redigindo parecer técnico no padrão FAB..."):
                        try:
                            pico_tq = df['E1 Torq ft lbs'].max() if 'E1 Torq ft lbs' in df.columns else 0
                            pico_itt = df['E1 ITT'].max() if 'E1 ITT' in df.columns else 0
                            pico_ng = df['E1 NG'].max() if 'E1 NG' in df.columns else 0
                            pico_nh = df['E1 NP %'].max() if 'E1 NP %' in df.columns else 0

                            texto_alertas = ", ".join(alertas) if alertas else "Nenhum alerta. Voo dentro dos limites normais."
                            
                            prompt = f"""
                            Atue como um Inspetor Sênior de Manutenção e Avaliador de Telemetria da Força Aérea Brasileira (especialista na aeronave T-27M Tucano e motor Pratt & Whitney PT6A-25C).
                            O seu objetivo é emitir um Parecer Técnico Oficial e um Estudo de Caso aprofundado, baseado estritamente nos dados deste voo e nos limites oficiais da aeronave.

                            LIMITES OPERACIONAIS OFICIAIS DO T-27M (REFERÊNCIA):
                            - Torque Máx Decolagem: 1795 lb.ft
                            - Temperatura T5 (ITT) Máx: Partida 1090°C (2s) | Aceleração 850°C (2s) | Decolagem 790°C | Subida 765°C | Cruzeiro 740°C
                            - Ng Máx (Decolagem/Subida/Cruzeiro): 101,5%
                            - Nh Máx (Decolagem/Subida/Cruzeiro): 100%
                            - Voo Invertido na Instrução: Máximo 15 segundos
                            - Limites de Aceleração (Força G): -3G a +6G (flapes recolhidos) e 0G a +2G (flapes baixados)
                            - Luz Pressão do Óleo: Abaixo de 40 psi

                            DADOS DE PICO COLETADOS NESTA MISSÃO:
                            - Força G: Máxima {max_g:.2f}G | Mínima {min_g:.2f}G
                            - Motor (Torque Máx): {pico_tq:.0f} lb.ft
                            - Motor (Temperatura T5 / ITT Máx): {pico_itt:.0f} ºC
                            - Motor (Rotação Ng Máx): {pico_ng:.1f}%
                            - Motor (Hélice Nh Máx): {pico_nh:.1f}%
                            - Ocorrências e Alertas de Ultrapassagem: {texto_alertas}

                            DIRETRIZES DE FORMATAÇÃO (ESTRUTURA VISUAL OBRIGATÓRIA):
                            - PROIBIDO escrever texto corrido tipo dissertação. O laudo deve ser altamente escaneável e estruturado.
                            - Use Markdown para criar hierarquia visual: utilize `#` para o Título Principal e `##` para Subtítulos.
                            - Use listas de marcadores (bullet points) extensivamente para listar eventos, facilitando a leitura rápida.
                            - Destaque em **negrito** todos os valores numéricos, nomes de sensores (T5/ITT, Ng, Nh, Torque) e peças críticas.
                            - Utilize Emojis como código de cores e status: 🔴 (Crítico/Limites Excedidos), 🟡 (Alerta/Monitoramento), 🟢 (Parâmetros Normais), 📊 (Referência a Gráfico), 🛠️ (Ação de Manutenção).

                            DIRETRIZES DE REDAÇÃO E CONTEÚDO (ESTUDO DE CASO):
                            1. INTEGRAÇÃO VISUAL COM GRÁFICOS: Guie o leitor. Comece os tópicos fazendo referência direta ao painel (ex: "📊 **Análise do Gráfico 1 e 2 (Dinâmica):** Como evidenciado pela curva...").
                            2. FOCO CIRÚRGICO NAS OCORRÊNCIAS (🔴 e 🟡): Se houver alertas, crie um subtítulo específico para o Estudo de Caso. Quebre a análise da falha em 3 pontos rápidos:
                               - **Momento do Evento:** (Cruze o alerta com Altitude, Velocidade e Hora).
                               - **Dinâmica Suspeita:** (Que tipo de manobra ou erro operacional causou o pico).
                               - **Consequência Física:** (Ex: fadiga estrutural severa, degradação térmica da seção quente).
                            3. CONDIÇÃO VERDE (🟢): Mesmo se não houver alertas, comente sobre o comportamento saudável e os picos atingidos perante os limites oficiais do T-27M.
                            4. RECOMENDAÇÕES DIRETAS: Termine com uma lista de ordens de serviço (ex: Inspeção Visual, END, Boroscopia, ou "Aeronave Liberada").
                            5. VOCABULÁRIO: Jargão militar e aeronáutico estrito.

                            ESTRUTURA EXATA EXIGIDA PARA O LAUDO:
                            # RELATÓRIO DE INSPETORIA TÉCNICA E TELEMETRIA
                            ## 1. 📊 VISÃO GERAL E DINÂMICA DE VOO
                            ## 2. 🔍 ESTUDO DE CASO: OCORRÊNCIAS E PICOS DE ESFORÇO
                            ## 3. 🛠️ DIAGNÓSTICO E RECOMENDAÇÕES DE MANUTENÇÃO
                            """
                            
                            texto_ia_solo = gerar_texto_ia(prompt)
                            st.session_state.obs_graficos = texto_ia_solo
                            st.rerun()
                        except Exception as e:
                            st.error(f"Erro na IA: {e}")
                
                obs_graf_tmp_rel = st.text_area("Parecer Técnico da Inspetoria (Editável para o Documento Oficial):", value=st.session_state.obs_graficos, key='k_graf_rel', height=260)
                if st.button("💾 Confirmar Alterações no Parecer", key="b_graf_rel"):
                    st.session_state.obs_graficos = obs_graf_tmp_rel
                    st.success("Parecer atualizado e pronto para exportação!")
                
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
                    label="📥 BAIXAR RELATÓRIO TÉCNICO OFICIAL (.DOCX)", 
                    data=docx_bytes_solo_editado, 
                    file_name=f"RTT_{projeto_atual.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.docx", 
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
                    use_container_width=True
                )
