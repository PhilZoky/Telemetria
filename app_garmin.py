import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import simplekml
from docx import Document
from docx.shared import RGBColor
import io
import os
import json
from datetime import datetime
import streamlit.components.v1 as components

# ==========================================
# 0. SISTEMA DE LOGIN E CADASTRO (SaaS)
# ==========================================
import json
import os

ARQUIVO_USUARIOS = "banco_usuarios.json"

# Carrega os usuários salvos ou cria o seu usuário master inicial estruturado
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

# Tela de bloqueio e formulários
if st.session_state.usuario_logado is None:
    st.markdown("<br><br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.title("🔒 Plataforma Telemetria")
        st.caption("Sistema Comercial de Telemetria e Inspetoria de Voo")
        
        # Cria abas separadas para entrar ou criar conta
        aba_login, aba_cadastro = st.tabs(["Fazer Login", "Criar Nova Conta"])
        
        with aba_login:
            st.markdown("#### Acesse o seu painel")
            usuario_input = st.text_input("Nome de Usuário", key="log_user")
            senha_input = st.text_input("Senha", type="password", key="log_pass")
            
            if st.button("Entrar no Sistema", use_container_width=True):
                if usuario_input in usuarios_permitidos:
                    dados_usuario = usuarios_permitidos[usuario_input]
                    
                    # Compatibilidade (verifica se é o formato antigo ou o novo formato completo)
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
                        st.error("❌ Credenciais incorretas.")
                else:
                    st.error("❌ Credenciais incorretas.")
                    
        with aba_cadastro:
            st.markdown("#### Dados do Cliente")
            novo_nome = st.text_input("Nome Completo", key="cad_nome")
            novo_email = st.text_input("E-mail de Contato", key="cad_email")
            nova_org = st.text_input("Organização / Esquadrão / Empresa", key="cad_org")
            
            st.markdown("#### Dados de Acesso")
            novo_usuario = st.text_input("Escolha um Nome de Usuário (Login)", key="cad_user")
            nova_senha = st.text_input("Crie uma Senha", type="password", key="cad_pass")
            
            if st.button("Cadastrar Cliente", use_container_width=True):
                if novo_usuario in usuarios_permitidos:
                    st.warning("⚠️ Este nome de usuário já existe. Escolha outro.")
                elif not novo_usuario or not nova_senha or not novo_nome or not novo_email:
                    st.warning("⚠️ Preencha todos os campos obrigatórios (Nome, E-mail, Usuário e Senha).")
                else:
                    # Estrutura profissional de dados do cliente (Futuro CRM)
                    usuarios_permitidos[novo_usuario] = {
                        "senha": nova_senha,
                        "nome": novo_nome,
                        "email": novo_email,
                        "organizacao": nova_org,
                        "data_cadastro": datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    }
                    salvar_usuarios(usuarios_permitidos)
                    st.success(f"✅ Conta de {novo_nome} criada com sucesso! Pode ir à aba de Login para aceder.")
                    
    st.stop() # Bloqueia o carregamento do resto do site

# ==========================================
# PAINEL DO USUÁRIO LOGADO (Barra Lateral)
# ==========================================
st.sidebar.markdown(f"👤 **Piloto/Analista:** `{st.session_state.nome_logado}`")
if st.sidebar.button("Sair (Logout)"):
    st.session_state.usuario_logado = None
    st.session_state.nome_logado = None
    st.rerun()

# Ajuste automático do ficheiro de perfis para cada cliente:
ARQUIVO_PERFIS = f"perfis_{st.session_state.usuario_logado}.json"


# ==========================================
# 1. CONFIGURAÇÃO E GERENCIADOR DE PERFIS
# ==========================================
st.set_page_config(page_title="Telemetria AFA", layout="wide")

ARQUIVO_PERFIS = f"perfis_{st.session_state.usuario_logado}.json"
PERFIS_PADRAO = {
    "T-27M (Oficial)": {
        "tipo_aeronave": "Asa Fixa",
        "qtd_motores": 1,
        "g_max": 3.0, 
        "g_min": -1.5,
        "tq_red": 1795.0, 
        "tq_yel": 1795.0,
        "itt_red": 790.0, 
        "itt_yel": 740.0,
        "ng_red": 101.5, 
        "ng_yel": 100.0,
        "nh_red": 101.0, 
        "nh_yel": 100.0,
        "oilt_red": 99.0, 
        "oilt_min_green": 10.0,
        "oilp_red": 95.0, 
        "oilp_min_green": 75.0
    },
    "Multi-Motor (Padrão)": {
        "tipo_aeronave": "Asa Fixa",
        "qtd_motores": 2,
        "g_max": 2.5, 
        "g_min": 0.0,
        "tq_red": 1500.0, 
        "tq_yel": 1400.0,
        "itt_red": 850.0, 
        "itt_yel": 800.0,
        "ng_red": 100.0, 
        "ng_yel": 98.0,
        "nh_red": 100.0, 
        "nh_yel": 98.0,
        "oilt_red": 100.0, 
        "oilt_min_green": 20.0,
        "oilp_red": 90.0, 
        "oilp_min_green": 60.0
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
# 3. BARRA LATERAL (MENU E ASSINATURA)
# ==========================================
st.sidebar.markdown("### 🛠️ Telemetria")
st.sidebar.caption("👨‍✈️ **Analista Responsável:** Filipe Iack")
st.sidebar.markdown("---")

st.sidebar.header("✈️ Projeto da Aeronave")
projeto_atual = st.sidebar.selectbox("Selecione o Perfil:", list(perfis_salvos.keys()))
p_dados = perfis_salvos[projeto_atual]

tipo_projeto = p_dados.get("tipo_aeronave", "Asa Fixa")
qtd_motores_config = int(p_dados.get("qtd_motores", 1))

st.sidebar.info(f"**Classe:** {tipo_projeto} | **Motores Configurados:** {qtd_motores_config}")

st.sidebar.markdown("---")
st.sidebar.header("🕹️ Modo de Análise")
modo_analise = st.sidebar.radio(
    "Selecione o tipo de missão:", 
    ["Voo Solo / Comparativo de Máquina", "Voo em Formatura (Líder e Ala)"]
)
st.sidebar.markdown("---")

with st.sidebar.expander("➕ Adicionar Novo Projeto", expanded=False):
    novo_projeto = st.text_input("Nome da nova aeronave:")
    nova_classe = st.selectbox("Categoria ao criar:", ["Asa Fixa", "Asa Rotativa"])
    qtd_mot_novo = st.number_input("Quantidade de Motores ao criar:", min_value=1, max_value=6, value=1)
    if st.button("Criar Projeto"):
        if novo_projeto and novo_projeto not in perfis_salvos:
            perfis_salvos[novo_projeto] = {
                "tipo_aeronave": nova_classe,
                "qtd_motores": int(qtd_mot_novo),
                "g_max": 3.0, "g_min": -1.5,
                "tq_red": 1795.0, "tq_yel": 1795.0,
                "itt_red": 790.0, "itt_yel": 740.0,
                "ng_red": 101.5, "ng_yel": 100.0,
                "nh_red": 101.0, "nh_yel": 100.0,
                "oilt_red": 99.0, "oilt_min_green": 10.0,
                "oilp_red": 95.0, "oilp_min_green": 75.0
            }
            salvar_perfis(perfis_salvos)
            st.success(f"Projeto '{novo_projeto}' criado!")
            st.rerun()

with st.sidebar.expander("⚙️ Arcos do Painel (Verde/Amarelo/Vermelho)", expanded=False):
    st.caption("Configure os limites estruturais e os motores do perfil selecionado.")
    
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
    l_itt_red = ler_limite("ITT Red Line", "itt_red")
    l_itt_yel = ler_limite("ITT Início Arco Amarelo", "itt_yel")
    l_oilt_red = ler_limite("Temp Óleo Red Line", "oilt_red")
    l_oilt_min = ler_limite("Temp Óleo Mínima (Arco Verde)", "oilt_min_green")
    l_oilp_red = ler_limite("Pressão Óleo Red Line", "oilp_red")
    l_oilp_min = ler_limite("Pressão Óleo Mínima (Arco Verde)", "oilp_min_green")

    if st.button("💾 Salvar Alterações de Arcos e Motores"):
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
        st.success(f"Configurações do {projeto_atual} salvas!")
        st.rerun()


# ==========================================
# 4. FUNÇÕES BASE (MATEMÁTICA E PROCESSAMENTO)
# ==========================================
st.title("📊 Sistema de Avaliação de Voo e Inspetoria")
st.markdown("Plataforma Avançada de Telemetria Técnica e Tática.")

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
    R = 6371000 
    phi1, phi2 = np.radians(lat1), np.radians(lat2)
    dphi = np.radians(lat2 - lat1)
    dlambda = np.radians(lon2 - lon1)
    a = np.sin(dphi/2)**2 + np.cos(phi1)*np.cos(phi2)*np.sin(dlambda/2)**2
    return 2 * R * np.arctan2(np.sqrt(a), np.sqrt(1 - a))

def verificar_temporizador(df, coluna, limite_valor, segundos_necessarios):
    if coluna not in df.columns: 
        return False
    acima_limite = df[coluna] > limite_valor
    consecutivos = acima_limite.groupby((~acima_limite).cumsum()).sum()
    return consecutivos.max() >= segundos_necessarios


# ==========================================
# 5. GERADORES DE INTERFACE
# ==========================================
def exibir_card_arcos(titulo, pico, red_line, yellow_line, min_green, unidade, formato=":.1f", e_minimo=False):
    cor, status = "#2ecc71", "ARCO VERDE"
    
    if red_line is None: 
        cor, status = "#aaaaaa", "SEM LIMITE"
    else:
        if e_minimo:
            if pico < red_line: 
                cor, status = "#e74c3c", "LINHA VERMELHA"
            elif yellow_line is not None and pico <= yellow_line: 
                cor, status = "#f1c40f", "ARCO AMARELO"
            elif min_green is not None and pico < min_green: 
                cor, status = "#f1c40f", "ARCO AMARELO"
        else:
            if pico >= red_line: 
                cor, status = "#e74c3c", "LINHA VERMELHA"
            elif yellow_line is not None and pico >= yellow_line: 
                cor, status = "#f1c40f", "ARCO AMARELO"
            elif min_green is not None and pico < min_green: 
                cor, status = "#f1c40f", "ARCO AMARELO"
            
    pico_str = ("{"+formato+"}").format(pico)
    limite_str = ("{"+formato+"}").format(red_line) if red_line is not None else 'N/A'
    
    html = f"""
    <div style="background-color: #1e1e1e; padding: 15px; border-radius: 8px; border-left: 6px solid {cor}; margin-bottom: 15px; height: 130px; display: flex; flex-direction: column; justify-content: space-between; box-shadow: 2px 2px 10px rgba(0,0,0,0.3);">
        <p style="margin: 0; font-size: 13px; color: #cccccc; font-weight: bold;">{titulo}</p>
        <h3 style="margin: 0; color: {cor}; font-size: 28px;">{pico_str} <span style="font-size: 14px;">{unidade}</span></h3>
        <div style="display: flex; justify-content: space-between; align-items: flex-end;">
            <span style="font-size: 11px; color: #888888;">Limite Max: {limite_str}</span>
            <span style="font-size: 11px; color: {cor}; font-weight: bold;">{status}</span>
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
    <style>
    body {{ background-color: #0e1117; color: #fafafa; font-family: sans-serif; margin: 0; }}
    .tabela-container {{ max-height: 400px; overflow-y: auto; overflow-x: auto; border: 1px solid #444; border-radius: 8px; }}
    .tabela-container::-webkit-scrollbar {{ width: 10px; height: 10px; }}
    .tabela-container::-webkit-scrollbar-thumb {{ background: #3498db; border-radius: 5px; }}
    .tabela-grande {{ width: 100%; font-size: 15px; text-align: center; border-collapse: collapse; white-space: nowrap; }}
    .tabela-grande th {{ background-color: #2c3e50; color: white; font-size: 15px; padding: 12px; position: sticky; top: 0; z-index: 1; border-bottom: 2px solid #3498db; }}
    .tabela-grande td {{ padding: 10px; border-bottom: 1px solid #333; }}
    .tabela-grande tr:hover {{ background-color: #1a252f; }}
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
        traducao[f'E{i} ITT'] = f'ITT M{i} (ºC)'
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


# ==========================================
# 6. GERADORES DE DOCX (WORD) - SUBSTITUI O PDF
# ==========================================
def gerar_docx_solo(tempos, alertas, obs_dict):
    doc = Document()
    titulo = doc.add_heading("RELATÓRIO OFICIAL DE INSPETORIA - VOO SOLO", 0)
    titulo.alignment = 1
    
    data_hoje = datetime.now().strftime('%d/%m/%Y')
    sub = doc.add_paragraph(f"Perfil: {projeto_atual} | Analista: Filipe Iack | Data: {data_hoje}")
    sub.alignment = 1
    
    doc.add_heading("1. LOGBOOK DA MISSÃO", level=1)
    doc.add_paragraph(f"Acionamento: {tempos.get('acionamento','-')}")
    doc.add_paragraph(f"Decolagem: {tempos.get('decolagem','-')}")
    doc.add_paragraph(f"Pouso: {tempos.get('pouso','-')}")
    doc.add_paragraph(f"Corte do Motor: {tempos.get('corte','-')}")
    doc.add_paragraph(f"TEMPO DE VOO: {tempos.get('t_voo','-')}").runs[0].bold = True
    
    doc.add_heading("2. AUDITORIA DE ALERTAS", level=1)
    if len(alertas) > 0:
        for alerta in alertas:
            p_al = doc.add_paragraph()
            run = p_al.add_run(alerta)
            if "LINHA VERMELHA" in alerta or "CRÍTICO" in alerta or "🔴" in alerta:
                run.font.color.rgb = RGBColor(220, 0, 0)
            elif "AMARELO" in alerta or "🟡" in alerta:
                run.font.color.rgb = RGBColor(200, 150, 0)
    else:
        p_al = doc.add_paragraph()
        run = p_al.add_run("✅ Nenhum alerta registado.")
        run.font.color.rgb = RGBColor(0, 150, 0)
        
    doc.add_heading("3. OBSERVAÇÕES DA INSPETORIA", level=1)
    tem_obs = False
    for titulo_obs, texto in obs_dict.items():
        if str(texto).strip() != "":
            tem_obs = True
            doc.add_heading(f"[{titulo_obs}]", level=2)
            doc.add_paragraph(str(texto))
            
    if not tem_obs:
        doc.add_paragraph("Nenhuma observação inserida.")
        
    doc.add_paragraph("\n\n____________________________________________").alignment = 1
    doc.add_paragraph("Filipe Iack - Especialista em Telemetria AFA").alignment = 1
    
    io_bytes = io.BytesIO()
    doc.save(io_bytes)
    return io_bytes.getvalue()

def gerar_docx_solo_editado(tempos_dict, alertas_texto, obs_dict):
    doc = Document()
    titulo = doc.add_heading("RELATÓRIO OFICIAL DE INSPETORIA - VOO SOLO", 0)
    titulo.alignment = 1
    
    data_hoje = datetime.now().strftime('%d/%m/%Y')
    sub = doc.add_paragraph(f"Perfil: {projeto_atual} | Analista: Filipe Iack | Data: {data_hoje}")
    sub.alignment = 1
    
    doc.add_heading("1. LOGBOOK DA MISSÃO", level=1)
    doc.add_paragraph(f"Acionamento: {tempos_dict.get('acionamento','-')}")
    doc.add_paragraph(f"Decolagem: {tempos_dict.get('decolagem','-')}")
    doc.add_paragraph(f"Pouso: {tempos_dict.get('pouso','-')}")
    doc.add_paragraph(f"Corte do Motor: {tempos_dict.get('corte','-')}")
    doc.add_paragraph(f"TEMPO DE VOO: {tempos_dict.get('t_voo','-')}").runs[0].bold = True
    
    doc.add_heading("2. AUDITORIA DE ALERTAS", level=1)
    linhas_alerta = [linha for linha in alertas_texto.split('\n') if linha.strip()]
    
    if len(linhas_alerta) > 0:
        for alerta in linhas_alerta:
            p_al = doc.add_paragraph()
            run = p_al.add_run(alerta)
            if "LINHA VERMELHA" in alerta or "CRÍTICO" in alerta or "🔴" in alerta:
                run.font.color.rgb = RGBColor(220, 0, 0)
            elif "AMARELO" in alerta or "🟡" in alerta:
                run.font.color.rgb = RGBColor(200, 150, 0)
    else:
        p_al = doc.add_paragraph()
        run = p_al.add_run("✅ Nenhum alerta registado.")
        run.font.color.rgb = RGBColor(0, 150, 0)
        
    doc.add_heading("3. OBSERVAÇÕES DA INSPETORIA", level=1)
    tem_obs = False
    for titulo_obs, texto in obs_dict.items():
        if str(texto).strip() != "":
            tem_obs = True
            doc.add_heading(f"[{titulo_obs}]", level=2)
            doc.add_paragraph(str(texto))
            
    if not tem_obs:
        doc.add_paragraph("Nenhuma observação inserida.")
        
    doc.add_paragraph("\n\n____________________________________________").alignment = 1
    doc.add_paragraph("Filipe Iack - Especialista em Telemetria AFA").alignment = 1
    
    io_bytes = io.BytesIO()
    doc.save(io_bytes)
    return io_bytes.getvalue()

def gerar_docx_formatura(arq_lider, arq_ala, obs_dict):
    doc = Document()
    titulo = doc.add_heading("RELATÓRIO TÁTICO DE FORMATURA", 0)
    titulo.alignment = 1
    
    data_hoje = datetime.now().strftime('%d/%m/%Y')
    sub = doc.add_paragraph(f"Analista: Filipe Iack | Data: {data_hoje}")
    sub.alignment = 1
    
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
            
    if not tem_obs:
        doc.add_paragraph("Nenhuma observação inserida no dossiê.")
        
    doc.add_paragraph("\n\n____________________________________________").alignment = 1
    doc.add_paragraph("Filipe Iack - Especialista em Telemetria AFA").alignment = 1
    
    io_bytes = io.BytesIO()
    doc.save(io_bytes)
    return io_bytes.getvalue()

def gerar_docx_formatura_editado(arq_lider, arq_ala, obs_dict):
    return gerar_docx_formatura(arq_lider, arq_ala, obs_dict)


# ==========================================
# 7. UPLOADER E FLUXO PRINCIPAL
# ==========================================
arquivos_up = st.file_uploader(f"Carregar arquivo(s) .csv do Garmin (Perfil: {projeto_atual})", type=['csv'], accept_multiple_files=True)

if arquivos_up:
    
    # ---------------------------------------------------------
    # FLUXO: VOO EM FORMATURA
    # ---------------------------------------------------------
    if modo_analise == "Voo em Formatura (Líder e Ala)":
        if len(arquivos_up) != 2:
            st.warning("⚠️ Para o Modo Formatura, você precisa carregar **exatamente 2 arquivos** simultaneamente (O Líder e o Ala).")
        else:
            st.markdown("### 🔀 Mapeamento de Aeronaves")
            st.markdown("Defina abaixo qual é o ficheiro do Líder e qual é o do Ala.")
            
            nome_arq1 = arquivos_up[0].name
            nome_arq2 = arquivos_up[1].name
            
            col_lider, col_ala = st.columns(2)
            with col_lider:
                arq_lider = st.selectbox("Selecione o ficheiro do LÍDER (Laranja):", [nome_arq1, nome_arq2], index=0)
            with col_ala:
                idx_padrao_ala = 1 if arq_lider == nome_arq1 else 0
                arq_ala = st.selectbox("Selecione o ficheiro do ALA (Azul):", [nome_arq1, nome_arq2], index=idx_padrao_ala)

            st.divider()
            
            c_info, c_sync = st.columns([2, 1])
            with c_info:
                st.info("💡 **Dica de Sincronização:** Se um arquivo estiver em Zulu (UTC) e outro em Local, ajuste o relógio do Ala para alinharem perfeitamente.")
            with c_sync:
                ajuste_horas = st.number_input("⏳ Ajuste de Fuso do Ala (Horas):", min_value=-12, max_value=12, value=0, step=1, help="Use -3 ou +3")
                
            st.sidebar.markdown("---")
            st.sidebar.success("✅ **Análise Carregada**")

            if arq_lider == arq_ala:
                st.error("⚠️ Você selecionou o mesmo ficheiro para o Líder e para o Ala. Mude um deles nas caixas acima.")
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
                    st.error("❌ Os relógios não coincidem. Tente ajustar o 'Ajuste de Fuso do Ala' acima para -3 ou +3.")
                else:
                    aba_dist, aba_comp, aba_3d, aba_rel_form = st.tabs(["📏 1. DISTÂNCIA", "⚙️ 2. COMPARATIVO", "✈️ 3. TRAJETÓRIA 3D", "📄 4. RELATÓRIO FINAL"])
                    
                    with aba_dist:
                        st.header("Análise de Distância na Formatura")
                        st.markdown("Verifique em que exato momento e manobra o Ala se afastou ou se aproximou excessivamente do Líder.")
                        
                        df_merge['Distancia_m'] = formula_haversine(
                            df_merge['Latitude_lider'], df_merge['Longitude_lider'], 
                            df_merge['Latitude_ala'], df_merge['Longitude_ala']
                        )
                        
                        fig_dist = make_subplots(
                            rows=3, cols=1, shared_xaxes=True, vertical_spacing=0.08, 
                            subplot_titles=("1. Distância Horizontal (Metros)", "2. Perfil de Altitude (ft)", "3. Atitude da Aeronave (Arfagem/Pitch º)")
                        )
                        
                        fig_dist.add_trace(go.Scattergl(
                            x=df_merge['Datetime'], y=df_merge['Distancia_m'], 
                            name="Distância (m)", fill='tozeroy', fillcolor='rgba(149, 165, 166, 0.2)', 
                            line=dict(color='#bdc3c7', width=2)
                        ), row=1, col=1)
                        
                        if 'AltMSL_lider' in df_merge.columns: 
                            fig_dist.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['AltMSL_lider'], name="Alt Líder", line=dict(color='#e67e22', width=2)), row=2, col=1)
                        if 'AltMSL_ala' in df_merge.columns: 
                            fig_dist.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['AltMSL_ala'], name="Alt Ala", line=dict(color='#3498db', width=2, dash='dot')), row=2, col=1)
                            
                        if 'Pitch_lider' in df_merge.columns: 
                            fig_dist.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['Pitch_lider'], name="Pitch Líder", line=dict(color='#e67e22', width=2)), row=3, col=1)
                        if 'Pitch_ala' in df_merge.columns: 
                            fig_dist.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['Pitch_ala'], name="Pitch Ala", line=dict(color='#3498db', width=2, dash='dot')), row=3, col=1)

                        fig_dist.update_xaxes(showspikes=True, spikecolor="white", spikesnap="cursor", spikemode="across", spikethickness=1, spikedash="dot")
                        fig_dist.update_layout(template="plotly_dark", height=850, hovermode="x unified", margin=dict(l=10, r=10, t=30, b=10))
                        st.plotly_chart(fig_dist, use_container_width=True)
                        
                        st.subheader("🚨 Top 15 Momentos de Maior Afastamento")
                        df_top_dist = df_merge.sort_values(by='Distancia_m', ascending=False).head(15).copy()
                        
                        df_formatada = pd.DataFrame()
                        df_formatada['Hora (Líder)'] = df_top_dist['Datetime'].dt.strftime('%H:%M:%S')
                        df_formatada['Distância (m)'] = df_top_dist['Distancia_m'].apply(lambda x: f"{x:.1f} m")
                        if 'AltMSL_lider' in df_top_dist.columns: 
                            df_formatada['Alt. Líder (ft)'] = df_top_dist['AltMSL_lider'].apply(lambda x: f"{x:.0f}")
                        if 'Pitch_lider' in df_top_dist.columns: 
                            df_formatada['Arfagem Líder (º)'] = df_top_dist['Pitch_lider'].apply(lambda x: f"{x:.1f}")
                        
                        render_html_table(df_formatada)

                        st.divider()
                        st.markdown("📝 **Observações - Distância** (Esses comentários irão para o Word)")
                        st.session_state.obs_distancia = st.text_area("Escreva aqui:", value=st.session_state.obs_distancia, key='k_dist')
                        if st.button("💾 Salvar Comentário", key="b_dist"): 
                            st.success("Salvo com sucesso!")

                    with aba_comp:
                        st.header("Comparativo Dinâmico e Motores")
                        st.markdown("Líder em **Laranja**, Ala em **Azul**.")
                        
                        fig_comp = make_subplots(
                            rows=6, cols=1, shared_xaxes=True, vertical_spacing=0.03, 
                            subplot_titles=("1. Altitude (ft) e Velocidade (kt)", "2. Força G", "3. Torque (Lbs)", "4. ITT (ºC)", "5. Ng (%)", "6. Nh (%)")
                        )
                        
                        if 'AltMSL_lider' in df_merge.columns: 
                            fig_comp.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['AltMSL_lider'], name="Alt Líder", line=dict(color='#e67e22', width=2)), row=1, col=1)
                        if 'AltMSL_ala' in df_merge.columns: 
                            fig_comp.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['AltMSL_ala'], name="Alt Ala", line=dict(color='#3498db', width=2, dash='dot')), row=1, col=1)
                        
                        if 'NormAc_lider' in df_merge.columns: 
                            fig_comp.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['NormAc_lider'], name="G Líder", line=dict(color='#e67e22', width=2)), row=2, col=1)
                        if 'NormAc_ala' in df_merge.columns: 
                            fig_comp.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['NormAc_ala'], name="G Ala", line=dict(color='#3498db', width=2)), row=2, col=1)
                        
                        if 'E1 Torq ft lbs_lider' in df_merge.columns: 
                            fig_comp.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['E1 Torq ft lbs_lider'], name="Tq Líder", line=dict(color='#e67e22', width=2)), row=3, col=1)
                        if 'E1 Torq ft lbs_ala' in df_merge.columns: 
                            fig_comp.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['E1 Torq ft lbs_ala'], name="Tq Ala", line=dict(color='#3498db', width=2)), row=3, col=1)
                        
                        if 'E1 ITT_lider' in df_merge.columns: 
                            fig_comp.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['E1 ITT_lider'], name="ITT Líder", line=dict(color='#e67e22', width=2)), row=4, col=1)
                        if 'E1 ITT_ala' in df_merge.columns: 
                            fig_comp.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['E1 ITT_ala'], name="ITT Ala", line=dict(color='#3498db', width=2)), row=4, col=1)
                        
                        if 'E1 NG_lider' in df_merge.columns: 
                            fig_comp.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['E1 NG_lider'], name="Ng Líder", line=dict(color='#e67e22', width=2)), row=5, col=1)
                        if 'E1 NG_ala' in df_merge.columns: 
                            fig_comp.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['E1 NG_ala'], name="Ng Ala", line=dict(color='#3498db', width=2)), row=5, col=1)
                        
                        if 'E1 NP %_lider' in df_merge.columns: 
                            fig_comp.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['E1 NP %_lider'], name="Nh Líder", line=dict(color='#e67e22', width=2)), row=6, col=1)
                        if 'E1 NP %_ala' in df_merge.columns: 
                            fig_comp.add_trace(go.Scattergl(x=df_merge['Datetime'], y=df_merge['E1 NP %_ala'], name="Nh Ala", line=dict(color='#3498db', width=2)), row=6, col=1)
                        
                        fig_comp.update_xaxes(showspikes=True, spikecolor="white", spikesnap="cursor", spikemode="across", spikethickness=1, spikedash="dot")
                        fig_comp.update_layout(template="plotly_dark", height=1200, hovermode="x unified", margin=dict(l=10, r=10, t=30, b=10))
                        st.plotly_chart(fig_comp, use_container_width=True)

                        st.divider()
                        st.markdown("📝 **Observações - Comparativo de Motores**")
                        st.session_state.obs_comparativo = st.text_area("Escreva aqui:", value=st.session_state.obs_comparativo, key='k_comp')
                        if st.button("💾 Salvar Comentário", key="b_comp"): 
                            st.success("Salvo com sucesso!")

                    with aba_3d:
                        st.header("Trajetória 3D Sincronizada")
                        fig_3d = go.Figure()
                        fig_3d.add_trace(go.Scatter3d(
                            x=df_merge['Longitude_lider'], y=df_merge['Latitude_lider'], z=df_merge['AltMSL_lider'], 
                            mode='lines', line=dict(color='#e67e22', width=4), name="Líder (Laranja)"
                        ))
                        fig_3d.add_trace(go.Scatter3d(
                            x=df_merge['Longitude_ala'], y=df_merge['Latitude_ala'], z=df_merge['AltMSL_ala'], 
                            mode='lines', line=dict(color='#3498db', width=4), name="Ala (Azul)"
                        ))
                        fig_3d.update_layout(template="plotly_dark", height=700, margin=dict(l=0, r=0, b=0, t=20), scene=dict(aspectmode='auto'))
                        st.plotly_chart(fig_3d, use_container_width=True)
                        
                        st.divider()
                        st.markdown("📝 **Observações - Trajetória 3D**")
                        st.session_state.obs_3d = st.text_area("Escreva aqui:", value=st.session_state.obs_3d, key='k_3d')
                        if st.button("💾 Salvar Comentário", key="b_3d"): 
                            st.success("Salvo com sucesso!")

                    with aba_rel_form:
                        st.header("📝 Editor do Relatório Final (Formatura)")
                        st.markdown("Confirme ou edite os dados extraídos antes de gerar o Word.")
                        
                        c1, c2 = st.columns(2)
                        with c1: lider_editado = st.text_input("Ficheiro Líder:", arq_lider)
                        with c2: ala_editado = st.text_input("Ficheiro Ala:", arq_ala)
                        
                        st.info("As observações editadas nas outras abas serão incluídas automaticamente.")
                        
                        obs_formatura_atual = {
                            "Posicionamento e Distancia": st.session_state.obs_distancia,
                            "Comparativo Dinamico (Motores)": st.session_state.obs_comparativo,
                            "Trajetoria Espacial 3D": st.session_state.obs_3d
                        }
                        
                        st.divider()
                        st.subheader("🖨️ Gerar Documento Word")
                        docx_bytes_form_edit = gerar_docx_formatura_editado(lider_editado, ala_editado, obs_formatura_atual)
                        st.download_button(
                            label="💾 BAIXAR DOSSIÊ DE FORMATURA (.DOCX)", 
                            data=docx_bytes_form_edit, 
                            file_name="Dossie_Formatura.docx", 
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
                            use_container_width=True
                        )

    # ---------------------------------------------------------
    # FLUXO: VOO SOLO
    # ---------------------------------------------------------
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
                vel = f"{vel_val:.0f} nós" if pd.notna(vel_val) else 'N/A'
                alertas.append(f"{mensagem} -> Ocorrência às {tempo} | Alt: {alt} | Vel: {vel}")
            except: 
                alertas.append(mensagem)

        if l_g_max is not None and max_g >= l_g_max: 
            add_alerta_contexto('NormAc', max_g, f"🔴 LINHA VERMELHA: Força G Positiva ({max_g:.2f}G)")
        if l_g_min is not None and min_g <= l_g_min: 
            add_alerta_contexto('NormAc', min_g, f"🔴 LINHA VERMELHA: Força G Negativa ({min_g:.2f}G)")
        
        # Regras de Negócios Restritas para T-27M
        if projeto_atual == "T-27M (Oficial)":
            max_torq = df['E1 Torq ft lbs'].max() if 'E1 Torq ft lbs' in df.columns else 0
            max_itt = df['E1 ITT'].max() if 'E1 ITT' in df.columns else 0
            max_ng = df['E1 NG'].max() if 'E1 NG' in df.columns else 0
            max_nh = df['E1 NP %'].max() if 'E1 NP %' in df.columns else 0

            if verificar_temporizador(df, 'E1 Torq ft lbs', 2100, 2): 
                add_alerta_contexto('E1 Torq ft lbs', max_torq, f"🔴 OVER-TORQUE CRÍTICO: >2100 Lbs por 2 seg")
            elif verificar_temporizador(df, 'E1 Torq ft lbs', 1970, 1): 
                add_alerta_contexto('E1 Torq ft lbs', max_torq, f"🔴 OVER-TORQUE: Passou de 1970 Lbs")
            elif l_tq_red is not None and max_torq >= l_tq_red: 
                add_alerta_contexto('E1 Torq ft lbs', max_torq, f"🔴 LINHA VERMELHA Torque: {max_torq:.0f} Lbs")
            elif l_tq_yel is not None and max_torq >= l_tq_yel: 
                add_alerta_contexto('E1 Torq ft lbs', max_torq, f"🟡 ARCO AMARELO Torque: {max_torq:.0f} Lbs")
            
            if verificar_temporizador(df, 'E1 NG', 102.6, 6): 
                add_alerta_contexto('E1 NG', max_ng, f"🔴 NG CRÍTICO: >102.6% por 6 seg")
            elif verificar_temporizador(df, 'E1 NG', 101.6, 1): 
                add_alerta_contexto('E1 NG', max_ng, f"🔴 LINHA VERMELHA Ng: Passou de 101.6%")
            elif l_ng_red is not None and max_ng >= l_ng_red: 
                add_alerta_contexto('E1 NG', max_ng, f"🔴 LINHA VERMELHA Ng: {max_ng:.1f}%")
            elif l_ng_yel is not None and max_ng >= l_ng_yel: 
                add_alerta_contexto('E1 NG', max_ng, f"🟡 ARCO AMARELO Ng: {max_ng:.1f}%")
            
            if verificar_temporizador(df, 'E1 ITT', 850, 2): 
                add_alerta_contexto('E1 ITT', max_itt, f"🔴 SUPER AQUECIMENTO ITT: >850ºC por 2 seg")
            elif verificar_temporizador(df, 'E1 ITT', 790, 1): 
                add_alerta_contexto('E1 ITT', max_itt, f"🔴 LINHA VERMELHA ITT: Passou de 790ºC")
            elif l_itt_yel is not None and max_itt >= l_itt_yel: 
                add_alerta_contexto('E1 ITT', max_itt, f"🟡 ARCO AMARELO ITT: {max_itt:.0f} ºC")
            
            if verificar_temporizador(df, 'E1 NP %', 110, 1): 
                add_alerta_contexto('E1 NP %', max_nh, f"🔴 OVERSPEED HÉLICE CRÍTICO: >110%")
            elif verificar_temporizador(df, 'E1 NP %', 106, 1): 
                add_alerta_contexto('E1 NP %', max_nh, f"🔴 OVERSPEED HÉLICE: >106%")
            elif l_nh_red is not None and max_nh >= l_nh_red: 
                add_alerta_contexto('E1 NP %', max_nh, f"🔴 LINHA VERMELHA Nh: {max_nh:.1f}%")
            elif l_nh_yel is not None and max_nh >= l_nh_yel: 
                add_alerta_contexto('E1 NP %', max_nh, f"🟡 ARCO AMARELO Nh: {max_nh:.1f}%")
        
        # Regras Gerais para Multimotores / Outras Aeronaves
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

        # Regras de Temperatura/Pressão do Óleo (Aplicadas a todos os motores do perfil)
        for i_mot in range(1, qtd_motores_config + 1):
            m_oilt = df[f'E{i_mot} OilT_Celsius'].max() if f'E{i_mot} OilT_Celsius' in df.columns else 0
            m_oilp = df[f'E{i_mot} OilP'].max() if f'E{i_mot} OilP' in df.columns else 0
            
            if l_oilt_red is not None and m_oilt >= l_oilt_red: add_alerta_contexto(f'E{i_mot} OilT_Celsius', m_oilt, f"🔴 Motor {i_mot} LINHA VERMELHA Temp Óleo: {m_oilt:.0f} ºC")
            if l_oilp_red is not None and m_oilp >= l_oilp_red: add_alerta_contexto(f'E{i_mot} OilP', m_oilp, f"🔴 Motor {i_mot} LINHA VERMELHA Pressão Óleo: {m_oilp:.0f} PSI")

        tempos = {'acionamento': '-', 'decolagem': '-', 'pouso': '-', 'corte': '-', 't_voo': '-'}
        
        # Ajuste dinâmico para Helicópteros ou Aviões no Gatilho de Motor
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

        st.sidebar.markdown("---")
        st.sidebar.success("✅ **Análise Carregada**")

        aba_didatica, aba_inspetoria, aba_manobras, aba_relatorio = st.tabs(["🎓 1. RESUMO DE PAINEL", "⚙️ 2. INSPETORIA E GRÁFICOS", "🛩️ 3. AVALIADOR DE MANOBRAS", "📄 4. RELATÓRIO FINAL"])

        with aba_didatica:
            st.header(f"🎯 Resumo de Picos - {projeto_atual}")
            c1, c2, c3, c4 = st.columns(4)
            
            with c1:
                exibir_card_arcos("Força G Máxima", max_g, l_g_max, None, None, "G", ":.2f")
                max_ng = df['E1 NG'].max() if 'E1 NG' in df.columns else 0
                exibir_card_arcos("Gerador de Gás (Ng)", max_ng, l_ng_red, l_ng_yel, None, "%", ":.1f")
            with c2:
                exibir_card_arcos("Força G Mínima", min_g, l_g_min, None, None, "G", ":.2f", e_minimo=True)
                max_nh = df['E1 NP %'].max() if 'E1 NP %' in df.columns else 0
                exibir_card_arcos("Hélice (Nh)", max_nh, l_nh_red, l_nh_yel, None, "%", ":.1f")
            with c3:
                max_torq = df['E1 Torq ft lbs'].max() if 'E1 Torq ft lbs' in df.columns else 0
                exibir_card_arcos("Torque Máximo", max_torq, l_tq_red, l_tq_yel, None, "Lbs", ":.0f")
                max_oilt = df['E1 OilT_Celsius'].max() if 'E1 OilT_Celsius' in df.columns else 0
                exibir_card_arcos("Temp. Óleo (OilT)", max_oilt, l_oilt_red, None, l_oilt_min, "ºC", ":.0f")
            with c4:
                max_itt = df['E1 ITT'].max() if 'E1 ITT' in df.columns else 0
                exibir_card_arcos("Temperatura ITT", max_itt, l_itt_red, l_itt_yel, None, "ºC", ":.0f")
                max_oilp = df['E1 OilP'].max() if 'E1 OilP' in df.columns else 0
                exibir_card_arcos("Pressão Óleo (OilP)", max_oilp, l_oilp_red, None, l_oilp_min, "PSI", ":.0f")
                
            if qtd_motores_config > 1:
                st.caption(f"*Valores do Motor 1 no painel. O rastreio dos {qtd_motores_config} motores está detalhado na aba de Inspetoria.*")

            st.divider()
            st.subheader("🗺️ Área de Instrução (Rota e Google Earth)")
            df_mapa = df[(df['IAS'] > 30)].dropna(subset=['Latitude', 'Longitude'])
            if not df_mapa.empty:
                fig_map = go.Figure()
                fig_map.add_trace(go.Scattermap(
                    lat=df_mapa['Latitude'], lon=df_mapa['Longitude'], 
                    mode='lines', line=dict(width=4, color='#3498db')
                ))
                fig_map.add_trace(go.Scattermap(
                    lat=[df_mapa['Latitude'].iloc[0]], lon=[df_mapa['Longitude'].iloc[0]], 
                    mode='markers', marker=dict(size=14, color='#2ecc71')
                ))
                fig_map.add_trace(go.Scattermap(
                    lat=[df_mapa['Latitude'].iloc[-1]], lon=[df_mapa['Longitude'].iloc[-1]], 
                    mode='markers', marker=dict(size=14, color='#e74c3c')
                ))
                fig_map.update_layout(
                    map_style="carto-darkmatter", map_zoom=8, 
                    map_center={"lat": df_mapa['Latitude'].mean(), "lon": df_mapa['Longitude'].mean()}, 
                    margin={"r":0,"t":0,"l":0,"b":0}, height=400, showlegend=False
                )
                st.plotly_chart(fig_map, use_container_width=True)
                
                # Geração do arquivo KML tridimensional compatível com o Google Earth (Absolute)
                kml = simplekml.Kml()
                linestring = kml.newlinestring(
                    name="Trajetoria de Voo 3D", 
                    coords=[(r['Longitude'], r['Latitude'], r['AltMSL'] * 0.3048) for i, r in df_mapa.iterrows()]
                )
                linestring.altitudemode = simplekml.AltitudeMode.absolute
                linestring.style.linestyle.color = simplekml.Color.cyan
                linestring.style.linestyle.width = 4

                ponto_inicio = kml.newpoint(
                    name="Início da Missão",
                    coords=[(df_mapa['Longitude'].iloc[0], df_mapa['Latitude'].iloc[0], df_mapa['AltMSL'].iloc[0] * 0.3048)]
                )
                ponto_inicio.altitudemode = simplekml.AltitudeMode.absolute
                ponto_inicio.style.iconstyle.icon.href = "http://maps.google.com/mapfiles/kml/paddle/grn-circle.png"

                ponto_fim = kml.newpoint(
                    name="Término / Pouso",
                    coords=[(df_mapa['Longitude'].iloc[-1], df_mapa['Latitude'].iloc[-1], df_mapa['AltMSL'].iloc[-1] * 0.3048)]
                )
                ponto_fim.altitudemode = simplekml.AltitudeMode.absolute
                ponto_fim.style.iconstyle.icon.href = "http://maps.google.com/mapfiles/kml/paddle/red-circle.png"
                
                st.download_button("⬇️ Baixar Rota 3D (Google Earth .KML)", data=kml.kml(), file_name="rota_voo.kml", mime="application/vnd.google-earth.kml+xml")

            st.divider()
            st.markdown("📝 **Observações - Resumo de Painel e Rota**")
            st.session_state.obs_resumo = st.text_area("Escreva aqui:", value=st.session_state.obs_resumo, key='k_res')
            if st.button("💾 Salvar Comentário", key="b_res"): 
                st.success("Salvo com sucesso!")

        with aba_inspetoria:
            st.subheader("Auditoria de Zonas Amarelas e Vermelhas")
            if len(alertas) > 0:
                for a in alertas:
                    if "LINHA VERMELHA" in a or "CRÍTICO" in a: 
                        st.error(a)
                    else: 
                        st.warning(a)
            else:
                st.success("✅ Voo totalmente dentro dos parâmetros Verdes.")
                
            st.divider()
            st.subheader("🔍 Investigação de Eventos")
            c_busca1, c_busca2, c_busca3 = st.columns([2, 1, 1])
            opcoes_busca = {
                'Força G': 'NormAc', 'Torque M1': 'E1 Torq ft lbs', 'ITT M1': 'E1 ITT', 
                'Ng M1 (%)': 'E1 NG', 'Nh M1 (%)': 'E1 NP %', 'Temp. Óleo M1': 'E1 OilT_Celsius', 
                'Pressão Óleo M1': 'E1 OilP', 'Velocidade': 'IAS', 'Altitude': 'AltMSL', 
                'Arfagem (Pitch)': 'Pitch', 'Rolagem (Roll)': 'Roll'
            }
            if qtd_motores_config > 1:
                opcoes_busca.update({
                    'Torque M2': 'E2 Torq ft lbs', 'ITT M2': 'E2 ITT',
                    'Ng M2 (%)': 'E2 NG', 'Nh M2 (%)': 'E2 NP %',
                    'Temp. Óleo M2': 'E2 OilT_Celsius', 'Pressão Óleo M2': 'E2 OilP'
                })
            
            with c_busca1: 
                escolha = st.selectbox("Parâmetro:", list(opcoes_busca.keys()))
            with c_busca2: 
                operador = st.selectbox("Condição:", ["Maior que >", "Menor que <"])
            with c_busca3: 
                valor = st.number_input("Valor:", value=3.0)
                
            coluna_filtro = opcoes_busca[escolha]
            if coluna_filtro in df.columns:
                resultado_busca = df[df[coluna_filtro] > valor] if operador == "Maior que >" else df[df[coluna_filtro] < valor]
                if not resultado_busca.empty:
                    colunas_ver = ['Lcl Time', 'AltMSL', 'IAS', 'NormAc', 'Pitch', 'Roll']
                    for j in range(1, qtd_motores_config + 1):
                        colunas_ver.extend([f'E{j} Torq ft lbs', f'E{j} ITT', f'E{j} NG', f'E{j} NP %', f'E{j} OilT_Celsius', f'E{j} OilP'])
                    colunas_ver = [c for c in colunas_ver if c in df.columns]
                    exibir_tabela_grande(resultado_busca[colunas_ver].head(20).copy())
                else: 
                    st.info("Nenhum evento encontrado.")

            st.divider()
            st.subheader("📈 Dashboard de Dinâmica e Motor")
            
            fig = make_subplots(
                rows=6, cols=1, shared_xaxes=True, vertical_spacing=0.03, 
                specs=[[{"secondary_y": True}], [{"secondary_y": False}], [{"secondary_y": False}], [{"secondary_y": False}], [{"secondary_y": False}], [{"secondary_y": False}]], 
                subplot_titles=("1. Altitude e Velocidade", "2. Força G", "3. Torque (Lbs)", "4. Temperatura ITT (ºC)", "5. Gerador de Gás - Ng (%)", "6. Rotação da Hélice - Nh (%)")
            )
            
            cores_motores = ['#27ae60', '#f1c40f', '#e67e22', '#8e44ad']
            cores_itt = ['#f39c12', '#e74c3c', '#d35400', '#c0392b']
            cores_ng = ['#9b59b6', '#3498db', '#1abc9c', '#e67e22']
            cores_nh = ['#f1c40f', '#2ecc71', '#34495e', '#bdc3c7']
            
            if 'AltMSL' in df.columns: 
                fig.add_trace(go.Scattergl(x=df['Datetime'], y=df['AltMSL'], name="Altitude (ft)", fill='tozeroy', fillcolor='rgba(52, 152, 219, 0.1)', line=dict(color='#3498db', width=2)), row=1, col=1, secondary_y=False)
            if 'IAS' in df.columns: 
                fig.add_trace(go.Scattergl(x=df['Datetime'], y=df['IAS'], name="Velocidade (kt)", line=dict(color='#00d2d3', width=2)), row=1, col=1, secondary_y=True)
                
            if 'NormAc' in df.columns: 
                fig.add_trace(go.Scattergl(x=df['Datetime'], y=df['NormAc'], name="Força G", fill='tozeroy', fillcolor='rgba(231, 76, 60, 0.1)', line=dict(color='#e74c3c', width=2)), row=2, col=1)
                if l_g_max is not None: 
                    fig.add_hline(y=l_g_max, line_dash="dash", line_color="red", row=2, col=1)
                    
            for i_m in range(1, qtd_motores_config + 1):
                col_tq = f'E{i_m} Torq ft lbs'
                if col_tq in df.columns: 
                    fig.add_trace(go.Scattergl(x=df['Datetime'], y=df[col_tq], name=f"Tq M{i_m}", line=dict(color=cores_motores[i_m-1], width=2)), row=3, col=1)
            if l_tq_red is not None: fig.add_hline(y=l_tq_red, line_dash="dash", line_color="red", row=3, col=1)
            if l_tq_yel is not None and l_tq_red is not None and l_tq_yel != l_tq_red: fig.add_hrect(y0=l_tq_yel, y1=l_tq_red, fillcolor="rgba(241, 196, 15, 0.2)", layer="below", line_width=0, row=3, col=1)
                    
            for i_m in range(1, qtd_motores_config + 1):
                col_itt = f'E{i_m} ITT'
                if col_itt in df.columns: 
                    fig.add_trace(go.Scattergl(x=df['Datetime'], y=df[col_itt], name=f"ITT M{i_m}", line=dict(color=cores_itt[i_m-1], width=2)), row=4, col=1)
            if l_itt_red is not None: fig.add_hline(y=l_itt_red, line_dash="dash", line_color="red", row=4, col=1)
            if l_itt_yel is not None and l_itt_red is not None: fig.add_hrect(y0=l_itt_yel, y1=l_itt_red, fillcolor="rgba(241, 196, 15, 0.2)", layer="below", line_width=0, row=4, col=1)
                    
            for i_m in range(1, qtd_motores_config + 1):
                col_ng = f'E{i_m} NG'
                if col_ng in df.columns:
                    fig.add_trace(go.Scattergl(x=df['Datetime'], y=df[col_ng], name=f"Ng M{i_m} (%)", line=dict(color=cores_ng[i_m-1], width=2)), row=5, col=1)
            if l_ng_red is not None: fig.add_hline(y=l_ng_red, line_dash="dash", line_color="red", row=5, col=1)
            if l_ng_yel is not None and l_ng_red is not None: fig.add_hrect(y0=l_ng_yel, y1=l_ng_red, fillcolor="rgba(241, 196, 15, 0.2)", layer="below", line_width=0, row=5, col=1)
                    
            for i_m in range(1, qtd_motores_config + 1):
                col_nh = f'E{i_m} NP %'
                if col_nh in df.columns:
                    fig.add_trace(go.Scattergl(x=df['Datetime'], y=df[col_nh], name=f"Nh M{i_m} (%)", line=dict(color=cores_nh[i_m-1], width=2)), row=6, col=1)
            if l_nh_red is not None: fig.add_hline(y=l_nh_red, line_dash="dash", line_color="red", row=6, col=1)
            if l_nh_yel is not None and l_nh_red is not None: fig.add_hrect(y0=l_nh_yel, y1=l_nh_red, fillcolor="rgba(241, 196, 15, 0.2)", layer="below", line_width=0, row=6, col=1)
            
            fig.update_xaxes(showspikes=True, spikecolor="white", spikesnap="cursor", spikemode="across", spikethickness=1, spikedash="dot")
            fig.update_yaxes(showspikes=True, spikecolor="gray", spikethickness=1, spikedash="dot")
            
            fig.update_layout(
                template="plotly_dark", height=1200, 
                hovermode="x unified", margin=dict(l=10, r=10, t=30, b=10), 
                legend=dict(title="Sensores", orientation="v", yanchor="top", y=1, xanchor="left", x=1.02, bgcolor="rgba(0,0,0,0.5)")
            )
            st.plotly_chart(fig, use_container_width=True)

            st.divider()
            st.markdown("📝 **Observações - Inspetoria e Gráficos**")
            st.session_state.obs_graficos = st.text_area("Escreva aqui:", value=st.session_state.obs_graficos, key='k_graf')
            if st.button("💾 Salvar Comentário", key="b_graf"): 
                st.success("Salvo com sucesso!")

        with aba_manobras:
            c_m1, c_m2 = st.columns(2)
            with c_m1:
                st.subheader("🛬 Avaliador de Pouso")
                if 'IAS' in df.columns and 'AltMSL' in df.columns and not voo_ativo.empty:
                    fase_final = voo_ativo.tail(300) 
                    pico_g_pouso = fase_final['NormAc'].max()
                    vel_aprox = fase_final['IAS'].mean()
                    
                    st.info(f"**Velocidade Média na Aproximação Final:** {vel_aprox:.1f} nós (kt)")
                    if pico_g_pouso > 2.0: 
                        st.error(f"🔴 **Impacto Duro:** {pico_g_pouso:.2f}G no toque.")
                    elif pico_g_pouso > 1.5: 
                        st.warning(f"🟡 **Pouso Firme:** {pico_g_pouso:.2f}G na pista.")
                    else: 
                        st.success(f"✅ **Pouso Suave:** {pico_g_pouso:.2f}G.")
            with c_m2:
                st.subheader("🤸 Isolador Automático de Acrobacias")
                if 'Pitch' in df.columns and 'Roll' in df.columns:
                    df_acro = df[(df['Pitch'] > 45) | (df['Pitch'] < -45) | (df['Roll'] > 60) | (df['Roll'] < -60)]
                    if not df_acro.empty:
                        st.markdown(f"**Atitudes anormais:** {len(df_acro)} momentos.")
                        colunas_acro = ['Lcl Time', 'AltMSL', 'IAS', 'Pitch', 'Roll', 'NormAc']
                        for j in range(1, qtd_motores_config + 1):
                            colunas_acro.extend([f'E{j} Torq ft lbs', f'E{j} ITT', f'E{j} NG', f'E{j} NP %'])
                        colunas_acro = [c for c in colunas_acro if c in df.columns]
                        exibir_tabela_grande(df_acro.sort_values(by='NormAc', ascending=False).head(20)[colunas_acro].copy())
                    else: 
                        st.info("Nenhuma manobra severa detectada.")
                        
            st.divider()
            st.subheader("✈️ Trajetória 3D de Acrobacias (Cor por Velocidade)")
            if not voo_ativo.empty and 'Latitude' in voo_ativo.columns and 'Longitude' in voo_ativo.columns:
                df_3d = voo_ativo.dropna(subset=['Latitude', 'Longitude', 'AltMSL'])
                
                fig_3d = go.Figure(data=[go.Scatter3d(
                    x=df_3d['Longitude'], y=df_3d['Latitude'], z=df_3d['AltMSL'], 
                    mode='lines', 
                    line=dict(color=df_3d['IAS'], colorscale='Turbo', width=4, colorbar=dict(title="Velocidade IAS", len=0.5, x=1.1))
                )])
                fig_3d.update_layout(template="plotly_dark", height=600, margin=dict(l=0, r=0, b=0, t=20), scene=dict(aspectmode='auto'))
                st.plotly_chart(fig_3d, use_container_width=True)

            st.divider()
            st.markdown("📝 **Observações - Avaliador de Manobras e Pouso**")
            st.session_state.obs_manobras = st.text_area("Escreva aqui:", value=st.session_state.obs_manobras, key='k_manob')
            if st.button("💾 Salvar Comentário", key="b_manob"): 
                st.success("Salvo com sucesso!")

        with aba_relatorio:
            st.header("📝 Editor do Relatório Final")
            st.markdown("Confirme ou edite os dados extraídos antes de gerar o documento Word oficial.")

            st.subheader("1. Edição do Logbook da Missão")
            col_t1, col_t2, col_t3, col_t4, col_t5 = st.columns(5)
            with col_t1: temp_ac = st.text_input("Acionamento:", tempos.get('acionamento','-'))
            with col_t2: temp_dec = st.text_input("Decolagem:", tempos.get('decolagem','-'))
            with col_t3: temp_pou = st.text_input("Pouso:", tempos.get('pouso','-'))
            with col_t4: temp_cor = st.text_input("Corte:", tempos.get('corte','-'))
            with col_t5: temp_voo = st.text_input("Tempo de Voo:", tempos.get('t_voo','-'))
            
            tempos_editados = {
                'acionamento': temp_ac,
                'decolagem': temp_dec,
                'pouso': temp_pou,
                'corte': temp_cor,
                't_voo': temp_voo
            }

            st.subheader("2. Edição de Alertas do Voo")
            st.caption("Pode apagar linhas falsas ou reescrever motivos. (Mantenha o texto 🔴 ou 🟡 para o Word aplicar a cor).")
            alertas_txt = st.text_area("Lista de Alertas (Um por linha):", value="\n".join(alertas), height=150)

            st.info("As observações (Campo 3) digitadas e salvas nas outras abas serão incluídas automaticamente no documento final.")
            obs_solo_atual = {
                "Resumo de Painel e Rota": st.session_state.obs_resumo,
                "Inspetoria e Graficos": st.session_state.obs_graficos,
                "Avaliador de Manobras e Pouso": st.session_state.obs_manobras
            }

            st.divider()
            st.subheader("🖨️ Gerar Documento Word")
            docx_bytes_solo_editado = gerar_docx_solo_editado(tempos_editados, alertas_txt, obs_solo_atual)
            st.download_button(
                label="💾 BAIXAR RELATÓRIO OFICIAL (.DOCX)", 
                data=docx_bytes_solo_editado, 
                file_name="Inspetoria_Relatorio.docx", 
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
                use_container_width=True
            )
