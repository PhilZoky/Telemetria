import customtkinter as ctk
import sqlite3
import json
from datetime import datetime
import os
import socket
import threading
from flask import Flask, render_template_string, request, redirect, url_for
import qrcode
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from PIL import Image

def configurar_banco():
    conn = sqlite3.connect('ceravelle_app.db')
    cursor = conn.cursor()
    cursor.execute('''CREATE TABLE IF NOT EXISTS receitas (id INTEGER PRIMARY KEY AUTOINCREMENT, nome TEXT UNIQUE, dados TEXT)''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS testes_completos (id INTEGER PRIMARY KEY AUTOINCREMENT, nome_teste TEXT UNIQUE, dados TEXT)''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS config (id INTEGER PRIMARY KEY, nome_marca TEXT, licenca_ativa TEXT)''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS estoque (id INTEGER PRIMARY KEY AUTOINCREMENT, item TEXT, quantidade REAL, unidade TEXT, validade TEXT)''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS producao (id INTEGER PRIMARY KEY AUTOINCREMENT, data TEXT, receita TEXT, quantidade INTEGER, custo_total REAL, faturamento REAL, lucro REAL)''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS encomendas (id INTEGER PRIMARY KEY AUTOINCREMENT, cliente TEXT, telefone TEXT, produto TEXT, quantidade INTEGER, status TEXT, data TEXT, canal TEXT)''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS clientes (id INTEGER PRIMARY KEY AUTOINCREMENT, nome TEXT, telefone TEXT, email TEXT, aniversario TEXT)''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS historico_compras (id INTEGER PRIMARY KEY AUTOINCREMENT, cliente_id INTEGER, produto TEXT, quantidade INTEGER, valor_gasto REAL, data_hora TEXT)''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS despesas (id INTEGER PRIMARY KEY AUTOINCREMENT, descricao TEXT, valor REAL)''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS orcamentos (id INTEGER PRIMARY KEY AUTOINCREMENT, cliente TEXT, valor REAL, data TEXT)''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS kits (id INTEGER PRIMARY KEY AUTOINCREMENT, nome_kit TEXT, descricao TEXT, preco REAL)''')
    
    try:
        cursor.execute("ALTER TABLE encomendas ADD COLUMN canal TEXT")
    except sqlite3.OperationalError:
        pass

    cursor.execute("INSERT OR IGNORE INTO config (id, nome_marca, licenca_ativa) VALUES (1, 'Ceravelle', 'PRO')")
    conn.commit()
    conn.close()

configurar_banco()

# --- CONFIGURAÇÃO DO SERVIDOR FLASK PARA O TELEMÓVEL ---
flask_app = Flask("CeravelleMobile")

def get_local_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except:
        return "127.0.0.1"

HTML_SUCESSO = """
<!DOCTYPE html>
<html lang="pt">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sucesso - Ceravelle</title>
    <style>
        body { background-color: #F7F2EC; color: #4A3B32; font-family: Arial, sans-serif; text-align: center; padding: 50px; }
        .box { background: #EFE5DC; padding: 30px; border-radius: 10px; border: 1px solid #C5A059; display: inline-block; }
        h2 { color: #C5A059; }
    </style>
</head>
<body>
    <div class="box">
        <h2>Sucesso!</h2>
        <p>Os dados do teste <b>{{ nome_teste }}</b> foram enviados e salvos no aplicativo do computador com sucesso!</p>
    </div>
</body>
</html>
"""

HTML_FORMULARIO = """
<!DOCTYPE html>
<html lang="pt">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Ceravelle - Registo de Teste</title>
    <style>
        body { background-color: #F7F2EC; color: #4A3B32; font-family: Arial, sans-serif; padding: 15px; margin: 0; }
        .container { max-width: 600px; margin: auto; background: #EFE5DC; padding: 15px; border-radius: 10px; border: 1px solid #C5A059; }
        h2 { text-align: center; color: #C5A059; font-family: 'Times New Roman', serif; margin-bottom: 5px; }
        h3 { color: #4A3B32; border-bottom: 1px solid #C5A059; padding-bottom: 5px; margin-top: 20px; font-size: 15px; }
        label { font-weight: bold; display: block; margin-top: 8px; font-size: 12px; }
        input, textarea { width: 100%; padding: 8px; margin-top: 2px; border: 1px solid #C5A059; border-radius: 4px; box-sizing: border-box; background: white; font-size: 13px; }
        .row { display: flex; gap: 10px; }
        .col { flex: 1; }
        table { width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 12px; }
        th, td { border: 1px solid #C5A059; padding: 6px; text-align: center; background: white; }
        th { background-color: #EFE5DC; color: #4A3B32; font-size: 11px; }
        td:first-child { text-align: left; font-weight: bold; font-size: 11px; background-color: #FAF5F0; }
        input.grid-input { width: 100%; border: none; text-align: center; padding: 4px; font-size: 13px; background: transparent; }
        textarea { height: 100px; }
        button { background-color: #C5A059; color: white; border: none; padding: 14px; width: 100%; border-radius: 5px; font-size: 16px; font-weight: bold; margin-top: 20px; cursor: pointer; }
        button:hover { background-color: #B3922B; }
    </style>
</head>
<body>
    <div class="container">
        <h2>CERAVELLE</h2>
        <p style="text-align: center; font-size: 13px; margin-top: 0;">Ficha de Teste: <b>{{ nome_teste }}</b></p>
        
        <form method="POST">
            <h3>Cabeçalho e Horários</h3>
            <div class="row">
                <div class="col"><label>Data:</label><input type="text" name="data" value="{{ dados.get('t_data', '') }}"></div>
                <div class="col"><label>Nº do Teste:</label><input type="text" name="num" value="{{ dados.get('t_num', '') }}"></div>
            </div>
            <div class="row">
                <div class="col"><label>Horário Início:</label><input type="text" name="hora_ini" value="{{ dados.get('t_hora_ini', '') }}"></div>
                <div class="col"><label>Horário Fim:</label><input type="text" name="hora_fim" value="{{ dados.get('t_hora_fim', '') }}"></div>
            </div>

            <h3>1. Formulação e Insumos</h3>
            <div class="row">
                <div class="col"><label>Tipo de Cera:</label><input type="text" name="cera_tipo" value="{{ dados.get('t_cera_tipo', '') }}"></div>
                <div class="col"><label>Qnt Cera (g):</label><input type="text" name="cera_qnt" value="{{ dados.get('t_cera_qnt', '') }}"></div>
            </div>
            <div class="row">
                <div class="col"><label>Temp. Derret. (°C):</label><input type="text" name="cera_temp" value="{{ dados.get('t_cera_temp', '') }}"></div>
                <div class="col"><label>Tipo de Pavio:</label><input type="text" name="pavio_tipo" value="{{ dados.get('t_pavio_tipo', '') }}"></div>
            </div>
            <div class="row">
                <div class="col"><label>Qnt Pavio:</label><input type="text" name="pavio_qnt" value="{{ dados.get('t_pavio_qnt', '') }}"></div>
                <div class="col"><label>Essência:</label><input type="text" name="essencia" value="{{ dados.get('t_essencia', '') }}"></div>
            </div>
            <div class="row">
                <div class="col"><label>Essência (%):</label><input type="text" name="essencia_perc" value="{{ dados.get('t_essencia_perc', '') }}"></div>
                <div class="col"><label>Essência (g):</label><input type="text" name="essencia_g" value="{{ dados.get('t_essencia_g', '') }}"></div>
            </div>
            <div class="row">
                <div class="col"><label>Temp. Mistura (°C):</label><input type="text" name="essencia_temp" value="{{ dados.get('t_essencia_temp', '') }}"></div>
                <div class="col"><label>Corante:</label><input type="text" name="corante" value="{{ dados.get('t_corante', '') }}"></div>
            </div>
            <div class="row">
                <div class="col"><label>Qnt Corante (g):</label><input type="text" name="corante_qnt" value="{{ dados.get('t_corante_qnt', '') }}"></div>
                <div class="col"><label>Temp. Envase (°C):</label><input type="text" name="envase_temp" value="{{ dados.get('t_envase_temp', '') }}"></div>
            </div>
            <div class="row">
                <div class="col"><label>Peso Total (g):</label><input type="text" name="peso_total" value="{{ dados.get('t_peso_total', '') }}"></div>
                <div class="col"><label>Cold Throw (0-10):</label><input type="text" name="cold_throw" value="{{ dados.get('t_cold_throw', '') }}"></div>
            </div>
            <div class="row">
                <div class="col"><label>Cômodo Teste:</label><input type="text" name="comodo" value="{{ dados.get('t_comodo', '') }}"></div>
                <div class="col"><label>Modelo Recipiente:</label><input type="text" name="modelo" value="{{ dados.get('t_modelo', '') }}"></div>
            </div>
            <label>Diâmetro Recipiente (cm):</label>
            <input type="text" name="diametro" value="{{ dados.get('t_diametro', '') }}">

            <h3>2. Grelha de Acompanhamento (1h a 4h)</h3>
            <div style="overflow-x:auto;">
                <table>
                    <tr>
                        <th>Parâmetro</th>
                        <th>1h</th>
                        <th>2h</th>
                        <th>3h</th>
                        <th>4h</th>
                    </tr>
                    {% set parametros = [
                        "Altura chama (cm)", "Chama estável?", "Aroma percebido (0-10)", 
                        "Fuligem (Sim/Não)", "Cogumelo", "Túnel", 
                        "Prof. piscina fusão", "Piscina até a borda", "Estabilidade Recipiente"
                    ] %}
                    {% for i in range(9) %}
                    <tr>
                        <td>{{ parametros[i] }}</td>
                        {% for j in range(4) %}
                        <td>
                            <input type="text" class="grid-input" name="t_grid_{{ i }}_{{ j }}" value="{{ dados.get('t_grid_' ~ i ~ '_' ~ j, '') }}">
                        </td>
                        {% endfor %}
                    </tr>
                    {% endfor %}
                </table>
            </div>

            <h3>3. Observações Finais e Conclusão</h3>
            <label>Tempo de Queima Total:</label>
            <input type="text" name="queima_total" value="{{ dados.get('t_queima_total', '') }}">
            
            <label>Piscina Completa?</label>
            <input type="text" name="piscina_comp" value="{{ dados.get('t_piscina_comp', '') }}">
            
            <label>Altura do Pavio (Início):</label>
            <input type="text" name="pavio_ini" value="{{ dados.get('t_pavio_ini', '') }}">
            
            <label>Formou Piscina às:</label>
            <input type="text" name="piscina_hora" value="{{ dados.get('t_piscina_hora', '') }}">
            
            <label>Observações Adicionais e Conclusões:</label>
            <textarea name="obs_adicionais">{{ dados.get('t_obs_adicionais', '') }}</textarea>
            
            <button type="submit">Guardar Tudo no Computador</button>
        </form>
    </div>
</body>
</html>
"""

@flask_app.route('/preencher/<path:nome_teste>', methods=['GET', 'POST'])
def preencher_via_celular(nome_teste):
    conn = sqlite3.connect('ceravelle_app.db')
    if request.method == 'POST':
        resultado = conn.execute("SELECT dados FROM testes_completos WHERE nome_teste = ?", (nome_teste,)).fetchone()
        dados = json.loads(resultado[0]) if resultado else {}
        
        dados['t_data'] = request.form.get('data', '')
        dados['t_num'] = request.form.get('num', '')
        dados['t_hora_ini'] = request.form.get('hora_ini', '')
        dados['t_hora_fim'] = request.form.get('hora_fim', '')

        dados['t_cera_tipo'] = request.form.get('cera_tipo', '')
        dados['t_cera_qnt'] = request.form.get('cera_qnt', '')
        dados['t_cera_temp'] = request.form.get('cera_temp', '')
        dados['t_pavio_tipo'] = request.form.get('pavio_tipo', '')
        dados['t_pavio_qnt'] = request.form.get('pavio_qnt', '')
        dados['t_essencia'] = request.form.get('essencia', '')
        dados['t_essencia_perc'] = request.form.get('essencia_perc', '')
        dados['t_essencia_g'] = request.form.get('essencia_g', '')
        dados['t_essencia_temp'] = request.form.get('essencia_temp', '')
        dados['t_corante'] = request.form.get('corante', '')
        dados['t_corante_qnt'] = request.form.get('corante_qnt', '')
        dados['t_envase_temp'] = request.form.get('envase_temp', '')
        dados['t_peso_total'] = request.form.get('peso_total', '')
        dados['t_cold_throw'] = request.form.get('cold_throw', '')
        dados['t_comodo'] = request.form.get('comodo', '')
        dados['t_modelo'] = request.form.get('modelo', '')
        dados['t_diametro'] = request.form.get('diametro', '')

        dados['t_queima_total'] = request.form.get('queima_total', '')
        dados['t_piscina_comp'] = request.form.get('piscina_comp', '')
        dados['t_pavio_ini'] = request.form.get('pavio_ini', '')
        dados['t_piscina_hora'] = request.form.get('piscina_hora', '')
        dados['t_obs_adicionais'] = request.form.get('obs_adicionais', '')

        for i in range(9):
            for j in range(4):
                chave = f"t_grid_{i}_{j}"
                dados[chave] = request.form.get(chave, '')

        conn.execute("INSERT OR REPLACE INTO testes_completos (nome_teste, dados) VALUES (?, ?)", (nome_teste, json.dumps(dados)))
        conn.commit()
        conn.close()
        return render_template_string(HTML_SUCESSO, nome_teste=nome_teste)

    resultado = conn.execute("SELECT dados FROM testes_completos WHERE nome_teste = ?", (nome_teste,)).fetchone()
    conn.close()
    dados = json.loads(resultado[0]) if resultado else {}

    return render_template_string(HTML_FORMULARIO, nome_teste=nome_teste, dados=dados)

def iniciar_servidor_flask():
    flask_app.run(host='0.0.0.0', port=5000, debug=False, use_reloader=False)

threading.Thread(target=iniciar_servidor_flask, daemon=True).start()

# --- PALETA DE CORES DA LOGO CERAVELLE ---
COR_BG = "#F7F2EC"          
COR_SECUNDARIA = "#EFE5DC"   
COR_TEXTO = "#4A3B32"        
COR_DESTAQUE = "#C5A059"     
COR_BOTAO = "#D9C3A9"        

ctk.set_appearance_mode("light")

class SistemaVelasApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.campos_dados = {} 
        self.campos_liquidos = {}
        self.campos_teste = {}
        self.campos_proporcao = {}
        
        self.title("Ceravelle - Sistema de Gestão")
        self.geometry("1350x950")
        self.configure(fg_color=COR_BG)

        self.criar_cabecalho_marca()

        self.tabs = ctk.CTkTabview(self, fg_color=COR_SECUNDARIA, segmented_button_selected_color=COR_DESTAQUE, segmented_button_selected_hover_color="#B3922B", text_color=COR_TEXTO)
        self.tabs.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.tab_dash = self.tabs.add("Dashboard")
        self.tab_prop = self.tabs.add("Proporção")
        self.tab_calc = self.tabs.add("Velas")
        self.tab_liquidos = self.tabs.add("Líquidos")
        self.tab_testes = self.tabs.add("Testes")
        self.tab_estoque = self.tabs.add("Estoque")
        self.tab_despesas = self.tabs.add("Despesas Fixas")
        self.tab_kits = self.tabs.add("Kits & Combos")
        self.tab_producao = self.tabs.add("Produção & Lotes")
        self.tab_pedidos = self.tabs.add("Encomendas & CRM")
        self.tab_clientes = self.tabs.add("Clientes & Fidelidade")
        self.tab_sistema = self.tabs.add("Orçamentos & Relatórios")

        self.construir_aba_dashboard()
        self.construir_aba_proporcao()
        self.construir_aba_calculadora()
        self.construir_aba_liquidos()
        self.construir_aba_testes()
        self.construir_aba_estoque()
        self.construir_aba_despesas()
        self.construir_aba_kits()
        self.construir_aba_producao()
        self.construir_aba_pedidos()
        self.construir_aba_clientes()
        self.construir_aba_sistema()

    def criar_cabecalho_marca(self):
        f_top = ctk.CTkFrame(self, fg_color="#F6EFE5", height=135)
        f_top.pack(fill="x", padx=0, pady=(0, 0))
        f_top.pack_propagate(False)
        
        if os.path.exists("Ceravelle.png"):
            try:
                img_pil = Image.open("Ceravelle.png")
                self.logo_esq = ctk.CTkImage(light_image=img_pil, dark_image=img_pil, size=(130, 130))
                lbl_esq = ctk.CTkLabel(f_top, image=self.logo_esq, text="")
                lbl_esq.pack(side="left", padx=25, pady=2)

                self.logo_dir = ctk.CTkImage(light_image=img_pil, dark_image=img_pil, size=(130, 130))
                lbl_dir = ctk.CTkLabel(f_top, image=self.logo_dir, text="")
                lbl_dir.pack(side="right", padx=25, pady=2)
            except Exception:
                pass

        lbl_titulo = ctk.CTkLabel(f_top, text="C    E    R    A    V    E    L    L    E", font=("Palatino Linotype", 44, "bold"), text_color="#C5A059")
        lbl_titulo.pack(expand=True, pady=38)

    def criar_campo(self, parent, texto, chave, row, col, dicionario):
        lbl = ctk.CTkLabel(parent, text=texto, text_color=COR_TEXTO, font=("Trebuchet MS", 13))
        lbl.grid(row=row, column=col, padx=10, pady=5, sticky="e")
        entry = ctk.CTkEntry(parent, placeholder_text="0", fg_color="white", border_color=COR_DESTAQUE, text_color=COR_TEXTO, font=("Trebuchet MS", 12))
        entry.grid(row=row, column=col+1, padx=10, pady=5, sticky="w")
        dicionario[chave] = entry
        return entry

    def obter_valor(self, campo):
        texto = campo.get().strip().replace(',', '.')
        return float(texto) if texto else 0.0

    # ==========================================
    # ABA: PROPORÇÃO DE PESO
    # ==========================================
    def construir_aba_proporcao(self):
        scroll_prop = ctk.CTkScrollableFrame(self.tab_prop, fg_color=COR_BG)
        scroll_prop.pack(fill="both", expand=True, padx=10, pady=10)

        ctk.CTkLabel(scroll_prop, text="Calculadora Rápida de Proporção (Bancada)", font=("Palatino Linotype", 22, "bold"), text_color=COR_TEXTO).pack(pady=15)
        ctk.CTkLabel(scroll_prop, text="Escolha o método de cálculo e informe os dados para ver exatamente o que separar:", font=("Trebuchet MS", 13), text_color=COR_TEXTO).pack(pady=5)

        self.modo_prop = ctk.StringVar(value="Peso Total da Vela (Vidro)")
        seletor_modo_prop = ctk.CTkSegmentedButton(
            scroll_prop, 
            values=["Peso Total da Vela (Vidro)", "Peso da Cera a Derreter"], 
            variable=self.modo_prop, 
            command=self.atualizar_rotulo_proporcao, 
            selected_color=COR_DESTAQUE, 
            text_color=COR_TEXTO
        )
        seletor_modo_prop.pack(pady=15)

        f_inputs = ctk.CTkFrame(scroll_prop, fg_color=COR_SECUNDARIA, border_color=COR_DESTAQUE, border_width=1)
        f_inputs.pack(pady=10, padx=20, fill="x")

        self.lbl_prop_input = ctk.CTkLabel(f_inputs, text="Peso Total Desejado da Vela (g):", font=("Trebuchet MS", 13, "bold"), text_color=COR_TEXTO)
        self.lbl_prop_input.grid(row=0, column=0, padx=15, pady=15, sticky="e")
        
        self.entry_prop_valor = ctk.CTkEntry(f_inputs, placeholder_text="Ex: 180", width=160, fg_color="white", border_color=COR_DESTAQUE, text_color=COR_TEXTO, font=("Trebuchet MS", 14))
        self.entry_prop_valor.grid(row=0, column=1, padx=15, pady=15, sticky="w")
        self.campos_proporcao["valor_base"] = self.entry_prop_valor

        ctk.CTkLabel(f_inputs, text="Carga de Essência (%):", font=("Trebuchet MS", 13, "bold"), text_color=COR_TEXTO).grid(row=1, column=0, padx=15, pady=15, sticky="e")
        self.entry_prop_carga = ctk.CTkEntry(f_inputs, placeholder_text="Ex: 14", width=160, fg_color="white", border_color=COR_DESTAQUE, text_color=COR_TEXTO, font=("Trebuchet MS", 14))
        self.entry_prop_carga.grid(row=1, column=1, padx=15, pady=15, sticky="w")
        self.entry_prop_carga.insert(0, "14")
        self.campos_proporcao["carga"] = self.entry_prop_carga

        ctk.CTkButton(scroll_prop, text="Calcular Proporção Exata", command=self.calcular_proporcao_rapida, font=("Trebuchet MS", 14, "bold"), fg_color=COR_DESTAQUE, text_color="white", height=40, width=250).pack(pady=15)

        self.frame_res_prop = ctk.CTkFrame(scroll_prop, fg_color="#FAF5F0", border_color=COR_DESTAQUE, border_width=2)
        self.frame_res_prop.pack(pady=10, padx=20, fill="x")

        ctk.CTkLabel(self.frame_res_prop, text="Resultados para a Pesagem no Ateliê", font=("Palatino Linotype", 18, "bold"), text_color=COR_TEXTO).pack(pady=(15, 10))
        
        self.lbl_res_cera = ctk.CTkLabel(self.frame_res_prop, text="• Peso da Cera a Separar: 0.00 g", font=("Trebuchet MS", 15, "bold"), text_color=COR_TEXTO)
        self.lbl_res_cera.pack(anchor="w", padx=30, pady=5)

        self.lbl_res_essencia = ctk.CTkLabel(self.frame_res_prop, text="• Peso da Essência a Separar: 0.00 g", font=("Trebuchet MS", 15, "bold"), text_color=COR_TEXTO)
        self.lbl_res_essencia.pack(anchor="w", padx=30, pady=5)

        self.lbl_res_total_final = ctk.CTkLabel(self.frame_res_prop, text="• Peso Total Final no Pote: 0.00 g", font=("Trebuchet MS", 15, "bold"), text_color="#C5A059")
        self.lbl_res_total_final.pack(anchor="w", padx=30, pady=(5, 15))

    def atualizar_rotulo_proporcao(self, modo):
        if modo == "Peso Total da Vela (Vidro)":
            self.lbl_prop_input.configure(text="Peso Total Desejado da Vela (g):")
        else:
            self.lbl_prop_input.configure(text="Peso da Cera a Derreter (g):")

    def calcular_proporcao_rapida(self):
        try:
            val_base = self.obter_valor(self.campos_proporcao["valor_base"])
            carga_perc = self.obter_valor(self.campos_proporcao["carga"]) / 100
            modo = self.modo_prop.get()

            if modo == "Peso Total da Vela (Vidro)":
                essencia = val_base * carga_perc
                cera = val_base - essencia
                total_final = val_base
            else:
                cera = val_base
                essencia = cera * carga_perc
                total_final = cera + essencia

            self.lbl_res_cera.configure(text=f"• Peso da Cera a Separar: {cera:.2f} g")
            self.lbl_res_essencia.configure(text=f"• Peso da Essência a Separar: {essencia:.2f} g")
            self.lbl_res_total_final.configure(text=f"• Peso Total Final no Pote: {total_final:.2f} g")
        except Exception:
            self.lbl_res_cera.configure(text="• Erro nos valores informados.")
            self.lbl_res_essencia.configure(text="")
            self.lbl_res_total_final.configure(text="")

    # ==========================================
    # ABA 0: DASHBOARD
    # ==========================================
    def construir_aba_dashboard(self):
        self.scroll_dash = ctk.CTkScrollableFrame(self.tab_dash, fg_color=COR_BG)
        self.scroll_dash.pack(fill="both", expand=True, padx=10, pady=10)

        ctk.CTkLabel(self.scroll_dash, text="Painel de Inteligência - Ceravelle", font=("Trebuchet MS", 22, "bold"), text_color=COR_TEXTO).pack(pady=15)

        f_cards = ctk.CTkFrame(self.scroll_dash, fg_color="transparent")
        f_cards.pack(fill="x", pady=5)

        self.card_receitas = ctk.CTkFrame(f_cards, width=220, height=90, fg_color=COR_SECUNDARIA, border_color=COR_DESTAQUE, border_width=1)
        self.card_receitas.pack(side="left", padx=5, expand=True, fill="x")
        self.lbl_dash_receitas = ctk.CTkLabel(self.card_receitas, text="Receitas: 0", font=("Trebuchet MS", 13, "bold"), text_color=COR_TEXTO)
        self.lbl_dash_receitas.pack(pady=20)

        self.card_producao = ctk.CTkFrame(f_cards, width=220, height=90, fg_color=COR_SECUNDARIA, border_color=COR_DESTAQUE, border_width=1)
        self.card_producao.pack(side="left", padx=5, expand=True, fill="x")
        self.lbl_dash_producao = ctk.CTkLabel(self.card_producao, text="Faturamento: R$ 0.00", font=("Trebuchet MS", 13, "bold"), text_color=COR_TEXTO)
        self.lbl_dash_producao.pack(pady=20)

        self.card_clientes = ctk.CTkFrame(f_cards, width=220, height=90, fg_color=COR_SECUNDARIA, border_color=COR_DESTAQUE, border_width=1)
        self.card_clientes.pack(side="left", padx=5, expand=True, fill="x")
        self.lbl_dash_clientes = ctk.CTkLabel(self.card_clientes, text="Clientes: 0", font=("Trebuchet MS", 13, "bold"), text_color=COR_TEXTO)
        self.lbl_dash_clientes.pack(pady=20)

        self.card_alerta = ctk.CTkFrame(self.scroll_dash, fg_color="#D4EFDF", border_color="#5CB85C", border_width=1)
        self.card_alerta.pack(fill="x", padx=5, pady=10)
        self.lbl_alerta_estoque = ctk.CTkLabel(self.card_alerta, text="✅ Sistema Operacional: Estoque e validades regulares.", font=("Trebuchet MS", 12, "bold"), text_color="#1E8449")
        self.lbl_alerta_estoque.pack(pady=10, padx=10)

        ctk.CTkButton(self.scroll_dash, text="Atualizar Painel e Indicadores", command=self.atualizar_dashboard, fg_color=COR_DESTAQUE, text_color="white", font=("Trebuchet MS", 12, "bold")).pack(pady=10)
        
        self.lbl_grafico = ctk.CTkLabel(self.scroll_dash, text="")
        self.lbl_grafico.pack(pady=10)
        
        self.atualizar_dashboard()

    def atualizar_dashboard(self):
        conn = sqlite3.connect('ceravelle_app.db')
        total_rec = conn.execute("SELECT COUNT(*) FROM receitas").fetchone()[0]
        total_cli = conn.execute("SELECT COUNT(*) FROM clientes").fetchone()[0]
        prod_info = conn.execute("SELECT SUM(faturamento), SUM(lucro) FROM producao").fetchone()
        tot_fat = prod_info[0] if prod_info and prod_info[0] else 0.0
        
        baixos = conn.execute("SELECT item, quantidade, unidade FROM estoque WHERE quantidade < 200").fetchall()
        lotes = conn.execute("SELECT receita, faturamento, lucro FROM producao ORDER BY id DESC LIMIT 5").fetchall()
        conn.close()

        self.lbl_dash_receitas.configure(text=f"Receitas: {total_rec}")
        self.lbl_dash_producao.configure(text=f"Faturamento: R$ {tot_fat:.2f}")
        self.lbl_dash_clientes.configure(text=f"Clientes: {total_cli}")

        if baixos:
            nomes_baixos = ", ".join([f"{b[0]} ({b[1]}{b[2] or ''})" for b in baixos])
            self.lbl_alerta_estoque.configure(text=f"⚠️ ALERTA DE ESTOQUE BAIXO: Itens precisam de reposição: {nomes_baixos}")
            self.card_alerta.configure(fg_color="#FADBD8")
        else:
            self.lbl_alerta_estoque.configure(text="✅ Estoque regular: Nenhum insumo crítico no momento.")
            self.card_alerta.configure(fg_color="#D4EFDF")

        if lotes:
            nomes = [l[0] for l in lotes[::-1]]
            fats = [l[1] for l in lotes[::-1]]
            lucros = [l[2] for l in lotes[::-1]]

            fig, ax = plt.subplots(figsize=(7, 3.2))
            fig.patch.set_facecolor('#F7F2EC')
            ax.set_facecolor('#EFE5DC')

            x = range(len(nomes))
            width = 0.35
            ax.bar([i - width/2 for i in x], fats, width, label='Faturamento', color='#C5A059')
            ax.bar([i + width/2 for i in x], lucros, width, label='Lucro Líquido', color='#4A3B32')

            ax.set_xticks(x)
            ax.set_xticklabels(nomes, color='#4A3B32')
            ax.tick_params(colors='#4A3B32')
            ax.legend(facecolor='#EFE5DC', edgecolor='#C5A059', labelcolor='#4A3B32')
            plt.tight_layout()

            grafico_path = "grafico_temp.png"
            plt.savefig(grafico_path, dpi=100, facecolor=fig.get_facecolor(), edgecolor='none')
            plt.close(fig)

            img_pil = Image.open(grafico_path)
            self.chart_image = ctk.CTkImage(light_image=img_pil, dark_image=img_pil, size=(700, 320))
            self.lbl_grafico.configure(image=self.chart_image)
        else:
            self.lbl_grafico.configure(text="Registe lotes de produção para visualizar o gráfico de desempenho.", text_color=COR_TEXTO)

    # ==========================================
    # ABA 1: VELAS
    # ==========================================
    def construir_aba_calculadora(self):
        self.frame_db = ctk.CTkFrame(self.tab_calc, fg_color=COR_SECUNDARIA)
        self.frame_db.pack(pady=5, fill="x")
        
        self.entry_nome_receita = ctk.CTkEntry(self.frame_db, placeholder_text="Nome da Vela", width=180, fg_color="white", border_color=COR_DESTAQUE, text_color=COR_TEXTO, font=("Trebuchet MS", 12))
        self.entry_nome_receita.pack(side="left", padx=5, pady=10)
        
        ctk.CTkButton(self.frame_db, text="Salvar/Atualizar", command=self.salvar_receita, fg_color=COR_DESTAQUE, text_color="white", font=("Trebuchet MS", 11, "bold"), width=110).pack(side="left", padx=2)
        ctk.CTkButton(self.frame_db, text="Excluir Receita", command=self.excluir_receita, fg_color="#C9302C", text_color="white", font=("Trebuchet MS", 11, "bold"), width=105).pack(side="left", padx=2)
        ctk.CTkButton(self.frame_db, text="Ficha Técnica (PDF)", command=self.gerar_ficha_tecnica_pdf, fg_color="#2980B9", text_color="white", font=("Trebuchet MS", 11, "bold"), width=130).pack(side="left", padx=2)

        self.seletor_receitas = ctk.CTkOptionMenu(self.frame_db, values=["Selecione..."], command=self.carregar_receita, fg_color=COR_BOTAO, button_color=COR_DESTAQUE, text_color=COR_TEXTO, width=150)
        self.seletor_receitas.pack(side="right", padx=10)
        self.atualizar_lista_receitas()

        self.scroll_calc = ctk.CTkScrollableFrame(self.tab_calc, fg_color=COR_BG)
        self.scroll_calc.pack(fill="both", expand=True, pady=5)

        ctk.CTkLabel(self.scroll_calc, text="1. Formulação da Massa", font=("Trebuchet MS", 14, "bold"), text_color=COR_TEXTO).grid(row=0, column=0, columnspan=4, pady=(10, 5), sticky="w")
        
        self.modo_calculo = ctk.StringVar(value="Peso Total do Vidro")
        self.seletor_modo = ctk.CTkSegmentedButton(self.scroll_calc, values=["Peso Total do Vidro", "Peso da Cera"], variable=self.modo_calculo, command=self.atualizar_texto_peso, selected_color=COR_DESTAQUE, text_color=COR_TEXTO)
        self.seletor_modo.grid(row=1, column=0, columnspan=4, pady=10)

        self.lbl_peso = ctk.CTkLabel(self.scroll_calc, text="Peso Total do Vidro (g):", text_color=COR_TEXTO)
        self.lbl_peso.grid(row=2, column=0, padx=10, pady=5, sticky="e")
        self.entry_peso = ctk.CTkEntry(self.scroll_calc, placeholder_text="Ex: 200", fg_color="white", border_color=COR_DESTAQUE, text_color=COR_TEXTO, font=("Trebuchet MS", 12))
        self.entry_peso.grid(row=2, column=1, padx=10, pady=5, sticky="w")
        self.campos_dados["peso"] = self.entry_peso

        self.criar_campo(self.scroll_calc, "Carga Essência (%):", "carga", 2, 2, self.campos_dados)
        self.criar_campo(self.scroll_calc, "Preço Cera (R$):", "preco_cera", 3, 0, self.campos_dados)
        self.criar_campo(self.scroll_calc, "Qtd. Comprada (g):", "qtd_cera", 3, 2, self.campos_dados)
        self.criar_campo(self.scroll_calc, "Preço Essência (R$):", "preco_essencia", 4, 0, self.campos_dados)
        self.criar_campo(self.scroll_calc, "Qtd. Comprada (g/ml):", "qtd_essencia", 4, 2, self.campos_dados)

        ctk.CTkLabel(self.scroll_calc, text="2. Materiais e Embalagem (R$ Unitário)", font=("Trebuchet MS", 14, "bold"), text_color=COR_TEXTO).grid(row=5, column=0, columnspan=4, pady=(15, 5), sticky="w")
        self.criar_campo(self.scroll_calc, "Vidro + Tampa:", "vidro", 6, 0, self.campos_dados)
        self.criar_campo(self.scroll_calc, "Pavio + Ilhós:", "pavio", 6, 2, self.campos_dados)
        self.criar_campo(self.scroll_calc, "Rótulos/Adesivos:", "rotulo", 7, 0, self.campos_dados)
        self.criar_campo(self.scroll_calc, "Caixa/Sacola:", "caixa", 7, 2, self.campos_dados)
        self.criar_campo(self.scroll_calc, "Outros (Fita, Cola):", "outros", 8, 0, self.campos_dados)

        ctk.CTkLabel(self.scroll_calc, text="3. Mão de Obra", font=("Trebuchet MS", 14, "bold"), text_color=COR_TEXTO).grid(row=9, column=0, columnspan=4, pady=(15, 5), sticky="w")
        self.criar_campo(self.scroll_calc, "Salário Desejado (R$):", "salario", 10, 0, self.campos_dados)
        self.criar_campo(self.scroll_calc, "Horas Trabalhadas/Mês:", "horas_mes", 10, 2, self.campos_dados)
        self.criar_campo(self.scroll_calc, "Tempo nesta Vela (Min):", "tempo_vela", 11, 0, self.campos_dados)

        ctk.CTkLabel(self.scroll_calc, text="4. Precificação e Taxas", font=("Trebuchet MS", 14, "bold"), text_color=COR_TEXTO).grid(row=12, column=0, columnspan=4, pady=(15, 5), sticky="w")
        self.criar_campo(self.scroll_calc, "Markup Varejo (Ex: 2.5):", "markup", 13, 0, self.campos_dados)
        self.criar_campo(self.scroll_calc, "Markup Atacado (Ex: 1.5):", "markup_atacado", 13, 2, self.campos_dados)
        self.criar_campo(self.scroll_calc, "Taxa Cartão/Plat. (%):", "taxa", 14, 0, self.campos_dados)

        self.frame_resultados = ctk.CTkFrame(self.tab_calc, fg_color=COR_SECUNDARIA)
        self.frame_resultados.pack(pady=5, fill="x")

        ctk.CTkButton(self.frame_resultados, text="Processar Custos e Precificação", command=self.calcular_velas, font=("Trebuchet MS", 14, "bold"), fg_color=COR_DESTAQUE, text_color="white", hover_color="#B3922B").pack(pady=10)
        self.lbl_res_custo = ctk.CTkLabel(self.frame_resultados, text="Custo Total: R$ 0.00", font=("Trebuchet MS", 12), text_color=COR_TEXTO)
        self.lbl_res_custo.pack()
        self.lbl_res_varejo = ctk.CTkLabel(self.frame_resultados, text="VAREJO: R$ 0.00 | Lucro Líquido: R$ 0.00", font=("Trebuchet MS", 16, "bold"), text_color=COR_TEXTO)
        self.lbl_res_varejo.pack(pady=2)
        self.lbl_res_atacado = ctk.CTkLabel(self.frame_resultados, text="ATACADO: R$ 0.00 | Lucro Líquido: R$ 0.00", font=("Trebuchet MS", 16, "bold"), text_color=COR_TEXTO)
        self.lbl_res_atacado.pack(pady=2)
        
        ctk.CTkButton(self.frame_resultados, text="Dar Baixa no Estoque (Cera e Essência)", command=self.baixa_estoque_vela, fg_color=COR_BOTAO, text_color=COR_TEXTO, font=("Trebuchet MS", 12, "bold")).pack(pady=8)

    def atualizar_texto_peso(self, modo):
        if modo == "Peso Total do Vidro":
            self.lbl_peso.configure(text="Peso Total do Vidro (g):")
        else:
            self.lbl_peso.configure(text="Peso apenas da Cera (g):")

    def calcular_velas(self):
        try:
            peso_input = self.obter_valor(self.campos_dados["peso"])
            carga = self.obter_valor(self.campos_dados["carga"]) / 100
            modo = self.modo_calculo.get()

            if modo == "Peso Total do Vidro":
                self.cera_usada = peso_input / (1 + carga) if carga > 0 else peso_input
                self.essencia_usada = peso_input - self.cera_usada
            else:
                self.cera_usada = peso_input
                self.essencia_usada = self.cera_usada * carga

            c_cera = (self.obter_valor(self.campos_dados["preco_cera"]) / (self.obter_valor(self.campos_dados["qtd_cera"]) or 1)) * self.cera_usada
            c_essencia = (self.obter_valor(self.campos_dados["preco_essencia"]) / (self.obter_valor(self.campos_dados["qtd_essencia"]) or 1)) * self.essencia_usada
            custo_massa = c_cera + c_essencia

            custos_secos = sum([self.obter_valor(self.campos_dados[k]) for k in ["vidro", "pavio", "rotulo", "caixa", "outros"]])
            
            salario = self.obter_valor(self.campos_dados["salario"])
            horas_mes = self.obter_valor(self.campos_dados["horas_mes"]) or 1
            minutos_vela = self.obter_valor(self.campos_dados["tempo_vela"])
            custo_mao_obra = ((salario / horas_mes) / 60) * minutos_vela

            self.custo_total_vela = custo_massa + custos_secos + custo_mao_obra
            
            markup_varejo = self.obter_valor(self.campos_dados["markup"]) or 1
            markup_atacado = self.obter_valor(self.campos_dados["markup_atacado"]) or 1
            taxa_perc = self.obter_valor(self.campos_dados["taxa"]) / 100

            self.preco_venda_varejo = self.custo_total_vela * markup_varejo
            lucro_varejo = self.preco_venda_varejo - self.custo_total_vela - (self.preco_venda_varejo * taxa_perc)

            preco_atacado = self.custo_total_vela * markup_atacado
            lucro_atacado = preco_atacado - self.custo_total_vela - (preco_atacado * taxa_perc)

            self.lbl_res_custo.configure(text=f"Custo Base Total: R$ {self.custo_total_vela:.2f} (Massa: R$ {custo_massa:.2f} | Secos: R$ {custos_secos:.2f} | Trab: R$ {custo_mao_obra:.2f})")
            self.lbl_res_varejo.configure(text=f"VAREJO (x{markup_varejo}): Vender por R$ {self.preco_venda_varejo:.2f} | Lucro Líquido: R$ {lucro_varejo:.2f}")
            self.lbl_res_atacado.configure(text=f"ATACADO (x{markup_atacado}): Vender por R$ {preco_atacado:.2f} | Lucro Líquido: R$ {lucro_atacado:.2f}")
        except Exception as e:
            self.lbl_res_custo.configure(text="Erro de cálculo nos campos numéricos.")

    def baixa_estoque_vela(self):
        if hasattr(self, 'cera_usada'):
            conn = sqlite3.connect('ceravelle_app.db')
            conn.execute("UPDATE estoque SET quantidade = quantidade - ? WHERE item = 'Cera'", (self.cera_usada,))
            conn.execute("UPDATE estoque SET quantidade = quantidade - ? WHERE item = 'Essência'", (self.essencia_usada,))
            conn.commit()
            conn.close()
            self.carregar_estoque()
            self.lbl_res_custo.configure(text=self.lbl_res_custo.cget("text") + " | [Baixa Realizada!]")

    def gerar_ficha_tecnica_pdf(self):
        nome_vela = self.entry_nome_receita.get().strip() or "Receita"
        ficheiro = f"Ficha_Tecnica_{nome_vela.replace(' ', '_')}.pdf"
        
        c = canvas.Canvas(ficheiro, pagesize=A4)
        c.setFont("Helvetica-Bold", 18)
        c.drawString(50, 800, "CERAVELLE - FICHA TÉCNICA DE PRODUÇÃO")
        c.setFont("Helvetica", 12)
        c.drawString(50, 780, f"Produto: {nome_vela} | Data: {datetime.now().strftime('%d/%m/%Y')}")
        
        y = 740
        c.drawString(50, y, f"• Cera Necessária: {getattr(self, 'cera_usada', 0):.2f}g")
        y -= 25
        c.drawString(50, y, f"• Essência Necessária: {getattr(self, 'essencia_usada', 0):.2f}g/ml")
        y -= 25
        c.drawString(50, y, f"• Custo Total de Produção: R$ {getattr(self, 'custo_total_vela', 0):.2f}")
        y -= 25
        c.drawString(50, y, f"• Preço de Venda Sugerido (Varejo): R$ {getattr(self, 'preco_venda_varejo', 0):.2f}")
        y -= 35
        
        c.line(50, y, 550, y)
        y -= 30
        c.drawString(50, y, "Instruções do Ateliê:")
        y -= 20
        c.drawString(50, y, "- Fundir a cera na temperatura recomendada.")
        y -= 20
        c.drawString(50, y, "- Adicionar a essência e homogeneizar bem.")
        y -= 20
        c.drawString(50, y, "- Aguardar o tempo de cura completo antes de acender.")
        
        c.save()
        self.lbl_res_custo.configure(text=self.lbl_res_custo.cget("text") + f" | [PDF Gerado: {ficheiro}]")

    def salvar_receita(self):
        nome = self.entry_nome_receita.get().strip()
        if not nome: return
        dados = json.dumps({k: v.get() for k, v in self.campos_dados.items()})
        conn = sqlite3.connect('ceravelle_app.db')
        conn.execute("INSERT OR REPLACE INTO receitas (nome, dados) VALUES (?, ?)", (nome, dados))
        conn.commit()
        conn.close()
        self.atualizar_lista_receitas()
        self.seletor_receitas.set(nome)
        self.atualizar_dashboard()

    def excluir_receita(self):
        nome = self.entry_nome_receita.get().strip()
        if not nome: return
        conn = sqlite3.connect('ceravelle_app.db')
        conn.execute("DELETE FROM receitas WHERE nome = ?", (nome,))
        conn.commit()
        conn.close()
        self.entry_nome_receita.delete(0, 'end')
        self.atualizar_lista_receitas()
        self.seletor_receitas.set("Selecione...")
        self.atualizar_dashboard()

    def atualizar_lista_receitas(self):
        conn = sqlite3.connect('ceravelle_app.db')
        receitas = [r[0] for r in conn.execute("SELECT nome FROM receitas").fetchall()]
        if receitas: 
            self.seletor_receitas.configure(values=receitas)
        else:
            self.seletor_receitas.configure(values=["Selecione..."])
        conn.close()

    def carregar_receita(self, nome):
        if nome == "Selecione...": return
        conn = sqlite3.connect('ceravelle_app.db')
        resultado = conn.execute("SELECT dados FROM receitas WHERE nome = ?", (nome,)).fetchone()
        if resultado:
            dados = json.loads(resultado[0])
            for k, v in self.campos_dados.items():
                v.delete(0, 'end')
                if k in dados: v.insert(0, dados[k])
            self.entry_nome_receita.delete(0, 'end')
            self.entry_nome_receita.insert(0, nome)
            self.calcular_velas()
        conn.close()

    # ==========================================
    # ABA 2: LÍQUIDOS
    # ==========================================
    def construir_aba_liquidos(self):
        self.scroll_liq = ctk.CTkScrollableFrame(self.tab_liquidos, fg_color=COR_BG)
        self.scroll_liq.pack(fill="both", expand=True, pady=5)
        
        ctk.CTkLabel(self.scroll_liq, text="Formulação de Produtos Líquidos", font=("Trebuchet MS", 14, "bold"), text_color=COR_TEXTO).grid(row=0, column=0, columnspan=4, pady=10)
        
        self.tipo_liquido = ctk.StringVar(value="Home Spray")
        self.seletor_liq = ctk.CTkSegmentedButton(self.scroll_liq, values=["Home Spray", "Body Splash"], variable=self.tipo_liquido, selected_color=COR_DESTAQUE, text_color=COR_TEXTO)
        self.seletor_liq.grid(row=1, column=0, columnspan=4, pady=10)

        self.criar_campo(self.scroll_liq, "Volume Frasco (ml):", "vol_frasco", 2, 0, self.campos_liquidos)
        self.criar_campo(self.scroll_liq, "Concentração Essência (%):", "conc_essencia", 2, 2, self.campos_liquidos)
        self.criar_campo(self.scroll_liq, "Preço Base (R$):", "preco_base", 3, 0, self.campos_liquidos)
        self.criar_campo(self.scroll_liq, "Qtd. Base (ml):", "qtd_base", 3, 2, self.campos_liquidos)
        self.criar_campo(self.scroll_liq, "Preço Essência (R$):", "preco_ess", 4, 0, self.campos_liquidos)
        self.criar_campo(self.scroll_liq, "Qtd. Essência (ml):", "qtd_ess", 4, 2, self.campos_liquidos)
        self.criar_campo(self.scroll_liq, "Frasco + Válvula (R$):", "frasco", 5, 0, self.campos_liquidos)
        self.criar_campo(self.scroll_liq, "Markup Varejo:", "markup_liq", 5, 2, self.campos_liquidos)

        ctk.CTkButton(self.scroll_liq, text="Calcular Produto Líquido", command=self.calcular_liquidos, fg_color=COR_DESTAQUE, text_color="white", font=("Trebuchet MS", 12, "bold")).grid(row=6, column=0, columnspan=4, pady=20)
        self.lbl_res_liq = ctk.CTkLabel(self.scroll_liq, text="", font=("Trebuchet MS", 14, "bold"), text_color=COR_TEXTO)
        self.lbl_res_liq.grid(row=7, column=0, columnspan=4)

    def calcular_liquidos(self):
        try:
            vol = self.obter_valor(self.campos_liquidos["vol_frasco"])
            conc = self.obter_valor(self.campos_liquidos["conc_essencia"]) / 100
            essencia_usada = vol * conc
            base_usada = vol - essencia_usada

            c_base = (self.obter_valor(self.campos_liquidos["preco_base"]) / (self.obter_valor(self.campos_liquidos["qtd_base"]) or 1)) * base_usada
            c_ess = (self.obter_valor(self.campos_liquidos["preco_ess"]) / (self.obter_valor(self.campos_liquidos["qtd_ess"]) or 1)) * essencia_usada
            
            custo_total = c_base + c_ess + self.obter_valor(self.campos_liquidos["frasco"])
            venda = custo_total * (self.obter_valor(self.campos_liquidos["markup_liq"]) or 1)
            
            tipo = self.tipo_liquido.get()
            self.lbl_res_liq.configure(text=f"[{tipo}] Custo: R$ {custo_total:.2f} | Venda Sugerida: R$ {venda:.2f}\n(Base: {base_usada:.2f}ml | Essência: {essencia_usada:.2f}ml)")
        except:
            self.lbl_res_liq.configure(text="Erro de preenchimento numérico.")

    # ==========================================
    # ABA 3: FICHA DE TESTES
    # ==========================================
    def construir_aba_testes(self):
        self.frame_db_testes = ctk.CTkFrame(self.tab_testes, fg_color=COR_SECUNDARIA)
        self.frame_db_testes.pack(pady=5, fill="x")
        
        self.entry_nome_teste = ctk.CTkEntry(self.frame_db_testes, placeholder_text="ID do Teste (Ex: T-01)", width=120, fg_color="white", border_color=COR_DESTAQUE, text_color=COR_TEXTO, font=("Trebuchet MS", 12))
        self.entry_nome_teste.pack(side="left", padx=5, pady=10)
        
        ctk.CTkButton(self.frame_db_testes, text="Salvar", command=self.salvar_ficha_teste, fg_color=COR_DESTAQUE, text_color="white", font=("Trebuchet MS", 11, "bold"), width=75).pack(side="left", padx=2)
        ctk.CTkButton(self.frame_db_testes, text="Limpar", command=self.limpar_campos_teste, fg_color="#7F8C8D", text_color="white", font=("Trebuchet MS", 11, "bold"), width=75).pack(side="left", padx=2)
        ctk.CTkButton(self.frame_db_testes, text="Excluir", command=self.excluir_ficha_teste, fg_color="#C9302C", text_color="white", font=("Trebuchet MS", 11, "bold"), width=75).pack(side="left", padx=2)
        ctk.CTkButton(self.frame_db_testes, text="Gerar PDF + QR", command=self.gerar_pdf_teste, fg_color="#2980B9", text_color="white", font=("Trebuchet MS", 11, "bold"), width=120).pack(side="left", padx=2)
        
        self.seletor_testes = ctk.CTkOptionMenu(self.frame_db_testes, values=["Carregar Ficha..."], command=self.carregar_ficha_teste, fg_color=COR_BOTAO, button_color=COR_DESTAQUE, text_color=COR_TEXTO, width=130)
        self.seletor_testes.pack(side="right", padx=10)
        self.atualizar_lista_testes()

        self.scroll_testes = ctk.CTkScrollableFrame(self.tab_testes, fg_color=COR_BG)
        self.scroll_testes.pack(fill="both", expand=True, pady=5)

        f_cab = ctk.CTkFrame(self.scroll_testes, fg_color="transparent")
        f_cab.pack(fill="x", pady=5)
        
        ctk.CTkLabel(f_cab, text="Data:", text_color=COR_TEXTO).grid(row=0, column=0, padx=5, pady=2, sticky="e")
        f_data_box = ctk.CTkFrame(f_cab, fg_color="transparent")
        f_data_box.grid(row=0, column=1, padx=5, pady=2, sticky="w")
        
        entry_data = ctk.CTkEntry(f_data_box, width=95, height=25, fg_color="white", border_color=COR_DESTAQUE, text_color=COR_TEXTO, font=("Trebuchet MS", 11))
        entry_data.pack(side="left", padx=(0, 2))
        entry_data.insert(0, datetime.now().strftime("%d/%m/%Y"))
        self.campos_teste["t_data"] = entry_data
        
        btn_hoje = ctk.CTkButton(f_data_box, text="📅 Hoje", width=55, height=25, fg_color=COR_DESTAQUE, text_color="white", font=("Trebuchet MS", 9, "bold"),
                                  command=lambda: [entry_data.delete(0, 'end'), entry_data.insert(0, datetime.now().strftime("%d/%m/%Y"))])
        btn_hoje.pack(side="left")

        self.add_campo_teste(f_cab, "Número do Teste:", "t_num", 0, 2)
        self.add_campo_teste(f_cab, "Horário Início:", "t_hora_ini", 0, 4)
        self.add_campo_teste(f_cab, "Horário Fim:", "t_hora_fim", 0, 6)

        ctk.CTkLabel(self.scroll_testes, text="Formulação e Temperaturas", font=("Trebuchet MS", 14, "bold"), text_color=COR_TEXTO).pack(anchor="w", pady=(10, 0))
        f_insumos = ctk.CTkFrame(self.scroll_testes, fg_color=COR_SECUNDARIA)
        f_insumos.pack(fill="x", pady=5, padx=5)
        
        self.add_campo_teste(f_insumos, "Tipo de Cera:", "t_cera_tipo", 0, 0)
        self.add_campo_teste(f_insumos, "Qnt (g):", "t_cera_qnt", 0, 2)
        self.add_campo_teste(f_insumos, "Temp. Derretimento:", "t_cera_temp", 0, 4)
        
        self.add_campo_teste(f_insumos, "Tipo de Pavio:", "t_pavio_tipo", 1, 0)
        self.add_campo_teste(f_insumos, "Qnt (und):", "t_pavio_qnt", 1, 2)
        
        self.add_campo_teste(f_insumos, "Essência:", "t_essencia", 2, 0)
        self.add_campo_teste(f_insumos, "Qnt (%):", "t_essencia_perc", 2, 2)
        self.add_campo_teste(f_insumos, "Temp. Mistura:", "t_essencia_temp", 2, 4)
        self.add_campo_teste(f_insumos, "Qnt (g):", "t_essencia_g", 2, 6)
        
        self.add_campo_teste(f_insumos, "Corante:", "t_corante", 3, 0)
        self.add_campo_teste(f_insumos, "Qnt (g):", "t_corante_qnt", 3, 2)
        
        self.add_campo_teste(f_insumos, "Temp. Envase:", "t_envase_temp", 4, 0)
        self.add_campo_teste(f_insumos, "Peso Total Vela:", "t_peso_total", 4, 2)
        
        self.add_campo_teste(f_insumos, "Cold Throw (0-10):", "t_cold_throw", 5, 0)
        self.add_campo_teste(f_insumos, "Cômodo Teste:", "t_comodo", 5, 2)
        
        self.add_campo_teste(f_insumos, "Diâmetro Recipiente:", "t_diametro", 6, 0)
        self.add_campo_teste(f_insumos, "Modelo Recipiente:", "t_modelo", 6, 2)

        ctk.CTkLabel(self.scroll_testes, text="Informações do Teste (Grelha de Acompanhamento)", font=("Trebuchet MS", 14, "bold"), text_color=COR_TEXTO).pack(anchor="w", pady=(15, 5))
        f_grelha = ctk.CTkFrame(self.scroll_testes, fg_color=COR_SECUNDARIA)
        f_grelha.pack(fill="x", pady=5, padx=5)

        colunas = ["Após 1h", "Após 2h", "Após 3h", "Após 4h"]
        linhas = ["Altura da chama (cm)", "Chama estável?", "Aroma percebido (0-10)", "Fuligem (Sim/Não)", 
                  "Cogumelo", "Túnel", "Prof. piscina fusão", "Piscina até a borda", "Estabilidade Recipiente"]

        for j, col_nome in enumerate(colunas):
            ctk.CTkLabel(f_grelha, text=col_nome, font=("Trebuchet MS", 12, "bold"), text_color=COR_TEXTO).grid(row=0, column=j+1, padx=5, pady=5)

        for i, linha_nome in enumerate(linhas):
            ctk.CTkLabel(f_grelha, text=linha_nome, text_color=COR_TEXTO).grid(row=i+1, column=0, padx=10, pady=2, sticky="e")
            for j in range(4):
                chave_grelha = f"t_grid_{i}_{j}"
                entry = ctk.CTkEntry(f_grelha, width=100, height=25, fg_color="white", border_color=COR_DESTAQUE, text_color=COR_TEXTO)
                entry.grid(row=i+1, column=j+1, padx=5, pady=2)
                self.campos_teste[chave_grelha] = entry

        ctk.CTkLabel(self.scroll_testes, text="Observações Finais e Adicionais", font=("Trebuchet MS", 14, "bold"), text_color=COR_TEXTO).pack(anchor="w", pady=(15, 5))
        
        f_finais = ctk.CTkFrame(self.scroll_testes, fg_color=COR_SECUNDARIA)
        f_finais.pack(fill="x", pady=5, padx=5)
        
        self.add_campo_teste(f_finais, "Tempo Queima Total:", "t_queima_total", 0, 0)
        self.add_campo_teste(f_finais, "Piscina completa?:", "t_piscina_comp", 0, 2)
        self.add_campo_teste(f_finais, "Altura Pavio Início:", "t_pavio_ini", 1, 0)
        self.add_campo_teste(f_finais, "Formou Piscina às:", "t_piscina_hora", 1, 2)

        ctk.CTkLabel(self.scroll_testes, text="Observações Adicionais", font=("Trebuchet MS", 12, "bold"), text_color=COR_TEXTO).pack(anchor="w", pady=(10, 0))
        self.teste_obs_texto = ctk.CTkTextbox(self.scroll_testes, height=80, fg_color="white", border_color=COR_DESTAQUE, text_color=COR_TEXTO)
        self.teste_obs_texto.pack(pady=5, padx=5, fill="x")
        self.campos_teste["t_obs_adicionais"] = self.teste_obs_texto

        self.lbl_status_teste = ctk.CTkLabel(self.scroll_testes, text="", text_color=COR_TEXTO)
        self.lbl_status_teste.pack(pady=5)

    def add_campo_teste(self, parent, texto, chave, row, col):
        ctk.CTkLabel(parent, text=texto, text_color=COR_TEXTO).grid(row=row, column=col, padx=5, pady=2, sticky="e")
        entry = ctk.CTkEntry(parent, width=120, height=25, fg_color="white", border_color=COR_DESTAQUE, text_color=COR_TEXTO, font=("Trebuchet MS", 11))
        entry.grid(row=row, column=col+1, padx=5, pady=2, sticky="w")
        self.campos_teste[chave] = entry

    def salvar_ficha_teste(self):
        nome_teste = self.entry_nome_teste.get().strip()
        if not nome_teste:
            self.lbl_status_teste.configure(text="Digite um ID para o Teste!", text_color="red")
            return
            
        dados_teste = {}
        for k, v in self.campos_teste.items():
            if isinstance(v, ctk.CTkEntry): dados_teste[k] = v.get()
            elif isinstance(v, ctk.CTkTextbox): dados_teste[k] = v.get("0.0", "end-1c")

        conn = sqlite3.connect('ceravelle_app.db')
        conn.execute("INSERT OR REPLACE INTO testes_completos (nome_teste, dados) VALUES (?, ?)", (nome_teste, json.dumps(dados_teste)))
        conn.commit()
        conn.close()
        
        self.atualizar_lista_testes()
        self.seletor_testes.set(nome_teste)
        self.lbl_status_teste.configure(text=f"Ficha '{nome_teste}' guardada/atualizada!", text_color=COR_TEXTO)

    def limpar_campos_teste(self):
        # Limpa o ID do teste
        self.entry_nome_teste.delete(0, 'end')
        
        # Percorre todos os campos e caixas de texto da aba de testes e limpa-os
        for k, v in self.campos_teste.items():
            if k == "t_data":
                # Mantém a data de hoje por conveniência
                v.delete(0, 'end')
                v.insert(0, datetime.now().strftime("%d/%m/%Y"))
            elif isinstance(v, ctk.CTkEntry):
                v.delete(0, 'end')
            elif isinstance(v, ctk.CTkTextbox):
                v.delete("0.0", "end")
                
        self.lbl_status_teste.configure(text="Campos limpos com sucesso!", text_color=COR_TEXTO)

    def excluir_ficha_teste(self):
        nome_teste = self.entry_nome_teste.get().strip()
        if not nome_teste: return
        conn = sqlite3.connect('ceravelle_app.db')
        conn.execute("DELETE FROM testes_completos WHERE nome_teste = ?", (nome_teste,))
        conn.commit()
        conn.close()
        
        # Limpa visualmente o ecrã após excluir
        self.limpar_campos_teste()
        self.atualizar_lista_testes()
        self.seletor_testes.set("Carregar Ficha...")
        self.lbl_status_teste.configure(text=f"Ficha '{nome_teste}' excluída e ecrã limpo!", text_color="#C9302C")

    def atualizar_lista_testes(self):
        conn = sqlite3.connect('ceravelle_app.db')
        testes = [r[0] for r in conn.execute("SELECT nome_teste FROM testes_completos").fetchall()]
        if testes: 
            self.seletor_testes.configure(values=testes)
        else:
            self.seletor_testes.configure(values=["Carregar Ficha..."])
        conn.close()

    def carregar_ficha_teste(self, nome_teste):
        if nome_teste == "Carregar Ficha...": return
        conn = sqlite3.connect('ceravelle_app.db')
        resultado = conn.execute("SELECT dados FROM testes_completos WHERE nome_teste = ?", (nome_teste,)).fetchone()
        if resultado:
            dados = json.loads(resultado[0])
            for k, v in self.campos_teste.items():
                if k in dados:
                    if isinstance(v, ctk.CTkEntry):
                        v.delete(0, 'end'); v.insert(0, dados[k])
                    elif isinstance(v, ctk.CTkTextbox):
                        v.delete("0.0", "end"); v.insert("0.0", dados[k])
            self.entry_nome_teste.delete(0, 'end')
            self.entry_nome_teste.insert(0, nome_teste)
            self.lbl_status_teste.configure(text=f"Ficha '{nome_teste}' carregada!", text_color=COR_TEXTO)
        conn.close()

    def gerar_pdf_teste(self):
        nome_t = self.entry_nome_teste.get().strip() or "Teste"
        ficheiro = f"Ficha_Teste_{nome_t.replace(' ', '_')}.pdf"
        
        self.salvar_ficha_teste()

        ip_computador = get_local_ip()
        url_telemovel = f"http://{ip_computador}:5000/preencher/{nome_t}"
        
        qr = qrcode.QRCode(box_size=3, border=1)
        qr.add_data(url_telemovel)
        qr.make(fit=True)
        qr_img = qr.make_image(fill_color="black", back_color="white")
        qr_path = f"qr_temp_{nome_t}.png"
        qr_img.save(qr_path)

        c = canvas.Canvas(ficheiro, pagesize=A4)
        
        # Cabeçalho da Marca (Afastado à esquerda para dar total liberdade ao QR Code à direita)
        c.setFont("Helvetica-Bold", 15)
        c.setFillColorRGB(0.29, 0.23, 0.19)
        c.drawString(40, 810, "CERAVELLE — FICHA DE TESTE DE QUEIMA")
        
        # QR Code posicionado no canto superior direito real (sem tocar em nada)
        c.drawImage(qr_path, 495, 775, width=55, height=55)

        d_data = self.campos_teste['t_data'].get() or "___/___/______"
        d_num = self.campos_teste['t_num'].get() or "____"
        d_ini = self.campos_teste['t_hora_ini'].get() or "___:___"
        d_fim = self.campos_teste['t_hora_fim'].get() or "___:___"

        c.setFont("Helvetica", 10)
        c.drawString(40, 790, f"ID: {nome_t}  |  Nº: {d_num}")
        c.drawString(180, 790, f"Data: {d_data}")
        c.drawString(310, 790, f"Início: {d_ini}  |  Fim: {d_fim}")
        
        def val(chave, padrao="________"):
            v = self.campos_teste.get(chave)
            if v and hasattr(v, 'get'):
                res = v.get().strip()
                return res if res else padrao
            return padrao

        # --- BLOCO 1: FORMULAÇÃO E INSUMOS (Com espaço seguro abaixo do cabeçalho) ---
        c.setFont("Helvetica-Bold", 11)
        c.drawString(40, 755, "1. Formulação e Insumos")
        
        box1_y = 615
        box1_h = 130
        c.setStrokeColorRGB(0.77, 0.63, 0.35)
        c.rect(40, box1_y, 515, box1_h)
        
        c.setFont("Helvetica", 9)
        # Linha 1
        c.drawString(48, box1_y + 110, f"• Tipo de Cera: {val('t_cera_tipo')}    |    Qnt Cera: {val('t_cera_qnt', '____')} g    |    Temp. Derretimento: {val('t_cera_temp', '____')} °C")
        # Linha 2
        c.drawString(48, box1_y + 90, f"• Tipo de Pavio: {val('t_pavio_tipo')}    |    Qnt Pavio: {val('t_pavio_qnt', '____')} und")
        # Linha 3
        c.drawString(48, box1_y + 70, f"• Essência: {val('t_essencia')}    |    Carga (%): {val('t_essencia_perc', '__')}%    |    Qnt: {val('t_essencia_g', '___')} g    |    Mistura: {val('t_essencia_temp', '___')} °C")
        # Linha 4
        c.drawString(48, box1_y + 50, f"• Corante: {val('t_corante')}    |    Qnt Corante: {val('t_corante_qnt', '____')} g    |    Temp. Envase: {val('t_envase_temp', '____')} °C")
        # Linha 5
        c.drawString(48, box1_y + 30, f"• Peso Total Vela: {val('t_peso_total', '____')} g    |    Cold Throw: {val('t_cold_throw', '__')}    |    Cômodo: {val('t_comodo', '______')}")
        # Linha 6
        c.drawString(48, box1_y + 10, f"• Recipiente (Modelo): {val('t_modelo')}    |    Diâmetro: {val('t_diametro', '____')} cm")

        # --- BLOCO 2: TABELA DE ACOMPANHAMENTO (GRELHA) ---
        c.setFont("Helvetica-Bold", 11)
        c.drawString(40, 590, "2. Grelha de Acompanhamento (Preenchimento no Ateliê)")
        
        y_tabela = 570
        row_h = 17
        c.setFillColorRGB(0.93, 0.90, 0.86)
        c.rect(40, y_tabela - row_h, 515, row_h, fill=1, stroke=1)
        
        c.setFillColorRGB(0.29, 0.23, 0.19)
        c.setFont("Helvetica-Bold", 9)
        c.drawString(45, y_tabela - 12, "Parâmetro / Horário")
        c.drawString(240, y_tabela - 12, "Após 1h")
        c.drawString(310, y_tabela - 12, "Após 2h")
        c.drawString(380, y_tabela - 12, "Após 3h")
        c.drawString(450, y_tabela - 12, "Após 4h")
        
        linhas_grelha = [
            "Altura da chama (cm)", "Chama estável? (S/N)", "Aroma percebido (0-10)", 
            "Fuligem (Sim/Não)", "Cogumelo", "Túnel", "Prof. piscina fusão", 
            "Piscina até a borda", "Estabilidade Recipiente"
        ]
        
        y_linha = y_tabela - row_h
        c.setFont("Helvetica", 9)
        for idx, linha_nome in enumerate(linhas_grelha):
            y_linha -= row_h
            c.rect(40, y_linha, 515, row_h)
            c.drawString(45, y_linha + 5, linha_nome)
            
            c.line(230, y_linha, 230, y_linha + row_h)
            c.line(300, y_linha, 300, y_linha + row_h)
            c.line(370, y_linha, 370, y_linha + row_h)
            c.line(440, y_linha, 440, y_linha + row_h)

        # --- BLOCO 3: OBSERVAÇÕES FINAIS (Expandido até o final da folha) ---
        y_obs_title = y_linha - 20
        c.setFont("Helvetica-Bold", 11)
        c.drawString(40, y_obs_title, "3. Observações Finais e Conclusão")
        
        box3_y = 40
        box3_h = y_obs_title - box3_y - 8
        c.rect(40, box3_y, 515, box3_h)
        c.setFont("Helvetica", 9)
        
        topo_quadro = box3_y + box3_h
        c.drawString(48, topo_quadro - 20, f"• Tempo Queima Total: {val('t_queima_total', '____________')}")
        c.drawString(315, topo_quadro - 20, f"• Piscina Completa?: {val('t_piscina_comp', '________')}")
        
        c.drawString(48, topo_quadro - 40, f"• Altura Pavio Início: {val('t_pavio_ini', '____________')}")
        c.drawString(315, topo_quadro - 40, f"• Formou Piscina às: {val('t_piscina_hora', '________')}")
        
        c.drawString(48, topo_quadro - 60, "• Observações Adicionais / Conclusões:")
        
        obs_texto = ""
        if 't_obs_adicionais' in self.campos_teste:
            ob = self.campos_teste['t_obs_adicionais']
            if hasattr(ob, 'get'):
                obs_texto = ob.get("0.0", "end-1c").strip()
        
        if obs_texto:
            c.drawString(48, topo_quadro - 80, obs_texto)

        c.save()
        
        if os.path.exists(qr_path):
            try: os.remove(qr_path)
            except: pass

        self.lbl_status_teste.configure(text=f"PDF ajustado gerado: {ficheiro}", text_color=COR_TEXTO)

    # ==========================================
    # ABA 4: ESTOQUE
    # ==========================================
    def construir_aba_estoque(self):
        ctk.CTkLabel(self.tab_estoque, text="Gestão de Insumos e Materiais", font=("Trebuchet MS", 16, "bold"), text_color=COR_TEXTO).pack(pady=10)
        
        self.frame_est = ctk.CTkScrollableFrame(self.tab_estoque, width=700, height=350, fg_color=COR_BG)
        self.frame_est.pack(pady=10, padx=20, fill="both", expand=True)
        
        f_add = ctk.CTkFrame(self.tab_estoque, fg_color=COR_SECUNDARIA)
        f_add.pack(pady=10, padx=20, fill="x")
        
        ctk.CTkLabel(f_add, text="Novo Insumo:", text_color=COR_TEXTO).pack(side="left", padx=5)
        self.est_nome_input = ctk.CTkEntry(f_add, placeholder_text="Nome (Ex: Essência Lavanda)", width=220, fg_color="white", border_color=COR_DESTAQUE, text_color=COR_TEXTO, font=("Trebuchet MS", 12))
        self.est_nome_input.pack(side="left", padx=5)
        
        self.est_qtd_input = ctk.CTkEntry(f_add, placeholder_text="Qtd", width=80, fg_color="white", border_color=COR_DESTAQUE, text_color=COR_TEXTO, font=("Trebuchet MS", 12))
        self.est_qtd_input.pack(side="left", padx=5)
        
        self.est_unidade_input = ctk.CTkOptionMenu(f_add, values=['g', 'g/ml', 'ml', 'unit', 'kg', 'L', 'frasco', 'pacote'], fg_color=COR_BOTAO, button_color=COR_DESTAQUE, text_color=COR_TEXTO)
        self.est_unidade_input.pack(side="left", padx=5)
        
        ctk.CTkButton(f_add, text="Cadastrar / Adicionar", command=self.adicionar_item_estoque, fg_color=COR_DESTAQUE, text_color="white", font=("Trebuchet MS", 12, "bold")).pack(side="left", padx=10)
        
        self.carregar_estoque()

    def carregar_estoque(self):
        for widget in self.frame_est.winfo_children(): widget.destroy()
        conn = sqlite3.connect('ceravelle_app.db')
        itens = conn.execute("SELECT item, quantidade, unidade FROM estoque").fetchall()
        
        for i, (item, qtd, unidade) in enumerate(itens):
            cor = "#C9302C" if qtd <= 0 else ("#F0AD4E" if qtd < 500 else "#4CAE4C")
            unidade_str = f" {unidade}" if unidade else ""
            texto_info = f"{item}: {qtd:.2f}{unidade_str}"
            
            lbl = ctk.CTkLabel(self.frame_est, text=texto_info, text_color=cor, font=("Trebuchet MS", 14, "bold"))
            lbl.grid(row=i, column=0, pady=5, padx=10, sticky="w")
            
            btn_remover = ctk.CTkButton(self.frame_est, text="Excluir Item", width=90, fg_color="#C9302C", text_color="white", 
                                        command=lambda it=item: self.excluir_item_estoque(it))
            btn_remover.grid(row=i, column=1, padx=10, pady=5)
        conn.close()

    def adicionar_item_estoque(self):
        nome_item = self.est_nome_input.get().strip()
        qtd_texto = self.est_qtd_input.get().strip().replace(',', '.')
        unidade = self.est_unidade_input.get()

        if not nome_item or not qtd_texto: return

        try:
            qtd = float(qtd_texto)
            conn = sqlite3.connect('ceravelle_app.db')
            existe = conn.execute("SELECT quantidade FROM estoque WHERE item = ? AND unidade = ?", (nome_item, unidade)).fetchone()
            if existe:
                conn.execute("UPDATE estoque SET quantidade = quantidade + ? WHERE item = ? AND unidade = ?", (qtd, nome_item, unidade))
            else:
                conn.execute("INSERT INTO estoque (item, quantidade, unidade) VALUES (?, ?, ?)", (nome_item, qtd, unidade))
            conn.commit()
            conn.close()
            
            self.est_nome_input.delete(0, 'end')
            self.est_qtd_input.delete(0, 'end')
            self.carregar_estoque()
        except ValueError:
            pass

    def excluir_item_estoque(self, item_nome):
        conn = sqlite3.connect('ceravelle_app.db')
        conn.execute("DELETE FROM estoque WHERE item = ?", (item_nome,))
        conn.commit()
        conn.close()
        self.carregar_estoque()

    # ==========================================
    # ABA 4.5: DESPESAS & CUSTOS FIXOS
    # ==========================================
    def construir_aba_despesas(self):
        ctk.CTkLabel(self.tab_despesas, text="Controlo de Despesas Fixas e Operacionais", font=("Trebuchet MS", 16, "bold"), text_color=COR_TEXTO).pack(pady=10)

        f_desp = ctk.CTkFrame(self.tab_despesas, fg_color=COR_SECUNDARIA)
        f_desp.pack(pady=5, padx=20, fill="x")

        self.desp_desc = ctk.CTkEntry(f_desp, placeholder_text="Descrição (Ex: Luz, Internet, Embalagens)", width=260, fg_color="white", border_color=COR_DESTAQUE, text_color=COR_TEXTO, font=("Trebuchet MS", 12))
        self.desp_desc.grid(row=0, column=0, padx=5, pady=5)

        self.desp_valor = ctk.CTkEntry(f_desp, placeholder_text="Valor Mensal (R$)", width=130, fg_color="white", border_color=COR_DESTAQUE, text_color=COR_TEXTO, font=("Trebuchet MS", 12))
        self.desp_valor.grid(row=0, column=1, padx=5, pady=5)

        ctk.CTkButton(f_desp, text="Registar Despesa", command=self.salvar_despesa, fg_color=COR_DESTAQUE, text_color="white", font=("Trebuchet MS", 12, "bold")).grid(row=0, column=2, padx=10, pady=5)

        ctk.CTkLabel(self.tab_despesas, text="Despesas Registadas", font=("Trebuchet MS", 14, "bold"), text_color=COR_TEXTO).pack(pady=5)
        self.frame_desp_lista = ctk.CTkScrollableFrame(self.tab_despesas, width=800, height=300, fg_color=COR_BG)
        self.frame_desp_lista.pack(pady=5, padx=20, fill="both", expand=True)
        self.carregar_despesas()

    def salvar_despesa(self):
        desc = self.desp_desc.get().strip()
        val_txt = self.desp_valor.get().strip().replace(',', '.')
        if not desc or not val_txt: return
        try:
            val = float(val_txt)
            conn = sqlite3.connect('ceravelle_app.db')
            conn.execute("INSERT INTO despesas (descricao, valor) VALUES (?, ?)", (desc, val))
            conn.commit()
            conn.close()
            self.desp_desc.delete(0, 'end')
            self.desp_valor.delete(0, 'end')
            self.carregar_despesas()
        except ValueError:
            pass

    def carregar_despesas(self):
        for w in self.frame_desp_lista.winfo_children(): w.destroy()
        conn = sqlite3.connect('ceravelle_app.db')
        despesas = conn.execute("SELECT id, descricao, valor FROM despesas").fetchall()
        
        total_desp = sum([d[2] for d in despesas])
        lbl_tot = ctk.CTkLabel(self.frame_desp_lista, text=f"Total de Custos Fixos Mensais: R$ {total_desp:.2f}", font=("Trebuchet MS", 13, "bold"), text_color="#C0392B")
        lbl_tot.grid(row=0, column=0, sticky="w", pady=5, padx=5)

        for i, d in enumerate(despesas):
            did, desc, val = d
            txt = f"• {desc}: R$ {val:.2f}"
            lbl = ctk.CTkLabel(self.frame_desp_lista, text=txt, font=("Trebuchet MS", 12), text_color=COR_TEXTO)
            lbl.grid(row=i+1, column=0, sticky="w", pady=3, padx=5)
            
            ctk.CTkButton(self.frame_desp_lista, text="Editar", width=65, fg_color=COR_DESTAQUE, text_color="white", command=lambda did=did, des=desc, v=val: self.editar_despesa(did, des, v)).grid(row=i+1, column=1, padx=5, pady=3)
            ctk.CTkButton(self.frame_desp_lista, text="Excluir", width=65, fg_color="#C9302C", text_color="white", command=lambda did=did: self.excluir_despesa(did)).grid(row=i+1, column=2, padx=5, pady=3)
        conn.close()

    def editar_despesa(self, did, desc, valor):
        self.desp_desc.delete(0, 'end'); self.desp_desc.insert(0, desc)
        self.desp_valor.delete(0, 'end'); self.desp_valor.insert(0, str(valor))
        self.excluir_despesa(did)

    def excluir_despesa(self, did):
        conn = sqlite3.connect('ceravelle_app.db')
        conn.execute("DELETE FROM despesas WHERE id = ?", (did,))
        conn.commit()
        conn.close()
        self.carregar_despesas()

    # ==========================================
    # ABA 4.8: KITS & COMBOS
    # ==========================================
    def construir_aba_kits(self):
        ctk.CTkLabel(self.tab_kits, text="Catálogo de Kits e Combos de Produtos", font=("Trebuchet MS", 16, "bold"), text_color=COR_TEXTO).pack(pady=10)

        f_kit = ctk.CTkFrame(self.tab_kits, fg_color=COR_SECUNDARIA)
        f_kit.pack(pady=5, padx=20, fill="x")

        self.kit_nome = ctk.CTkEntry(f_kit, placeholder_text="Nome do Kit (Ex: Kit Dia das Mães)", width=220, fg_color="white", border_color=COR_DESTAQUE, text_color=COR_TEXTO, font=("Trebuchet MS", 12))
        self.kit_nome.grid(row=0, column=0, padx=5, pady=5)

        self.kit_desc = ctk.CTkEntry(f_kit, placeholder_text="Itens inclusos (Ex: 1 Vela + 1 Home Spray)", width=280, fg_color="white", border_color=COR_DESTAQUE, text_color=COR_TEXTO, font=("Trebuchet MS", 12))
        self.kit_desc.grid(row=0, column=1, padx=5, pady=5)

        self.kit_preco = ctk.CTkEntry(f_kit, placeholder_text="Preço Sugerido (R$)", width=140, fg_color="white", border_color=COR_DESTAQUE, text_color=COR_TEXTO, font=("Trebuchet MS", 12))
        self.kit_preco.grid(row=0, column=2, padx=5, pady=5)

        ctk.CTkButton(f_kit, text="Salvar Kit", command=self.salvar_kit, fg_color=COR_DESTAQUE, text_color="white", font=("Trebuchet MS", 12, "bold")).grid(row=0, column=3, padx=10, pady=5)

        ctk.CTkLabel(self.tab_kits, text="Kits Registados", font=("Trebuchet MS", 14, "bold"), text_color=COR_TEXTO).pack(pady=5)
        self.frame_kit_lista = ctk.CTkScrollableFrame(self.tab_kits, width=900, height=300, fg_color=COR_BG)
        self.frame_kit_lista.pack(pady=5, padx=20, fill="both", expand=True)
        self.carregar_kits()

    def salvar_kit(self):
        nome = self.kit_nome.get().strip()
        desc = self.kit_desc.get().strip()
        preco_txt = self.kit_preco.get().strip().replace(',', '.')
        if not nome or not preco_txt: return
        try:
            preco = float(preco_txt)
            conn = sqlite3.connect('ceravelle_app.db')
            conn.execute("INSERT INTO kits (nome_kit, descricao, preco) VALUES (?, ?, ?)", (nome, desc, preco))
            conn.commit()
            conn.close()
            self.kit_nome.delete(0, 'end')
            self.kit_desc.delete(0, 'end')
            self.kit_preco.delete(0, 'end')
            self.carregar_kits()
        except ValueError:
            pass

    def carregar_kits(self):
        for w in self.frame_kit_lista.winfo_children(): w.destroy()
        conn = sqlite3.connect('ceravelle_app.db')
        kits = conn.execute("SELECT id, nome_kit, descricao, preco FROM kits").fetchall()
        conn.close()

        for i, k in enumerate(kits):
            kid, nome, desc, preco = k
            txt = f"🎁 {nome} — Itens: {desc} | Preço de Venda: R$ {preco:.2f}"
            ctk.CTkLabel(self.frame_kit_lista, text=txt, font=("Trebuchet MS", 13, "bold"), text_color=COR_TEXTO).grid(row=i, column=0, sticky="w", pady=4, padx=5)
            
            ctk.CTkButton(self.frame_kit_lista, text="Excluir", width=65, fg_color="#C9302C", text_color="white", command=lambda kid=kid: self.excluir_kit(kid)).grid(row=i, column=1, padx=10, pady=2)

    def excluir_kit(self, kid):
        conn = sqlite3.connect('ceravelle_app.db')
        conn.execute("DELETE FROM kits WHERE id = ?", (kid,))
        conn.commit()
        conn.close()
        self.carregar_kits()

    # ==========================================
    # ABA 5: PRODUÇÃO & LOTES
    # ==========================================
    def construir_aba_producao(self):
        ctk.CTkLabel(self.tab_producao, text="Registo e Histórico de Produção em Lote", font=("Trebuchet MS", 16, "bold"), text_color=COR_TEXTO).pack(pady=10)
        
        f_reg = ctk.CTkFrame(self.tab_producao, fg_color=COR_SECUNDARIA)
        f_reg.pack(pady=10, padx=20, fill="x")
        
        ctk.CTkLabel(f_reg, text="Nome da Vela/Produto:", text_color=COR_TEXTO).grid(row=0, column=0, padx=10, pady=5, sticky="e")
        self.prod_nome = ctk.CTkEntry(f_reg, placeholder_text="Ex: Vela Lavanda", width=200, fg_color="white", border_color=COR_DESTAQUE, text_color=COR_TEXTO, font=("Trebuchet MS", 12))
        self.prod_nome.grid(row=0, column=1, padx=10, pady=5)

        ctk.CTkLabel(f_reg, text="Quantidade Produzida:", text_color=COR_TEXTO).grid(row=0, column=2, padx=10, pady=5, sticky="e")
        self.prod_qtd = ctk.CTkEntry(f_reg, placeholder_text="Ex: 10", width=100, fg_color="white", border_color=COR_DESTAQUE, text_color=COR_TEXTO, font=("Trebuchet MS", 12))
        self.prod_qtd.grid(row=0, column=3, padx=10, pady=5)

        ctk.CTkButton(f_reg, text=" Registar Lote Produzido ", command=self.registar_producao, fg_color=COR_DESTAQUE, text_color="white", font=("Trebuchet MS", 12, "bold")).grid(row=1, column=0, columnspan=4, pady=15)

        ctk.CTkLabel(self.tab_producao, text="Histórico de Lotes Registados", font=("Trebuchet MS", 14, "bold"), text_color=COR_TEXTO).pack(pady=5)
        self.frame_hist = ctk.CTkScrollableFrame(self.tab_producao, width=800, height=300, fg_color=COR_BG)
        self.frame_hist.pack(pady=5, padx=20, fill="both", expand=True)
        self.carregar_historico_producao()

    def registar_producao(self):
        nome = self.prod_nome.get().strip()
        qtd_txt = self.prod_qtd.get().strip()
        if not nome or not qtd_txt: return
        
        try:
            qtd = int(qtd_txt)
            custo_unit = getattr(self, 'custo_total_vela', 0.0)
            preco_unit = getattr(self, 'preco_venda_varejo', 0.0)
            
            custo_total = custo_unit * qtd
            faturamento = preco_unit * qtd
            lucro = faturamento - custo_total
            data_atual = datetime.now().strftime("%d/%m/%Y %H:%M")

            conn = sqlite3.connect('ceravelle_app.db')
            conn.execute("INSERT INTO producao (data, receita, quantidade, custo_total, faturamento, lucro) VALUES (?, ?, ?, ?, ?, ?)",
                         (data_atual, nome, qtd, custo_total, faturamento, lucro))
            conn.commit()
            conn.close()

            self.prod_nome.delete(0, 'end')
            self.prod_qtd.delete(0, 'end')
            self.carregar_historico_producao()
            self.atualizar_dashboard()
        except ValueError:
            pass

    def carregar_historico_producao(self):
        for w in self.frame_hist.winfo_children(): w.destroy()
        conn = sqlite3.connect('ceravelle_app.db')
        lotes = conn.execute("SELECT id, data, receita, quantidade, custo_total, faturamento, lucro FROM producao ORDER BY id DESC").fetchall()
        for i, l in enumerate(lotes):
            txt = f"Data: {l[1]} | Produto: {l[2]} | Qtd: {l[3]} un | Custo: R$ {l[4]:.2f} | Venda: R$ {l[5]:.2f} | Lucro: R$ {l[6]:.2f}"
            ctk.CTkLabel(self.frame_hist, text=txt, font=("Trebuchet MS", 12), text_color=COR_TEXTO).grid(row=i, column=0, sticky="w", pady=3, padx=5)
            
            ctk.CTkButton(self.frame_hist, text="Editar", width=65, fg_color=COR_DESTAQUE, text_color="white", command=lambda lid=l[0], r=l[2], q=l[3]: self.editar_producao(lid, r, q)).grid(row=i, column=1, padx=5, pady=3)
            ctk.CTkButton(self.frame_hist, text="Excluir", width=65, fg_color="#C9302C", text_color="white", command=lambda lid=l[0]: self.excluir_producao(lid)).grid(row=i, column=2, padx=5, pady=3)
        conn.close()

    def editar_producao(self, lid, receita, qtd):
        self.prod_nome.delete(0, 'end'); self.prod_nome.insert(0, receita)
        self.prod_qtd.delete(0, 'end'); self.prod_qtd.insert(0, str(qtd))
        self.excluir_producao(lid)

    def excluir_producao(self, lid):
        conn = sqlite3.connect('ceravelle_app.db')
        conn.execute("DELETE FROM producao WHERE id = ?", (lid,))
        conn.commit()
        conn.close()
        self.carregar_historico_producao()
        self.atualizar_dashboard()

    # ==========================================
    # ABA 6: ENCOMENDAS & CRM
    # ==========================================
    def construir_aba_pedidos(self):
        ctk.CTkLabel(self.tab_pedidos, text="Gestão de Clientes e Encomendas (com Canal de Venda)", font=("Trebuchet MS", 16, "bold"), text_color=COR_TEXTO).pack(pady=10)
        
        f_crm = ctk.CTkFrame(self.tab_pedidos, fg_color=COR_SECUNDARIA)
        f_crm.pack(pady=10, padx=20, fill="x")
        
        self.crm_cliente = ctk.CTkEntry(f_crm, placeholder_text="Nome do Cliente", width=160, fg_color="white", border_color=COR_DESTAQUE, text_color=COR_TEXTO, font=("Trebuchet MS", 12))
        self.crm_cliente.grid(row=0, column=0, padx=4, pady=5)
        
        self.crm_tel = ctk.CTkEntry(f_crm, placeholder_text="Telefone", width=120, fg_color="white", border_color=COR_DESTAQUE, text_color=COR_TEXTO, font=("Trebuchet MS", 12))
        self.crm_tel.grid(row=0, column=1, padx=4, pady=5)
        
        self.crm_produto = ctk.CTkEntry(f_crm, placeholder_text="Produto", width=150, fg_color="white", border_color=COR_DESTAQUE, text_color=COR_TEXTO, font=("Trebuchet MS", 12))
        self.crm_produto.grid(row=0, column=2, padx=4, pady=5)

        self.crm_qtd = ctk.CTkEntry(f_crm, placeholder_text="Qtd", width=50, fg_color="white", border_color=COR_DESTAQUE, text_color=COR_TEXTO, font=("Trebuchet MS", 12))
        self.crm_qtd.grid(row=0, column=3, padx=4, pady=5)

        self.crm_status = ctk.CTkOptionMenu(f_crm, values=['Em Produção', 'Pronto', 'Entregue', 'Cancelado'], fg_color=COR_BOTAO, button_color=COR_DESTAQUE, text_color=COR_TEXTO, width=130)
        self.crm_status.grid(row=0, column=4, padx=4, pady=5)

        self.crm_canal = ctk.CTkOptionMenu(f_crm, values=['Instagram', 'WhatsApp', 'Loja Física', 'Indicação', 'Feira / Evento', 'Outro'], fg_color=COR_BOTAO, button_color=COR_DESTAQUE, text_color=COR_TEXTO, width=130)
        self.crm_canal.grid(row=0, column=5, padx=4, pady=5)

        ctk.CTkButton(f_crm, text="Registar Encomenda", command=self.salvar_encomenda, fg_color=COR_DESTAQUE, text_color="white", font=("Trebuchet MS", 12, "bold")).grid(row=1, column=0, columnspan=6, pady=10)

        self.frame_crm_lista = ctk.CTkScrollableFrame(self.tab_pedidos, width=850, height=300, fg_color=COR_BG)
        self.frame_crm_lista.pack(pady=5, padx=20, fill="both", expand=True)
        self.carregar_encomendas()

    def salvar_encomenda(self):
        cliente = self.crm_cliente.get().strip()
        tel = self.crm_tel.get().strip()
        produto = self.crm_produto.get().strip()
        qtd_txt = self.crm_qtd.get().strip()
        status = self.crm_status.get()
        canal = self.crm_canal.get()
        data_atual = datetime.now().strftime("%d/%m/%Y")

        if not cliente or not produto or not qtd_txt: return

        try:
            qtd = int(qtd_txt)
            conn = sqlite3.connect('ceravelle_app.db')
            conn.execute("INSERT INTO encomendas (cliente, telefone, produto, quantidade, status, data, canal) VALUES (?, ?, ?, ?, ?, ?, ?)",
                         (cliente, tel, produto, qtd, status, data_atual, canal))
            conn.commit()
            conn.close()

            self.crm_cliente.delete(0, 'end')
            self.crm_tel.delete(0, 'end')
            self.crm_produto.delete(0, 'end')
            self.crm_qtd.delete(0, 'end')
            self.carregar_encomendas()
        except ValueError:
            pass

    def carregar_encomendas(self):
        for w in self.frame_crm_lista.winfo_children(): w.destroy()
        conn = sqlite3.connect('ceravelle_app.db')
        pedidos = conn.execute("SELECT id, cliente, telefone, produto, quantidade, status, data, canal FROM encomendas ORDER BY id DESC").fetchall()
        for i, p in enumerate(pedidos):
            cor = "#C9302C" if p[5] == "Em Produção" else ("#4CAE4C" if p[5] == "Pronto" else "#777777")
            canal_str = f" | 📢 Canal: {p[7]}" if len(p) > 7 and p[7] else ""
            txt = f"Data: {p[6]} | Cliente: {p[1]} ({p[2]}) | Item: {p[3]} (Qtd: {p[4]}) | Status: [{p[5]}]{canal_str}"
            ctk.CTkLabel(self.frame_crm_lista, text=txt, text_color=cor, font=("Trebuchet MS", 12, "bold")).grid(row=i, column=0, sticky="w", pady=4, padx=5)
            
            ctk.CTkButton(self.frame_crm_lista, text="Editar", width=65, fg_color=COR_DESTAQUE, text_color="white", command=lambda pid=p[0], cli=p[1], t=p[2], pr=p[3], q=p[4], st=p[5], cn=p[7] if len(p)>7 else 'Instagram': self.editar_encomenda(pid, cli, t, pr, q, st, cn)).grid(row=i, column=1, padx=5, pady=3)
            ctk.CTkButton(self.frame_crm_lista, text="Excluir", width=65, fg_color="#C9302C", text_color="white", command=lambda pid=p[0]: self.excluir_encomenda(pid)).grid(row=i, column=2, padx=5, pady=3)
        conn.close()

    def editar_encomenda(self, pid, cliente, tel, produto, qtd, status, canal):
        self.crm_cliente.delete(0, 'end'); self.crm_cliente.insert(0, cliente)
        self.crm_tel.delete(0, 'end'); self.crm_tel.insert(0, tel)
        self.crm_produto.delete(0, 'end'); self.crm_produto.insert(0, produto)
        self.crm_qtd.delete(0, 'end'); self.crm_qtd.insert(0, str(qtd))
        self.crm_status.set(status)
        if canal: self.crm_canal.set(canal)
        self.excluir_encomenda(pid)

    def excluir_encomenda(self, pid):
        conn = sqlite3.connect('ceravelle_app.db')
        conn.execute("DELETE FROM encomendas WHERE id = ?", (pid,))
        conn.commit()
        conn.close()
        self.carregar_encomendas()

    # ==========================================
    # ABA 7: CLIENTES & FIDELIDADE
    # ==========================================
    def construir_aba_clientes(self):
        ctk.CTkLabel(self.tab_clientes, text="Registo de Clientes e Histórico de Consumo", font=("Trebuchet MS", 16, "bold"), text_color=COR_TEXTO).pack(pady=10)

        f_cli = ctk.CTkFrame(self.tab_clientes, fg_color=COR_SECUNDARIA)
        f_cli.pack(pady=5, padx=20, fill="x")

        self.cli_nome = ctk.CTkEntry(f_cli, placeholder_text="Nome Completo", width=200, fg_color="white", border_color=COR_DESTAQUE, text_color=COR_TEXTO, font=("Trebuchet MS", 12))
        self.cli_nome.grid(row=0, column=0, padx=5, pady=5)

        self.cli_tel = ctk.CTkEntry(f_cli, placeholder_text="Telefone / WhatsApp", width=140, fg_color="white", border_color=COR_DESTAQUE, text_color=COR_TEXTO, font=("Trebuchet MS", 12))
        self.cli_tel.grid(row=0, column=1, padx=5, pady=5)

        self.cli_email = ctk.CTkEntry(f_cli, placeholder_text="E-mail", width=180, fg_color="white", border_color=COR_DESTAQUE, text_color=COR_TEXTO, font=("Trebuchet MS", 12))
        self.cli_email.grid(row=0, column=2, padx=5, pady=5)

        self.cli_aniv = ctk.CTkEntry(f_cli, placeholder_text="Aniversário (DD/MM)", width=120, fg_color="white", border_color=COR_DESTAQUE, text_color=COR_TEXTO, font=("Trebuchet MS", 12))
        self.cli_aniv.grid(row=0, column=3, padx=5, pady=5)

        ctk.CTkButton(f_cli, text="Cadastrar Cliente", command=self.salvar_cliente, fg_color=COR_DESTAQUE, text_color="white", font=("Trebuchet MS", 12, "bold")).grid(row=0, column=4, padx=10, pady=5)

        f_compra = ctk.CTkFrame(self.tab_clientes, fg_color=COR_SECUNDARIA)
        f_compra.pack(pady=10, padx=20, fill="x")

        ctk.CTkLabel(f_compra, text="Registar Compra:", text_color=COR_TEXTO).grid(row=0, column=0, padx=5, pady=5)
        
        self.cli_seletor = ctk.CTkOptionMenu(f_compra, values=["Selecione o Cliente..."], fg_color=COR_BOTAO, button_color=COR_DESTAQUE, text_color=COR_TEXTO)
        self.cli_seletor.grid(row=0, column=1, padx=5, pady=5)
        self.atualizar_lista_clientes_dropdown()

        self.compra_produto = ctk.CTkEntry(f_compra, placeholder_text="O que comprou (Ex: Vela Lavanda)", width=200, fg_color="white", border_color=COR_DESTAQUE, text_color=COR_TEXTO, font=("Trebuchet MS", 12))
        self.compra_produto.grid(row=0, column=2, padx=5, pady=5)

        self.compra_qtd = ctk.CTkEntry(f_compra, placeholder_text="Qtd", width=60, fg_color="white", border_color=COR_DESTAQUE, text_color=COR_TEXTO, font=("Trebuchet MS", 12))
        self.compra_qtd.grid(row=0, column=3, padx=5, pady=5)

        self.compra_valor = ctk.CTkEntry(f_compra, placeholder_text="Total Gasto (R$)", width=110, fg_color="white", border_color=COR_DESTAQUE, text_color=COR_TEXTO, font=("Trebuchet MS", 12))
        self.compra_valor.grid(row=0, column=4, padx=5, pady=5)

        ctk.CTkButton(f_compra, text="Registar Venda", command=self.registar_compra_cliente, fg_color=COR_DESTAQUE, text_color="white", font=("Trebuchet MS", 12, "bold")).grid(row=0, column=5, padx=10, pady=5)

        ctk.CTkLabel(self.tab_clientes, text="Histórico Consolidado de Clientes e Frequência", font=("Trebuchet MS", 14, "bold"), text_color=COR_TEXTO).pack(pady=5)
        self.frame_cli_lista = ctk.CTkScrollableFrame(self.tab_clientes, width=900, height=300, fg_color=COR_BG)
        self.frame_cli_lista.pack(pady=5, padx=20, fill="both", expand=True)
        self.carregar_clientes_historico()

    def salvar_cliente(self):
        nome = self.cli_nome.get().strip()
        tel = self.cli_tel.get().strip()
        email = self.cli_email.get().strip()
        aniv = self.cli_aniv.get().strip()

        if not nome: return

        conn = sqlite3.connect('ceravelle_app.db')
        conn.execute("INSERT INTO clientes (nome, telefone, email, aniversario) VALUES (?, ?, ?, ?)", (nome, tel, email, aniv))
        conn.commit()
        conn.close()

        self.cli_nome.delete(0, 'end')
        self.cli_tel.delete(0, 'end')
        self.cli_email.delete(0, 'end')
        self.cli_aniv.delete(0, 'end')
        self.atualizar_lista_clientes_dropdown()
        self.carregar_clientes_historico()
        self.atualizar_dashboard()

    def atualizar_lista_clientes_dropdown(self):
        conn = sqlite3.connect('ceravelle_app.db')
        clientes = conn.execute("SELECT id, nome FROM clientes").fetchall()
        conn.close()
        
        if clientes:
            self.clientes_dict = {nome: cid for cid, nome in clientes}
            self.cli_seletor.configure(values=list(self.clientes_dict.keys()))
            self.cli_seletor.set(list(self.clientes_dict.keys())[0])
        else:
            self.clientes_dict = {}
            self.cli_seletor.configure(values=["Nenhum cliente registado"])
            self.cli_seletor.set("Nenhum cliente registado")

    def registar_compra_cliente(self):
        cliente_nome = self.cli_seletor.get()
        if cliente_nome not in self.clientes_dict: return

        cliente_id = self.clientes_dict[cliente_nome]
        produto = self.compra_produto.get().strip()
        qtd_txt = self.compra_qtd.get().strip()
        valor_txt = self.compra_valor.get().strip().replace(',', '.')
        data_hora = datetime.now().strftime("%d/%m/%Y %H:%M")

        if not produto or not qtd_txt or not valor_txt: return

        try:
            qtd = int(qtd_txt)
            valor = float(valor_txt)

            conn = sqlite3.connect('ceravelle_app.db')
            conn.execute("INSERT INTO historico_compras (cliente_id, produto, quantidade, valor_gasto, data_hora) VALUES (?, ?, ?, ?, ?)",
                         (cliente_id, produto, qtd, valor, data_hora))
            conn.commit()
            conn.close()

            self.compra_produto.delete(0, 'end')
            self.compra_qtd.delete(0, 'end')
            self.compra_valor.delete(0, 'end')
            self.carregar_clientes_historico()
        except ValueError:
            pass

    def carregar_clientes_historico(self):
        for w in self.frame_cli_lista.winfo_children(): w.destroy()
        conn = sqlite3.connect('ceravelle_app.db')
        clientes = conn.execute("SELECT id, nome, telefone, email, aniversario FROM clientes").fetchall()

        for i, c in enumerate(clientes):
            cid, nome, tel, email, aniv = c
            res = conn.execute("SELECT COUNT(*), SUM(valor_gasto) FROM historico_compras WHERE cliente_id = ?", (cid,)).fetchone()
            total_compras = res[0] if res[0] else 0
            total_gasto = res[1] if res[1] else 0.0

            header_txt = f"👤 {nome} | Tel: {tel or 'N/I'} | E-mail: {email or 'N/I'} | Aniv: {aniv or 'N/I'} — 🛍️ Compras: {total_compras} | Total Gasto: R$ {total_gasto:.2f}"
            lbl_cli = ctk.CTkLabel(self.frame_cli_lista, text=header_txt, font=("Trebuchet MS", 12, "bold"), text_color=COR_TEXTO)
            lbl_cli.grid(row=i*2, column=0, sticky="w", pady=(10, 2), padx=5)
            
            ctk.CTkButton(self.frame_cli_lista, text="Editar", width=65, fg_color=COR_DESTAQUE, text_color="white", command=lambda cid=cid, n=nome, t=tel, e=email, a=aniv: self.editar_cliente(cid, n, t, e, a)).grid(row=i*2, column=1, padx=5, pady=2)
            ctk.CTkButton(self.frame_cli_lista, text="Excluir", width=65, fg_color="#C9302C", text_color="white", command=lambda cid=cid: self.excluir_cliente(cid)).grid(row=i*2, column=2, padx=5, pady=2)

            compras = conn.execute("SELECT produto, quantidade, valor_gasto, data_hora FROM historico_compras WHERE cliente_id = ?", (cid,)).fetchall()
            if compras:
                detalhes = "   ↳ Histórico: " + " | ".join([f"{comp[3]} - {comp[0]} ({comp[1]} un) - R$ {comp[2]:.2f}" for comp in compras])
            else:
                detalhes = "   ↳ Nenhuma compra registada ainda."
            
            lbl_det = ctk.CTkLabel(self.frame_cli_lista, text=detalhes, font=("Trebuchet MS", 11), text_color="#7A6555")
            lbl_det.grid(row=i*2+1, column=0, sticky="w", pady=(0, 5), padx=15)

        conn.close()

    def editar_cliente(self, cid, nome, tel, email, aniv):
        self.cli_nome.delete(0, 'end'); self.cli_nome.insert(0, nome)
        self.cli_tel.delete(0, 'end'); self.cli_tel.insert(0, tel or "")
        self.cli_email.delete(0, 'end'); self.cli_email.insert(0, email or "")
        self.cli_aniv.delete(0, 'end'); self.cli_aniv.insert(0, aniv or "")
        self.excluir_cliente(cid)

    def excluir_cliente(self, cid):
        conn = sqlite3.connect('ceravelle_app.db')
        conn.execute("DELETE FROM clientes WHERE id = ?", (cid,))
        conn.execute("DELETE FROM historico_compras WHERE cliente_id = ?", (cid,))
        conn.commit()
        conn.close()
        self.atualizar_lista_clientes_dropdown()
        self.carregar_clientes_historico()
        self.atualizar_dashboard()

    # ==========================================
    # ABA 8: ORÇAMENTOS & RELATÓRIOS
    # ==========================================
    def construir_aba_sistema(self):
        f_pdf = ctk.CTkFrame(self.tab_sistema, fg_color=COR_SECUNDARIA)
        f_pdf.pack(pady=10, fill="x", padx=20)
        ctk.CTkLabel(f_pdf, text="Gerador de Orçamentos e Relatórios (PDF & CSV)", font=("Trebuchet MS", 14, "bold"), text_color=COR_TEXTO).pack(pady=5)
        
        self.nome_cliente_pdf = ctk.CTkEntry(f_pdf, placeholder_text="Nome do Cliente/Empresa", width=250, fg_color="white", border_color=COR_DESTAQUE, text_color=COR_TEXTO, font=("Trebuchet MS", 12))
        self.nome_cliente_pdf.pack(pady=5)
        
        self.valor_orcamento_pdf = ctk.CTkEntry(f_pdf, placeholder_text="Valor (R$) - Opcional", width=250, fg_color="white", border_color=COR_DESTAQUE, text_color=COR_TEXTO, font=("Trebuchet MS", 12))
        self.valor_orcamento_pdf.pack(pady=5)

        f_botoes = ctk.CTkFrame(f_pdf, fg_color="transparent")
        f_botoes.pack(pady=5)

        ctk.CTkButton(f_botoes, text="Gerar PDF do Orçamento", command=self.gerar_pdf, fg_color=COR_DESTAQUE, text_color="white", font=("Trebuchet MS", 12, "bold")).pack(side="left", padx=5)
        ctk.CTkButton(f_botoes, text="Exportar Relatório p/ Excel (CSV)", command=self.exportar_csv, fg_color="#27AE60", text_color="white", font=("Trebuchet MS", 12, "bold")).pack(side="left", padx=5)

        self.lbl_pdf_status = ctk.CTkLabel(f_pdf, text="", text_color=COR_TEXTO)
        self.lbl_pdf_status.pack()

        ctk.CTkLabel(self.tab_sistema, text="Orçamentos Emitidos Recentemente", font=("Trebuchet MS", 14, "bold"), text_color=COR_TEXTO).pack(pady=5)
        self.frame_orc_lista = ctk.CTkScrollableFrame(self.tab_sistema, width=900, height=250, fg_color=COR_BG)
        self.frame_orc_lista.pack(pady=5, padx=20, fill="both", expand=True)
        self.carregar_historico_orcamentos()

    def gerar_pdf(self):
        cliente = self.nome_cliente_pdf.get().strip() or "Cliente"
        valor_custom = self.valor_orcamento_pdf.get().strip().replace(',', '.')
        
        if valor_custom:
            try:
                valor_final = float(valor_custom)
            except ValueError:
                valor_final = getattr(self, 'preco_venda_varejo', 0.0)
        else:
            valor_final = getattr(self, 'preco_venda_varejo', 0.0)

        data_atual = datetime.now().strftime("%d/%m/%Y")
        ficheiro = f"Orcamento_{cliente.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        
        c = canvas.Canvas(ficheiro, pagesize=A4)
        c.setFont("Helvetica-Bold", 20)
        c.drawString(50, 800, "CERAVELLE")
        c.setFont("Helvetica", 12)
        c.drawString(50, 780, f"Orçamento Comercial - {data_atual}")
        c.drawString(50, 760, f"Cliente: {cliente}")
        
        c.line(50, 750, 550, 750)
        c.drawString(50, 720, f"Valor Sugerido / Venda: R$ {valor_final:.2f}")
        c.drawString(50, 700, "Documento gerado automaticamente pelo sistema Ceravelle.")
        c.save()

        conn = sqlite3.connect('ceravelle_app.db')
        conn.execute("INSERT INTO orcamentos (cliente, valor, data) VALUES (?, ?, ?)", (cliente, valor_final, data_atual))
        conn.commit()
        conn.close()

        self.lbl_pdf_status.configure(text=f"PDF gerado com sucesso: {ficheiro}", text_color=COR_TEXTO)
        self.carregar_historico_orcamentos()

    def exportar_csv(self):
        try:
            conn = sqlite3.connect('ceravelle_app.db')
            lotes = conn.execute("SELECT data, receita, quantidade, custo_total, faturamento, lucro FROM producao").fetchall()
            conn.close()

            ficheiro_csv = f"Relatorio_Financeiro_Ceravelle_{datetime.now().strftime('%Y%m%d')}.csv"
            with open(ficheiro_csv, "w", encoding="utf-8-sig") as f:
                f.write("Data;Produto;Quantidade;Custo Total (R$);Faturamento (R$);Lucro Líquido (R$)\n")
                for l in lotes:
                    f.write(f"{l[0]};{l[1]};{l[2]};{l[3]:.2f};{l[4]:.2f};{l[5]:.2f}\n".replace('.', ','))

            self.lbl_pdf_status.configure(text=f"Relatório exportado com sucesso: {ficheiro_csv}", text_color="#27AE60")
        except Exception as e:
            self.lbl_pdf_status.configure(text="Erro ao exportar relatório financeiro.", text_color="red")

    def carregar_historico_orcamentos(self):
        for w in self.frame_orc_lista.winfo_children(): w.destroy()
        conn = sqlite3.connect('ceravelle_app.db')
        orcamentos = conn.execute("SELECT id, cliente, valor, data FROM orcamentos ORDER BY id DESC").fetchall()
        conn.close()

        for i, o in enumerate(orcamentos):
            oid, cliente, valor, data = o
            txt = f"Data: {data} | Cliente: {cliente} — Valor: R$ {valor:.2f}"
            ctk.CTkLabel(self.frame_orc_lista, text=txt, font=("Trebuchet MS", 12, "bold"), text_color=COR_TEXTO).grid(row=i, column=0, sticky="w", pady=4, padx=5)
            
            ctk.CTkButton(self.frame_orc_lista, text="Excluir", width=65, fg_color="#C9302C", text_color="white", command=lambda oid=oid: self.excluir_orcamento(oid)).grid(row=i, column=1, padx=10, pady=2)

    def excluir_orcamento(self, oid):
        conn = sqlite3.connect('ceravelle_app.db')
        conn.execute("DELETE FROM orcamentos WHERE id = ?", (oid,))
        conn.commit()
        conn.close()
        self.carregar_historico_orcamentos()

if __name__ == "__main__":
    app = SistemaVelasApp()
    app.mainloop()
