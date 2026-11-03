import streamlit as st
import webbrowser
import time
import re
import pandas as pd
import openpyxl
import os
import json
import urllib.parse
import urllib.request
from datetime import datetime, timedelta

IS_CLOUD = os.path.exists("/mount/src")

if not IS_CLOUD:
    import pyautogui
    import pyperclip
else:
    pyautogui = None
    pyperclip = None

# ==============================================================================
# CONFIGURAÇÃO DA PLANILHA DE AGENDAMENTOS (GOOGLE SHEETS)
# ==============================================================================
URL_PLANILHA_AGENDAMENTOS_CSV = "https://docs.google.com/spreadsheets/d/1jXUw3oZKRPRXV53YCoqBxHUZN95igzpRN_r7yAhhoTk/export?format=csv"
GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyw3jBYk7_FUYZuhzP9fNR7cvp12liyzF7hTwu8N8yqA4ul59brVHtyGg5Mz7WKwbNZ/exec"
ARQUIVO_AGENDAMENTOS_LOCAL = "agendamentos_backup.csv"

# Configuração de página alterada para centrado e com limite elegante
st.set_page_config(page_title="Sistema de Atendimento - FibraWer", layout="centered")

# CSS base global otimizado
st.markdown(
    """
    <style>
    .block-container {
        max-width: 900px;
        padding-top: 1.5rem !important;
        padding-bottom: 1rem !important;
        padding-left: 1rem !important;
        padding-right: 1rem !important;
    }
    h2, h3 {
        font-size: 1.2rem !important;
        margin-top: 0.2rem !important;
        margin-bottom: 0.2rem !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Cabeçalho unificado com margem superior segura para evitar qualquer corte visual
st.markdown(
    """
    <div style="margin-top: 10px; margin-bottom: 15px;">
        <h1 style="font-size: 2.2rem !important; font-weight: 700 !important; color: #262730; margin: 0; padding: 0; line-height: 1.2;">FibraWer ⚡</h1>
        <p style="font-size: 8pt !important; color: #666666; margin: 4px 0 0 0; padding: 0;">Desenvolvido por Anderson Wolfart Möllerke</p>
    </div>
    """,
    unsafe_allow_html=True
)

# --- LISTAS OFICIAIS DE APOIO ---
lista_cidades_assertec = [
    "Alto Feliz", "Bom Principio", "Caxias do Sul", "Feliz", "Harmonia", "Ivoti", 
    "Linha Nova", "Pareci Novo", "Presidente Lucena", "São José do Hortêncio", 
    "São Sebastião do Caí", "São Vendelino", "Tupandi", "Vale Real"
]

lista_servicos_assertec_global = [
    "Escolher...",
    "Atendimento Presencial - REPARO",
    "Atendimento Presencial - REPARO - ROMPIMENTO/LANÇAMENTO até 500 mt",
    "Atendimento presencial - RNC",
    "Visita improdutiva",
    "Desativações de serviços completa",
    "Desativações de serviços parcial / troca de endereço / migração de tecnologia",
    "Diária de atendimentos (sob demanda - dupla)",
    "Km rodado acima de 50km",
    "Ativação de Serviço - Acesso à Internet (0 ~ 300 mt)",
    "Ativação de Serviço - Acesso à Internet (301 ~ 500 mt)",
    "Ativação de Serviço - Agregado (câmeras)",
    "Ativação de Serviço - Agregado (câmeras wifi)",
    "Ativação de Serviço - Avulso (câmeras wifi)",
    "Ativação de Serviço - Agregado (wifi/tv/fone/ponto extra)",
    "Ativação de Serviço - Agregado Avulso (wifi/tv/fone/ponto extra)",
    "Batimento de caixa 1x16 p",
    "Batimento de caixa 1x8p",
    "Metragem excedente (acima de 500 mt)",
    "Troca de ponto interno (ponto principal/wifi/tv/fone/câmeras/ponto extra)",
    "Lançamento de fibra avulsa até f.o 6v",
    "Vistoria Técnica/ Viabilidade técnica",
    "Fusão em caixa",
    "Emenda de cabos rompidos na rede em infraestrutura própria",
    "Instalação de nova CTO",
    "Lançamento de cabo AS",
    "PLANTAO - duplo",
    "PLANTAO - simples",
    "PLANTAO - Atendimento Presencial - REPARO",
    "PLANTAO - Atendimento Presencial - REPARO - ROMPIMENTO/LANÇAMENTO até 500 mt",
    "CORPORATIVO - ATENDIMENTO PRESENCIAL - REPARO",
    "CORPORATIVO - ATENDIMENTO PRESENCIAL - REPARO - ROMPIMENTO/LANÇAMENTO ATÉ 500MT",
    "CORPORATIVO - ATENDIMENTO PRESENCIAL - RNC",
    "CORPORATIVO - ATIVAÇÃO DE SERVIÇO - ACESSO À INTERNET (0 ~ 300 MT)",
    "CORPORATIVO - ATIVAÇÃO DE SERVIÇO - ACESSO À INTERNET (301 ~ 500MT)",
    "CORPORATIVO - ATIVAÇÃO DE SERVIÇO - ACESSO À INTERNET (ACIMA DE 500MT)",
    "CORPORATIVO - ATIVAÇÃO DE SERVIÇO - ACESSO À INTERNET (RÁDIO)",
    "CORPORATIVO - ATIVAÇÃO DE SERVIÇO AGREGADO (CÂMERAS WIFI)",
    "CORPORATIVO - ATIVAÇÃO DE SERVIÇO AGREGADO (CÂMERAS)",
    "CORPORATIVO - ATIVAÇÃO DE SERVIÇO AGREGADO (WIFI/TV/FONE/PONTO EXTRA)",
    "CORPORATIVO - ATIVAÇÃO DE SERVIÇO AGREGADO AVULSO (CÂMERAS WIFI)",
    "CORPORATIVO - ATIVAÇÃO DE SERVIÇO AGREGADO AVULSO (WIFI/TV/FONE/PONTO EXTRA)",
    "CORPORATIVO - BATIMENTO DE CAIXA 1X16 PORTAS",
    "CORPORATIVO - BATIMENTO DE CAIXA 1X8 PORTAS",
    "CORPORATIVO - DESATIVAÇÃO DE SERVIÇO - COMPLETA",
    "CORPORATIVO - DESATIVAÇÃO DE SERVIÇO - PARCIAL",
    "CORPORATIVO - DIÁRIAS (SOB DEMANDA - DUPLA)",
    "CORPORATIVO - KM RODADO (ACIMA DE 50KM)",
    "CORPORATIVO - LANÇAMENTO DE FIBRA AVULSA ATÉ - FIBRA ATÉ 12V",
    "CORPORATIVO - METRAGEM EXCEDENTE (ACIMA DE 500MT)",
    "CORPORATIVO - TROCA DE PONTO INTERNO (PONTO PRINCIPAL/WIFI/TV/FONE/CÂMERAS/PONTO EXTRA)",
    "CORPORATIVO - VISITA IMPRODUTIVA",
    "CORPORATIVO - VISTORIA TÉCNICA / VIABILIDADE TÉCNICA"
]

lista_servicos_amigo_global = [
    "Escolher...", 
    "ATENDIMENTO PRESENCIAL - RNC", 
    "ATIVAÇÃO DE SERVIÇO - ACESSO À INTERNET (0 ~ 300 MT)", 
    "ATIVAÇÃO DE SERVIÇO - ACESSO À INTERNET (301 ~ 500MT)", 
    "ATIVAÇÃO DE SERVIÇO - ACESSO À INTERNET (ACIMA DE 500MT)", 
    "ATIVAÇÃO DE SERVIÇO - ACESSO À INTERNET (RÁDIO)", 
    "ATIVAÇÃO DE SERVIÇO AGREGADO (CÂMERAS WIFI)", 
    "ATIVAÇÃO DE SERVIÇO AGREGADO (Câmeras)", 
    "ATIVAÇÃO DE SERVIÇO AGREGADO (WIFI/TV/FONE/PONTO EXTRA)", 
    "ATIVAÇÃO DE SERVIÇO AGREGADO AVULSO (WIFI/TV/FONE/PONTO EXTRA)", 
    "ATIVAÇÃO DE SERVIÇO AGREGADO AVULSO (CÂMERAS WIFI)", 
    "BATIMENTO DE CAIXA 1X16 PORTAS", 
    "BATIMENTO DE CAIXA 1X8 PORTAS", 
    "DESATIVAÇÃO DE SERVIÇO - PARCIAL", 
    "DIÁRIAS (SOB DEMANDA - DUPLA)", 
    "LANÇAMENTO DE FIBRA AVULSA ATÉ - FIBRA ATÉ 12V", 
    "TROCA DE PONTO INTERNO (PONTO PRINCIPAL/WIFI/TV/FONE/CÂMERAS/PONTO EXTRA)", 
    "VISTORIA TÉCNICA / VIABILIDADE TÉCNICA"
]

# Função global para renderizar Selectbox com borda reativa (Vermelho se "Escolher...", Verde se preenchido)
def selectbox_com_borda_dinamica(label, opcoes, chave):
    valor_atual = st.session_state.get(chave, "Escolher...")
    esta_preenchido = valor_atual != "Escolher..."
    
    cor_borda = "#2e7d32" if esta_preenchido else "#ff4b4b"
    bg_cor = "rgba(46, 125, 50, 0.03)" if esta_preenchido else "rgba(255, 75, 75, 0.03)"

    st.markdown(f"""
        <div style="
            border: 2px solid {cor_borda};
            background-color: {bg_cor};
            border-radius: 6px;
            padding: 4px 8px 0px 8px;
            margin-bottom: 8px;
        ">
        <span style="font-size: 0.85rem; font-weight: 700; color: {cor_borda};">{label}</span>
        </div>
    """, unsafe_allow_html=True)
    
    return st.selectbox(label, opcoes, key=chave, label_visibility="collapsed")

# --- FUNÇÕES DE LEITURA E GRAVAÇÃO DE AGENDAMENTOS (5 COLUNAS: NOME, CLIENTE, PROTOCOLO, CIDADE, OBS) ---
def carregar_agendamentos():
    colunas_padrao = ["NOME", "CLIENTE", "PROTOCOLO", "CIDADE", "OBS"]
    try:
        url_com_cache_buster = f"{URL_PLANILHA_AGENDAMENTOS_CSV}&t={int(time.time())}"
        df = pd.read_csv(url_com_cache_buster)
        df.columns = [str(c).strip().upper() for c in df.columns]
        if not df.empty:
            return df
    except Exception:
        pass
    
    if os.path.exists(ARQUIVO_AGENDAMENTOS_LOCAL):
        try:
            df_local = pd.read_csv(ARQUIVO_AGENDAMENTOS_LOCAL)
            df_local.columns = [str(c).strip().upper() for c in df_local.columns]
            return df_local
        except Exception:
            pass
    return pd.DataFrame(columns=colunas_padrao)

def salvar_agendamento(dados):
    # Salva no Backup Local CSV com as 5 colunas exatas
    df_atual = carregar_agendamentos()
    novo_df = pd.DataFrame([{
        "NOME": dados["nome"],
        "CLIENTE": dados["cliente"],
        "PROTOCOLO": dados["protocolo"],
        "CIDADE": dados["cidade"],
        "OBS": dados["obs"]
    }])
    df_final = pd.concat([df_atual, novo_df], ignore_index=True)
    df_final.to_csv(ARQUIVO_AGENDAMENTOS_LOCAL, index=False)

    # Envia para o Google Sheets via Apps Script
    if GOOGLE_SCRIPT_URL.strip():
        try:
            req = urllib.request.Request(
                GOOGLE_SCRIPT_URL.strip(),
                data=json.dumps(dados).encode("utf-8"),
                headers={"Content-Type": "application/json"},
                method="POST"
            )
            with urllib.request.urlopen(req, timeout=10) as resp:
                return True, "Agendamento salvo na Planilha do Google Sheets!"
        except Exception as e:
            return True, f"Salvo localmente (Aviso Google Sheets: {e})"
    return True, "Salvo localmente!"

# --- CRIAÇÃO DAS ABAS PRINCIPAIS ---
if IS_CLOUD:
    aba_gerador, aba_agendamentos = st.tabs([
        "📱 Gerador de Mensagens (Telemóvel)", 
        "📅 Agendamentos"
    ])
    aba_gestao = None
else:
    aba_gestao, aba_gerador, aba_agendamentos = st.tabs([
        "💻 Gestão & Computador (Assertec & Excel)", 
        "📱 Gerador de Mensagens (Telemóvel)",
        "📅 Agendamentos"
    ])

# ==========================================
# ABA 1: GESTÃO & COMPUTADOR
# ==========================================
if aba_gestao is not None:
    with aba_gestao:
        def limpar_campo():
            st.session_state["msg_input"] = ""
            st.session_state["mensagem_processada"] = False
            st.session_state["tec_msg"] = ""
            st.session_state["tec"] = ""
            st.session_state["cli"] = ""
            st.session_state["proto"] = ""
            st.session_state["cidade"] = ""
            st.session_state["servico"] = ""
            st.session_state["qtd"] = ""
            st.session_state["obs"] = ""

        def extrair(padrao, texto):
            match = re.search(padrao, texto, re.IGNORECASE)
            if match:
                for val in match.groups():
                    if val is not None and val.strip():
                        return val.strip()
            return ""

        def processar_texto_atual(texto):
            linhas = [l.strip() for l in texto.strip().split("\n") if l.strip()]
            primeiro_texto = linhas[0].lower() if len(linhas) > 0 else ""
            if "amigo" in primeiro_texto or "tecpar" in primeiro_texto:
                st.session_state["formulario_escolhido"] = "Brasil TecPar / Amigo (Apenas Planilha Excel)"
            else:
                st.session_state["formulario_escolhido"] = "Assertec / Fibrawer"

            if len(linhas) > 1:
                st.session_state["tec_msg"] = linhas[1].upper()
            else:
                st.session_state["tec_msg"] = ""

            st.session_state["tec"] = extrair(r"Técnico:\s*(.*)", texto)
            st.session_state["cli"] = extrair(r"Cliente:\s*(.*)", texto)
            st.session_state["proto"] = extrair(r"Protocolo:\s*(.*)", texto)
            st.session_state["cidade"] = extrair(r"Cidade:\s*(.*)", texto)
            
            raw_servico = extrair(r"Serviço:\s*(.*)", texto)
            raw_servico = re.sub(r'ATATIVAÇÃO', 'ATIVAÇÃO', raw_servico, flags=re.IGNORECASE)
            st.session_state["servico"] = raw_servico
            
            qtd_extraida = extrair(r"Metragem[:]?\s*(.*)|Cabo[:]?\s*(.*)|Qtd[:]?\s*(.*)|Quantidade[:]?\s*(.*)", texto)
            if not qtd_extraida:
                qtd_extraida = extrair(r"LPU[:]?\s*(.*)", texto)
            st.session_state["qtd"] = qtd_extraida if qtd_extraida else "0"
            
            st.session_state["obs"] = extrair(r"Obs:\s*(.*)|Observação:\s*(.*)", texto)
            st.session_state["mensagem_processada"] = True

        if "mensagem_processada" not in st.session_state:
            st.session_state["mensagem_processada"] = False

        msg_texto = st.text_area(
            "Cole a mensagem do WhatsApp aqui (Ex: ASSERTEC \n MARCIO ...)...", 
            height=280, 
            key="msg_input"
        )

        if msg_texto.strip():
            processar_texto_atual(msg_texto)

        if "formulario_escolhido" not in st.session_state:
            st.session_state["formulario_escolhido"] = "Assertec / Fibrawer"

        col_proc, col_limpar, col_espaco = st.columns([1.2, 1.2, 3])

        with col_proc:
            processar = st.button("Processar Mensagem")

        with col_limpar:
            st.button("Limpar Mensagem", on_click=limpar_campo)

        if processar:
            if not msg_texto.strip():
                st.error("Cole uma mensagem válida.")
            else:
                processar_texto_atual(msg_texto)
                st.success(f"Mensagem processada! Fluxo: **{st.session_state['formulario_escolhido']}** | Técnico Identificado: **{st.session_state['tec_msg']}**")

        def traduzir_para_excel(servico_wpp, metragem_str, cidade_atendimento):
            s = servico_wpp.upper()
            cidade_limpa = str(cidade_atendimento).strip().lower()
            
            cidades_principais = [
                "ivoti", 
                "são josé do hortêncio", 
                "sao jose do hortencio", 
                "linha nova", 
                "são sebastião do caí", 
                "sao sebastiao do cai", 
                "presidente lucena"
            ]
            
            cidade_fora_da_base = cidade_limpa and cidade_limpa not in cidades_principais
            eh_plantao_fb = "PLANTAO" in s and ("REPARO" in s)
            sufixo_fb = " FB" if (cidade_fora_da_base or eh_plantao_fb) else ""

            try:
                m = int(re.sub(r'\D', '', str(metragem_str)))
            except:
                m = 0
            
            if "BATIDA" in s or "BATIMENTO" in s:
                if "1/16" in s or "1X16" in s:
                    return f"BATIDA DE CTO 1/16{sufixo_fb}"
                return f"BATIDA DE CTO 1/8{sufixo_fb}"
            elif "ROMPIMENTO" in s:
                return f"ROMPIMENTO{sufixo_fb}"
            elif "REPARO" in s or "MANUTENÇÃO" in s or "MANUTENCAO" in s or "RNC" in s:
                return f"MANUTENÇAO{sufixo_fb}"
            elif "TROCA DE ENDEREÇO" in s:
                if m > 300:
                    return f"TROCA DE ENDEREÇO MAIS DE 300M{sufixo_fb}"
                return f"TROCA DE ENDEREÇO{sufixo_fb}"
            elif "IMPRODUTIVA" in s or "VISITA IMPRODUTIVA" in s:
                return f"IMPRODUTIVA{sufixo_fb}"
            elif "ATIVAÇÃO" in s or "INSTALAÇÃO" in s or "INTERNET" in s:
                if m > 300:
                    return f"INSTALAÇAO MAIS DE 300M{sufixo_fb}"
                return f"INSTALAÇAO{sufixo_fb}"
            
            return f"INSTALAÇAO{sufixo_fb}"

        def obter_caminho_excel_tecnico(tec_nome):
            nome_limpo = str(tec_nome).strip().upper()
            if not nome_limpo:
                nome_limpo = "MARCIO"
            return rf"C:\Users\wolfa\OneDrive\Documentos\SERVIÇOS AUTOMATIZADOS {nome_limpo}.xlsx"

        def adicionar_ao_excel(caminho_arquivo, dados):
            if not os.path.exists(caminho_arquivo):
                return False, f"Ficheiro não encontrado: {caminho_arquivo}"
            try:
                wb = openpyxl.load_workbook(caminho_arquivo)
                if "LANÇAMENTOS" in wb.sheetnames:
                    ws = wb["LANÇAMENTOS"]
                else:
                    ws = wb.active 

                proxima_linha = ws.max_row + 1
                for row in range(2, ws.max_row + 2):
                    if ws.cell(row=row, column=2).value is None:
                        proxima_linha = row
                        break

                data_atual = datetime.now().strftime("%d/%m/%Y")

                ws.cell(row=proxima_linha, column=1, value=data_atual)
                ws.cell(row=proxima_linha, column=2, value=dados["cliente"])
                ws.cell(row=proxima_linha, column=3, value=dados["protocolo"])
                ws.cell(row=proxima_linha, column=4, value=dados["servico"])
                ws.cell(row=proxima_linha, column=5, value=dados["cidade"])
                ws.cell(row=proxima_linha, column=6, value=dados["metragem"])
                ws.cell(row=proxima_linha, column=8, value=dados["equipamento"])

                wb.save(caminho_arquivo)
                return True, proxima_linha
            except Exception as e:
                return False, str(e)

        def conferir_duplicidade_servico(caminho_arquivo, protocolo, servico_desejado):
            if not os.path.exists(caminho_arquivo):
                return False, 0
            try:
                wb = openpyxl.load_workbook(caminho_arquivo, data_only=True)
                if "LANÇAMENTOS" in wb.sheetnames:
                    ws = wb["LANÇAMENTOS"]
                else:
                    ws = wb.active
                    
                proto_limpo = str(protocolo).strip().lower()
                serv_limpo = str(servico_desejado).strip().lower()
                
                for row in range(2, ws.max_row + 1):
                    col_proto = str(ws.cell(row=row, column=3).value or "").strip().lower()
                    col_serv = str(ws.cell(row=row, column=4).value or "").strip().lower()
                    
                    if col_proto == proto_limpo and col_serv == serv_limpo:
                        return True, row
                return False, 0
            except Exception as e:
                return False, 0

        def conferir_na_planilha_local(caminho_arquivo, termo):
            if not os.path.exists(caminho_arquivo):
                return False, "Ficheiro Excel não encontrado."
            try:
                wb = openpyxl.load_workbook(caminho_arquivo, data_only=True)
                if "LANÇAMENTOS" in wb.sheetnames:
                    ws = wb["LANÇAMENTOS"]
                else:
                    ws = wb.active
                    
                termo_limpo = termo.strip().lower()
                for row in range(2, ws.max_row + 1):
                    cliente = str(ws.cell(row=row, column=2).value or "").lower()
                    protocolo = str(ws.cell(row=row, column=3).value or "").lower()
                    if termo_limpo in cliente or termo_limpo in protocolo:
                        return True, row
                return False, 0
            except Exception as e:
                return False, str(e)

        if st.session_state["mensagem_processada"]:
            st.markdown("---")
            formulario_escolhido = st.session_state["formulario_escolhido"]
            
            if formulario_escolhido == "Assertec / Fibrawer":
                btn_executar = st.button("🚀 Executar Assertec", key="btn_exec_ass_topo")
                if btn_executar:
                    tecnico_val = st.session_state.get("tec", "")
                    cliente_val = st.session_state.get("cli", "")
                    protocolo_val = st.session_state.get("proto", "")
                    cidade_importada = st.session_state.get("cidade", "")
                    servico_importado = st.session_state.get("servico", "")
                    quantidade_val = st.session_state.get("qtd", "") if st.session_state.get("qtd", "").strip() else "1"
                    observacao_val = st.session_state.get("obs", "")
                    
                    cidade_idx = 0
                    for idx, c in enumerate(lista_cidades_assertec):
                        if c.lower() in cidade_importada.lower() and cidade_importada.strip():
                            cidade_idx = idx
                            break
                    
                    servico_idx = 0
                    serv_imp_clean = servico_importado.strip().lower()
                    if serv_imp_clean:
                        for idx, s in enumerate(lista_servicos_assertec_global):
                            if s.strip().lower() == serv_imp_clean:
                                servico_idx = idx
                                break
                        if servico_idx == 0:
                            for idx, s in enumerate(lista_servicos_assertec_global):
                                if serv_imp_clean in s.strip().lower() or s.strip().lower() in serv_imp_clean:
                                    servico_idx = idx
                                    break

                    if not tecnico_val.strip() or not cliente_val.strip() or not protocolo_val.strip():
                        st.error("Preencha o técnico, o cliente e o protocolo.")
                    else:
                        st.warning("⚡ A executar automação Assertec... Não mexa no rato.")
                        webbrowser.open("https://forms.gle/2MENn9xuKjj7pKVm6")
                        time.sleep(2.0)
                        
                        for _ in range(3):
                            pyautogui.press('tab')
                            time.sleep(0.04)
                        pyautogui.press('tab')
                        pyautogui.press('tab')
                        pyautogui.press('tab')
                        pyautogui.press('enter')
                        time.sleep(0.2)
                        pyautogui.press('enter')
                        time.sleep(0.08)
                        
                        pyautogui.press('tab')
                        time.sleep(0.04)
                        pyautogui.press('enter')
                        time.sleep(0.25)
                        pyautogui.write(tecnico_val)
                        time.sleep(0.1)
                        pyautogui.press('enter')
                        time.sleep(0.1)
                        
                        pyautogui.press('tab')
                        pyautogui.press('tab')
                        pyautogui.press('space')
                        time.sleep(0.1)
                        
                        for _ in range(3):
                            pyautogui.press('tab')
                            time.sleep(0.04)
                        
                        pyperclip.copy(str(cliente_val))
                        pyautogui.hotkey('ctrl', 'v')
                        time.sleep(0.08)
                        
                        pyautogui.press('tab')
                        pyperclip.copy(str(protocolo_val))
                        pyautogui.hotkey('ctrl', 'v')
                        time.sleep(0.08)
                        
                        pyautogui.press('tab')
                        for _ in range(cidade_idx):
                            pyautogui.press('tab')
                            time.sleep(0.04)
                        
                        pyautogui.press('space')
                        time.sleep(0.1)
                        
                        passos_restantes = len(lista_cidades_assertec) - cidade_idx
                        for _ in range(passos_restantes):
                            pyautogui.press('tab')
                            time.sleep(0.04)
                        
                        pyautogui.press('enter')
                        time.sleep(0.3)
                        
                        pyautogui.press('home')
                        time.sleep(0.1)
                        for _ in range(servico_idx):
                            pyautogui.press('down')
                            time.sleep(0.03)
                        
                        pyautogui.press('enter')
                        time.sleep(0.2)
                        
                        pyautogui.press('tab')
                        time.sleep(0.1)
                        
                        pyautogui.press('tab')
                        time.sleep(0.1)
                        
                        pyperclip.copy(str(quantidade_val))
                        pyautogui.hotkey('ctrl', 'v')
                        time.sleep(0.1)
                        
                        pyautogui.press('tab')
                        time.sleep(0.1)
                        
                        if observacao_val.strip():
                            pyperclip.copy(str(observacao_val))
                            pyautogui.hotkey('ctrl', 'v')
                            time.sleep(0.1)
                        
                        st.success("Automação Assertec concluída com sucesso! 🚀")

            st.markdown("")
            col_lancar, col_conferir_plan, col_vazio_botoes = st.columns([1.5, 1.5, 2])
            
            tec_msg_atual = st.session_state.get("tec_msg", "MARCIO")
            caminho_excel = obter_caminho_excel_tecnico(tec_msg_atual)
            
            with col_lancar:
                if st.button("🚀 Lançar na Planilha Excel"):
                    cli_v = st.session_state.get("cli", "")
                    proto_v = st.session_state.get("proto", "")
                    serv_v = st.session_state.get("servico", "")
                    cid_v = st.session_state.get("cidade", "")
                    qtd_v = st.session_state.get("qtd", "0")
                    obs_v = st.session_state.get("obs", "")
                    
                    servico_convertido = traduzir_para_excel(serv_v, qtd_v, cid_v)
                    ja_existe, linha_existente = conferir_duplicidade_servico(caminho_excel, proto_v, servico_convertido)
                    
                    if ja_existe:
                        st.warning(f"⚠️ Atenção: O serviço **'{servico_convertido}'** para o protocolo **{proto_v}** já foi lançado na linha **{linha_existente}** do ficheiro de **{tec_msg_atual}**!")
                    else:
                        dados_registo = {
                            "cliente": cli_v,
                            "protocolo": proto_v,
                            "servico": servico_convertido,
                            "cidade": cid_v,
                            "metragem": re.sub(r'\D', '', str(qtd_v)) or "0",
                            "equipamento": obs_v
                        }
                        sucesso, res = adicionar_ao_excel(caminho_excel, dados_registo)
                        if sucesso:
                            st.success(f"✅ Atendimento de **{cli_v}** ({servico_convertido}) adicionado na linha **{res}** da planilha de **{tec_msg_atual}** (aba LANÇAMENTOS)!")
                        else:
                            st.error(f"Erro ao gravar no ficheiro de {tec_msg_atual}: {res}")
                        
            with col_conferir_plan:
                if st.button("🔍 Conferir se Já Existe"):
                    proto_v = st.session_state.get("proto", "")
                    cli_v = st.session_state.get("cli", "")
                    termo_busca = proto_v if proto_v.strip() else cli_v
                    
                    if not termo_busca.strip():
                        st.warning("Nenhum protocolo ou cliente detetado para conferir.")
                    else:
                        encontrado, linha_ou_erro = conferir_na_planilha_local(caminho_excel, termo_busca)
                        if encontrado:
                            st.success(f"✅ Encontrado na planilha de **{tec_msg_atual}** na linha **{linha_ou_erro}**!")
                        else:
                            st.info(f"ℹ️ O termo **'{termo_busca}'** ainda não foi encontrado na planilha de **{tec_msg_atual}**.")

        st.markdown("---")
        st.subheader("🔍 Consulta Global de Cliente / Protocolo (Assertec)")
        termo_pesquisa_global = st.text_input("Consultar por Nome do Cliente ou Protocolo:", value=st.session_state.get("cli", ""), key="pesquisa_global_topo")

        if st.button("🔍 Conferir Assertec", key="btn_conf_ass_topo_key"):
            if not termo_pesquisa_global.strip():
                st.warning("Insira o nome de um cliente ou protocolo na caixa de pesquisa.")
            else:
                with st.spinner("A consultar a planilha da Assertec..."):
                    try:
                        url_planilha = "https://docs.google.com/spreadsheets/d/1MMQ8PAjr2GQfWxPnfYNfHSiiu9AUdE0pvfJikSGQgqU/export?format=csv"
                        df = pd.read_csv(url_planilha)
                        encontrado = False
                        termo_busca = termo_pesquisa_global.strip().lower()
                        for col in df.columns:
                            matches = df[df[col].astype(str).str.lower().str.contains(termo_busca, na=False)]
                            if not matches.empty:
                                encontrado = True
                                break
                        if encontrado:
                            st.success(f"✅ O registo com **'{termo_pesquisa_global}'** foi encontrado na planilha da Assertec!")
                            st.dataframe(matches)
                        else:
                            st.info(f"ℹ️ O termo **'{termo_pesquisa_global}'** ainda não consta na planilha da Assertec.")
                    except Exception:
                        st.warning("Não foi possível ler a planilha automaticamente.")

        st.markdown("---")
        
        if st.session_state.get("mensagem_processada", False):
            formulario_escolhido = st.session_state["formulario_escolhido"]
            st.info(f"📋 Fluxo ativo: **{formulario_escolhido}** | Técnico da Mensagem: **{st.session_state.get('tec_msg', 'N/D')}**")

            st.subheader(f"📋 Relatório de Conferência - {formulario_escolhido}")
            
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                st.markdown(f"**Cliente:** {st.session_state.get('cli', 'N/D')}")
                st.markdown(f"**Protocolo:** {st.session_state.get('proto', 'N/D')}")
                st.markdown(f"**Cidade:** {st.session_state.get('cidade', 'N/D')}")
            with col_d2:
                st.markdown(f"**Serviço:** {st.session_state.get('servico', 'N/D')}")
                st.markdown(f"**Técnico:** {st.session_state.get('tec', 'N/D')}")
                st.markdown(f"**Quantidade / Metragem:** {st.session_state.get('qtd', '0')}")
            
            st.markdown(f"**Observações Finais:** {st.session_state.get('obs', 'Nenhuma observação.')}")
        else:
            st.info("ℹ️ Cole e processe uma mensagem do WhatsApp na caixa acima para visualizar o relatório de conferência.")

# ==========================================
# ABA 2: GERADOR DE MENSAGENS (Telemóvel)
# ==========================================
with aba_gerador:
    st.subheader("📱 Gerador de Mensagem Padronizada para WhatsApp")
    st.markdown("Selecione a empresa e preencha todos os campos obrigatórios para liberar a geração da mensagem.")

    opcoes_empresa = ["Escolher...", "ASSERTEC", "AMIGO"]
    empresa_gen = selectbox_com_borda_dinamica("Empresa / Tipo *", opcoes_empresa, "gen_empresa")

    campos_preenchidos = False

    if empresa_gen != "Escolher...":
        if empresa_gen == "ASSERTEC":
            st.error("🚨 **Atenção: Preencha todos os campos obrigatórios em destaque abaixo.**")
            
            tec_opcoes = ["Escolher...", "ANDERSON", "MARCIO"]
            tecnico_selecionado = selectbox_com_borda_dinamica("Técnico Responsável *", tec_opcoes, "gen_tecnico_assertec")
            
            cidades_opcoes = ["Escolher..."] + lista_cidades_assertec
            cidade_gen = selectbox_com_borda_dinamica("Cidade *", cidades_opcoes, "gen_cidade_assertec")
            
            servicos_opcoes = lista_servicos_assertec_global 
            servico_gen = selectbox_com_borda_dinamica("Serviço Realizado *", servicos_opcoes, "gen_servico_assertec")
            
            cliente_gen = st.text_input("Nome do Cliente *", placeholder="EX: ELISETE ANTONIA DOS SANTOS", key="gen_cliente_assertec")
            protocolo_gen = st.text_input("Número do Protocolo *", placeholder="EX: 114619", key="gen_protocolo_assertec")
            metragem_gen = st.number_input("Metragem / Quantidade", min_value=0, value=1, step=1, key="gen_metragem_assertec")
            
            t_ok = tecnico_selecionado != "Escolher..."
            c_ok = cidade_gen != "Escolher..."
            s_ok = servico_gen != "Escolher..."
            cli_ok = cliente_gen.strip() != "" and cliente_gen.strip() != "EX: ELISETE ANTONIA DOS SANTOS"
            p_ok = protocolo_gen.strip() != "" and protocolo_gen.strip() != "EX: 114619"
            
            if t_ok and c_ok and s_ok and cli_ok and p_ok:
                campos_preenchidos = True

        else: # AMIGO
            st.error("🚨 **Atenção: Preencha todos os campos obrigatórios em destaque abaixo.**")
            
            tec_opcoes = ["Escolher...", "ANDERSON", "MARCIO"]
            tecnico_selecionado = selectbox_com_borda_dinamica("Técnico Responsável *", tec_opcoes, "gen_tecnico_amigo")
            
            cidades_opcoes = ["Escolher..."] + lista_cidades_assertec
            cidade_gen = selectbox_com_borda_dinamica("Cidade *", cidades_opcoes, "gen_cidade_amigo")
            
            servicos_opcoes = lista_servicos_amigo_global 
            servico_gen = selectbox_com_borda_dinamica("Serviço Realizado *", servicos_opcoes, "gen_servico_amigo")
            
            cliente_gen = st.text_input("Nome do Cliente *", placeholder="EX: ELISETE ANTONIA DOS SANTOS", key="gen_cliente_amigo")
            protocolo_gen = st.text_input("Número do Protocolo *", placeholder="EX: 114619", key="gen_protocolo_amigo")
            
            lpu_gen = st.number_input("LPU", min_value=0, value=1, step=1, key="gen_lpu_amigo")
            metragem_amigo_gen = st.number_input("Metragem / Cabo", min_value=0, value=0, step=1, key="gen_metragem_amigo")
            km_rodado_gen = 0
            st.text_input("Quilômetro Rodado (Km)", value="0", disabled=True, key="gen_km_amigo_fixo")
            
            t_ok = tecnico_selecionado != "Escolher..."
            c_ok = cidade_gen != "Escolher..."
            s_ok = servico_gen != "Escolher..."
            cli_ok = cliente_gen.strip() != "" and cliente_gen.strip() != "EX: ELISETE ANTONIA DOS SANTOS"
            p_ok = protocolo_gen.strip() != "" and protocolo_gen.strip() != "EX: 114619"
            
            if t_ok and c_ok and s_ok and cli_ok and p_ok:
                campos_preenchidos = True

        cli_valido = cliente_gen.strip() != "" and cliente_gen.strip() != "EX: ELISETE ANTONIA DOS SANTOS"
        proto_valido = protocolo_gen.strip() != "" and protocolo_gen.strip() != "EX: 114619"
        
        col_cli = "#2e7d32" if cli_valido else "#ff4b4b"
        col_proto = "#2e7d32" if proto_valido else "#ff4b4b"

        st.markdown(f"""
            <style>
            input[aria-label="Nome do Cliente *"] {{
                border: 2px solid {col_cli} !important;
            }}
            input[aria-label="Número do Protocolo *"] {{
                border: 2px solid {col_proto} !important;
            }}
            </style>
        """, unsafe_allow_html=True)

        st.markdown("---")
        obs_gen = st.text_area("Observações (Opcional)", placeholder="EX: Refeita configurações nos equipamentos...", key="gen_obs")

        if tecnico_selecionado == "ANDERSON":
            tecnico_gen = "ANDERSON"
            nome_tec_completo = "ANDERSON WOLFART MOLLERKE" if empresa_gen == "AMIGO" else "Anderson"
        else:
            tecnico_gen = "MARCIO"
            nome_tec_completo = "ANDERSON WOLFART MOLLERKE" if empresa_gen == "AMIGO" else "Márcio"

        if campos_preenchidos:
            if st.button("✨ Gerar Mensagem Pronta"):
                if empresa_gen == "AMIGO":
                    mensagem_gerada = f"""{empresa_gen}
{tecnico_gen}
Técnico: {nome_tec_completo}
Cliente: {cliente_gen.upper()}
Protocolo: {protocolo_gen}
Cidade: {cidade_gen}
Serviço: {servico_gen}
LPU: {lpu_gen}
Metragem: {metragem_amigo_gen}
Km Rodado: {km_rodado_gen}
Obs: {obs_gen}"""
                else:
                    mensagem_gerada = f"""{empresa_gen}
{tecnico_gen}
Técnico: {nome_tec_completo}
Cliente: {cliente_gen.upper()}
Protocolo: {protocolo_gen}
Cidade: {cidade_gen}
Serviço: {servico_gen}
Metragem: {metragem_gen}
Obs: {obs_gen}"""
                st.session_state["mensagem_gerada_final"] = mensagem_gerada
        else:
            st.warning("⚠️ Por favor, preencha todos os campos obrigatórios destacados para liberar a geração da mensagem.")

    if "mensagem_gerada_final" in st.session_state and st.session_state["mensagem_gerada_final"]:
        st.markdown("### 📋 Mensagem Gerada:")
        st.code(st.session_state["mensagem_gerada_final"], language="text")
        
        mensagem_url = urllib.parse.quote(st.session_state["mensagem_gerada_final"])
        link_whatsapp = f"https://wa.me/?text={mensagem_url}"
        st.link_button("📲 Enviar diretamente para o WhatsApp", link_whatsapp)

# ==========================================
# ABA 3: CENTRAL DE AGENDAMENTOS (NOME, CLIENTE, PROTOCOLO, CIDADE, OBS)
# ==========================================
with aba_agendamentos:
    data_amanha_dt = datetime.now() + timedelta(days=1)
    data_amanha_str = data_amanha_dt.strftime("%d/%m/%Y")

    st.subheader(f"🔔 Agendamentos para Amanhã ({data_amanha_str})")
    
    df_agend = carregar_agendamentos()
    if not df_agend.empty and "OBS" in df_agend.columns:
        df_amanha = df_agend[
            df_agend.apply(lambda r: data_amanha_str in str(r.values), axis=1)
        ]
        if not df_amanha.empty:
            st.success(f"📌 Existem **{len(df_amanha)}** atendimento(s) agendado(s) para amanhã ({data_amanha_str}):")
            for idx, row in df_amanha.iterrows():
                with st.container():
                    st.markdown(
                        f"""
                        <div style="border-left: 5px solid #2e7d32; background-color: rgba(46, 125, 50, 0.06); padding: 10px 14px; border-radius: 6px; margin-bottom: 8px;">
                            <strong>🏙️ Cidade: {row.get('CIDADE', 'N/D')} &nbsp;|&nbsp; 👤 Registado por: {row.get('NOME', 'N/D')}</strong><br>
                            👤 <strong>Cliente:</strong> {row.get('CLIENTE', 'N/D')} &nbsp;|&nbsp; 🔢 <strong>Protocolo:</strong> {row.get('PROTOCOLO', 'N/D')}<br>
                            📝 <strong>Obs / Data:</strong> {row.get('OBS', '-')}
                        </div>
                        """,
                        unsafe_allow_html=True
                    )
        else:
            st.info(f"✅ Nenhum agendamento registado para amanhã ({data_amanha_str}) até ao momento.")
    else:
        st.info(f"✅ Nenhum agendamento registado para amanhã ({data_amanha_str}) até ao momento.")

    with st.expander("📆 Ver Todos os Agendamentos da Planilha"):
        if not df_agend.empty:
            st.dataframe(df_agend, use_container_width=True)
        else:
            st.write("Ainda não há agendamentos na planilha.")

    st.markdown("---")
    st.subheader("📝 Novo Agendamento")
    st.markdown("Preencha os campos obrigatórios para gravar na planilha e gerar a mensagem para o grupo do WhatsApp.")

    col_data, col_nome = st.columns(2)
    with col_data:
        data_escolhida = st.date_input("Data do Agendamento *", value=data_amanha_dt, format="DD/MM/YYYY", key="ag_data")
    with col_nome:
        nome_ag = st.text_input("Nome (Quem agendou / Técnico) *", placeholder="EX: MAIARA / MARCIO", key="ag_nome")

    cliente_ag = st.text_input("Cliente (Agendamento) *", placeholder="EX: JOÃO DA SILVA", key="ag_cliente")
    protocolo_ag = st.text_input("Protocolo (Agendamento) *", placeholder="EX: 2129988", key="ag_protocolo")
    cidade_ag = selectbox_com_borda_dinamica("Cidade *", ["Escolher..."] + lista_cidades_assertec, "ag_cidade")
    obs_ag = st.text_area("Observações (Turno, Endereço, Detalhes)", placeholder="EX: Manhã - Rua das Flores, 123 - Ligar antes", key="ag_obs")

    # Validação visual dos campos obrigatórios (Vermelho se vazio, Verde se preenchido)
    nome_ag_ok = nome_ag.strip() != ""
    cli_ag_ok = cliente_ag.strip() != ""
    proto_ag_ok = protocolo_ag.strip() != ""
    cid_ag_ok = cidade_ag != "Escolher..."

    cor_nome_ag = "#2e7d32" if nome_ag_ok else "#ff4b4b"
    cor_cli_ag = "#2e7d32" if cli_ag_ok else "#ff4b4b"
    cor_proto_ag = "#2e7d32" if proto_ag_ok else "#ff4b4b"

    st.markdown(f"""
        <style>
        input[aria-label="Nome (Quem agendou / Técnico) *"] {{
            border: 2px solid {cor_nome_ag} !important;
        }}
        input[aria-label="Cliente (Agendamento) *"] {{
            border: 2px solid {cor_cli_ag} !important;
        }}
        input[aria-label="Protocolo (Agendamento) *"] {{
            border: 2px solid {cor_proto_ag} !important;
        }}
        </style>
    """, unsafe_allow_html=True)

    agendamento_valido = nome_ag_ok and cli_ag_ok and proto_ag_ok and cid_ag_ok

    if agendamento_valido:
        if st.button("💾 Salvar na Planilha e Gerar Mensagem"):
            data_fmt = data_escolhida.strftime("%d/%m/%Y")
            obs_final_planilha = f"[{data_fmt}] {obs_ag.strip()}" if obs_ag.strip() else f"[{data_fmt}]"
            
            dados_novo_ag = {
                "nome": nome_ag.strip().upper(),
                "cliente": cliente_ag.strip().upper(),
                "protocolo": protocolo_ag.strip(),
                "cidade": cidade_ag,
                "obs": obs_final_planilha
            }
            
            ok, msg_retorno = salvar_agendamento(dados_novo_ag)
            if ok:
                st.success(f"✅ {msg_retorno}")

            msg_whatsapp_ag = f"""📅 *AGENDAMENTO ({data_fmt})*
👤 *Nome:* {nome_ag.strip().upper()}
🏠 *Cliente:* {cliente_ag.strip().upper()}
🔢 *Protocolo:* {protocolo_ag.strip()}
🏙️ *Cidade:* {cidade_ag}
📝 *Obs:* {obs_ag.strip() or 'Nenhuma'}"""
            st.session_state["msg_agendamento_gerada"] = msg_whatsapp_ag
    else:
        st.warning("⚠️ Preencha todos os campos obrigatórios (em vermelho) para liberar o registo do agendamento.")

    if "msg_agendamento_gerada" in st.session_state and st.session_state["msg_agendamento_gerada"]:
        st.markdown("### 📋 Mensagem de Agendamento Pronta:")
        st.code(st.session_state["msg_agendamento_gerada"], language="text")
        
        msg_ag_url = urllib.parse.quote(st.session_state["msg_agendamento_gerada"])
        link_wpp_ag = f"https://wa.me/?text={msg_ag_url}"
        st.link_button("📲 Enviar para o Grupo do WhatsApp", link_wpp_ag)