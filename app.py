import streamlit as st
import time
import re
import io
import csv
import html
import os
import json
import urllib.parse
import urllib.request
from datetime import datetime, timedelta

IS_CLOUD = os.path.exists("/mount/src") or os.environ.get("STREAMLIT_RUNTIME_ENV") == "cloud"

if not IS_CLOUD:
    try:
        import pandas as pd
        import webbrowser
        import openpyxl
        import pyautogui
        import pyperclip
    except ImportError:
        IS_CLOUD = True
        pd = None
        webbrowser = None
        openpyxl = None
        pyautogui = None
        pyperclip = None
else:
    pd = None
    webbrowser = None
    openpyxl = None
    pyautogui = None
    pyperclip = None

# ==============================================================================
# CONFIGURAÇÃO DA PLANILHA DE AGENDAMENTOS (GOOGLE SHEETS)
# ==============================================================================
SHEET_ID_AGENDAMENTOS = "1jXUw3oZKRPRXV53YCoqBxHUZN95igzpRN_r7yAhhoTk"
URL_PLANILHA_AGENDAMENTOS_CSV = f"https://docs.google.com/spreadsheets/d/{SHEET_ID_AGENDAMENTOS}/gviz/tq?tqx=out:csv"
GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyaSys2BNpDUVeM74BPvDxmseMYo2L57LnaZspAYIhBZ-qpaXIArbEqrvp_Cb1IhigF/exec"
ARQUIVO_AGENDAMENTOS_LOCAL = "agendamentos_backup.csv"
COLUNAS_PADRAO_AG = ["DATA", "CLIENTE", "PROTOCOLO", "CIDADE", "OBS"]

st.set_page_config(page_title="Sistema de Atendimento - FibraWer", layout="centered")

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

def selectbox_com_borda_dinamica(label, opcoes, chave):
    if chave not in st.session_state:
        st.session_state[chave] = opcoes[0]
    valor_atual = st.session_state.get(chave, opcoes[0])
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

def limpar_protocolo_str(val):
    texto = str(val or "").strip()
    if texto.endswith(".0"):
        texto = texto[:-2]
    return texto

def normalizar_lista_agendamentos(linhas_raw):
    resultado = []
    if not linhas_raw:
        return resultado
    for item in linhas_raw:
        if not isinstance(item, dict):
            continue
        norm = {str(k or "").strip().upper(): str(v or "").strip() for k, v in item.items()}
        if "DATA" not in norm and "NOME" in norm:
            norm["DATA"] = norm.get("NOME", "")
        reg = {
            "DATA": norm.get("DATA", ""),
            "CLIENTE": norm.get("CLIENTE", ""),
            "PROTOCOLO": limpar_protocolo_str(norm.get("PROTOCOLO", "")),
            "CIDADE": norm.get("CIDADE", ""),
            "OBS": norm.get("OBS", "")
        }
        if reg["CLIENTE"] != "" or reg["PROTOCOLO"] != "":
            resultado.append(reg)
    return resultado

def salvar_backup_csv(lista_ag):
    try:
        with open(ARQUIVO_AGENDAMENTOS_LOCAL, mode="w", encoding="utf-8", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=COLUNAS_PADRAO_AG)
            writer.writeheader()
            for row in lista_ag:
                writer.writerow({c: row.get(c, "") for c in COLUNAS_PADRAO_AG})
    except Exception:
        pass

def ler_csv_texto(conteudo_csv):
    leitor = csv.DictReader(io.StringIO(conteudo_csv))
    return normalizar_lista_agendamentos(list(leitor))

def buscar_agendamentos_remoto():
    try:
        url_csv = f"{URL_PLANILHA_AGENDAMENTOS_CSV}&t={int(time.time())}"
        req = urllib.request.Request(url_csv, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=4) as resp:
            conteudo_csv = resp.read().decode("utf-8", errors="replace")
        lista = ler_csv_texto(conteudo_csv)
        salvar_backup_csv(lista)
        return lista
    except Exception:
        if os.path.exists(ARQUIVO_AGENDAMENTOS_LOCAL):
            try:
                with open(ARQUIVO_AGENDAMENTOS_LOCAL, mode="r", encoding="utf-8") as f:
                    return ler_csv_texto(f.read())
            except Exception:
                pass
    return []

def obter_df_agendamentos():
    if "df_agendamentos_mem" not in st.session_state:
        st.session_state["df_agendamentos_mem"] = buscar_agendamentos_remoto()
    return st.session_state["df_agendamentos_mem"]

def atualizar_memoria_e_backup(lista_nova):
    lista_limpa = normalizar_lista_agendamentos(lista_nova)
    st.session_state["df_agendamentos_mem"] = lista_limpa
    salvar_backup_csv(lista_limpa)

def enviar_post_sheets(payload):
    if not GOOGLE_SCRIPT_URL.strip():
        return False, "URL do Apps Script não configurada."
    try:
        req = urllib.request.Request(
            GOOGLE_SCRIPT_URL.strip(),
            data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json"},
            method="POST"
        )
        with urllib.request.urlopen(req, timeout=8) as resp:
            return True, "Sincronizado com a Planilha Google!"
    except Exception as e:
        return False, f"Aviso Google Sheets: {e}"

def extrair_data_da_linha(row):
    val_data = str(row.get("DATA", "")).strip()
    if re.search(r"\d{2}/\d{2}/\d{4}", val_data):
        m = re.search(r"(\d{2}/\d{2}/\d{4})", val_data)
        return m.group(1)
    val_obs = str(row.get("OBS", ""))
    match = re.search(r"(\d{2}/\d{2}/\d{4})", val_obs)
    if match:
        return match.group(1)
    return val_data

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
            except Exception:
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
            except Exception:
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
    st.markdown("Selecione a empresa e preencha os campos obrigatórios para gerar a mensagem.")

    opcoes_empresa = ["Escolher...", "ASSERTEC", "AMIGO"]
    empresa_gen = selectbox_com_borda_dinamica("Empresa / Tipo *", opcoes_empresa, "gen_empresa")

    if empresa_gen != "Escolher...":
        with st.form(f"form_gerador_{empresa_gen}", clear_on_submit=False):
            st.caption("Preencha os campos abaixo e toque em Gerar Mensagem Pronta:")
            tec_opcoes = ["Escolher...", "ANDERSON", "MARCIO"]
            cidades_opcoes = ["Escolher..."] + lista_cidades_assertec

            if empresa_gen == "ASSERTEC":
                tecnico_selecionado = st.selectbox("Técnico Responsável *", tec_opcoes, key="gen_tecnico_assertec")
                cidade_gen = st.selectbox("Cidade *", cidades_opcoes, key="gen_cidade_assertec")
                servico_gen = st.selectbox("Serviço Realizado *", lista_servicos_assertec_global, key="gen_servico_assertec")
                cliente_gen = st.text_input("Nome do Cliente *", placeholder="EX: ELISETE ANTONIA DOS SANTOS", key="gen_cliente_assertec")
                protocolo_gen = st.text_input("Número do Protocolo *", placeholder="EX: 114619", key="gen_protocolo_assertec")
                metragem_gen = st.number_input("Metragem / Quantidade", min_value=0, value=1, step=1, key="gen_metragem_assertec")
                obs_gen = st.text_area("Observações (Opcional)", placeholder="EX: Refeita configurações nos equipamentos...", key="gen_obs_assertec")
            else:
                tecnico_selecionado = st.selectbox("Técnico Responsável *", tec_opcoes, key="gen_tecnico_amigo")
                cidade_gen = st.selectbox("Cidade *", cidades_opcoes, key="gen_cidade_amigo")
                servico_gen = st.selectbox("Serviço Realizado *", lista_servicos_amigo_global, key="gen_servico_amigo")
                cliente_gen = st.text_input("Nome do Cliente *", placeholder="EX: ELISETE ANTONIA DOS SANTOS", key="gen_cliente_amigo")
                protocolo_gen = st.text_input("Número do Protocolo *", placeholder="EX: 114619", key="gen_protocolo_amigo")
                lpu_gen = st.number_input("LPU", min_value=0, value=1, step=1, key="gen_lpu_amigo")
                metragem_amigo_gen = st.number_input("Metragem / Cabo", min_value=0, value=0, step=1, key="gen_metragem_amigo")
                km_rodado_gen = 0
                st.text_input("Quilômetro Rodado (Km)", value="0", disabled=True, key="gen_km_amigo_fixo")
                obs_gen = st.text_area("Observações (Opcional)", placeholder="EX: Refeita configurações nos equipamentos...", key="gen_obs_amigo")

            clicou_gerar_msg = st.form_submit_button("✨ Gerar Mensagem Pronta", type="primary", use_container_width=True)

        if clicou_gerar_msg:
            t_ok = tecnico_selecionado != "Escolher..."
            c_ok = cidade_gen != "Escolher..."
            s_ok = servico_gen != "Escolher..."
            cli_ok = cliente_gen.strip() != ""
            p_ok = protocolo_gen.strip() != ""

            if not (t_ok and c_ok and s_ok and cli_ok and p_ok):
                st.error("⚠️ Preencha todos os campos obrigatórios (*) antes de gerar a mensagem.")
            else:
                if tecnico_selecionado == "ANDERSON":
                    tecnico_gen = "ANDERSON"
                    nome_tec_completo = "ANDERSON WOLFART MOLLERKE" if empresa_gen == "AMIGO" else "Anderson"
                else:
                    tecnico_gen = "MARCIO"
                    nome_tec_completo = "ANDERSON WOLFART MOLLERKE" if empresa_gen == "AMIGO" else "Márcio"

                obs_texto_limpo = obs_gen.strip()
                if empresa_gen == "AMIGO":
                    mensagem_gerada = f"""{empresa_gen}
{tecnico_gen}
Técnico: {nome_tec_completo}
Cliente: {cliente_gen.strip().upper()}
Protocolo: {protocolo_gen.strip()}
Cidade: {cidade_gen}
Serviço: {servico_gen}
LPU: {lpu_gen}
Metragem: {metragem_amigo_gen}
Km Rodado: {km_rodado_gen}
Obs: {obs_texto_limpo}"""
                else:
                    mensagem_gerada = f"""{empresa_gen}
{tecnico_gen}
Técnico: {nome_tec_completo}
Cliente: {cliente_gen.strip().upper()}
Protocolo: {protocolo_gen.strip()}
Cidade: {cidade_gen}
Serviço: {servico_gen}
Metragem: {metragem_gen}
Obs: {obs_texto_limpo}"""
                st.session_state["mensagem_gerada_final"] = mensagem_gerada

    if "mensagem_gerada_final" in st.session_state and st.session_state["mensagem_gerada_final"]:
        st.markdown("### 📋 Mensagem Gerada:")
        st.code(st.session_state["mensagem_gerada_final"], language="text")
        
        mensagem_url = urllib.parse.quote(st.session_state["mensagem_gerada_final"])
        link_whatsapp = f"https://wa.me/?text={mensagem_url}"
        st.link_button("📲 Enviar diretamente para o WhatsApp", link_whatsapp, use_container_width=True)

# ==========================================
# ABA 3: CENTRAL DE AGENDAMENTOS (PC & SMARTPHONE)
# ==========================================
with aba_agendamentos:
    data_amanha_dt = (datetime.now() + timedelta(days=1)).date()
    data_amanha_str = data_amanha_dt.strftime("%d/%m/%Y")

    if "ag_data" not in st.session_state:
        st.session_state["ag_data"] = data_amanha_dt
    if "ag_nome" not in st.session_state:
        st.session_state["ag_nome"] = ""
    if "ag_cliente" not in st.session_state:
        st.session_state["ag_cliente"] = ""
    if "ag_protocolo" not in st.session_state:
        st.session_state["ag_protocolo"] = ""
    if "ag_cidade" not in st.session_state:
        st.session_state["ag_cidade"] = "Escolher..."
    if "ag_obs" not in st.session_state:
        st.session_state["ag_obs"] = ""
    if "modo_edicao_idx" not in st.session_state:
        st.session_state["modo_edicao_idx"] = None
    if "modo_edicao_proto_original" not in st.session_state:
        st.session_state["modo_edicao_proto_original"] = None
    if "modo_exclusao_idx" not in st.session_state:
        st.session_state["modo_exclusao_idx"] = None
    if "aviso_topo_ag" not in st.session_state:
        st.session_state["aviso_topo_ag"] = None
    if "erro_form_ag" not in st.session_state:
        st.session_state["erro_form_ag"] = None

    def cb_acao_card(chave_select, idx_linha, proto_item, data_item, cli_item, cid_item, obs_item):
        acao = st.session_state.get(chave_select, "Ações...")
        st.session_state[chave_select] = "Ações..."
        st.session_state["erro_form_ag"] = None

        if acao == "✅ Concluído":
            lista_mem = [dict(r) for r in obter_df_agendamentos()]
            if 0 <= idx_linha < len(lista_mem):
                lista_mem[idx_linha]["OBS"] = "✅ [CONCLUÍDO] " + str(lista_mem[idx_linha].get("OBS", ""))
                atualizar_memoria_e_backup(lista_mem)
            enviar_post_sheets({"acao": "concluir", "protocolo": str(proto_item).strip()})
            st.session_state["modo_exclusao_idx"] = None
            st.session_state["modo_edicao_idx"] = None
            st.session_state["aviso_topo_ag"] = ("sucesso", f"✅ Atendimento de **{cli_item}** marcado como CONCLUÍDO (verde na planilha)!")

        elif acao == "🗑️ Excluir":
            st.session_state["modo_exclusao_idx"] = idx_linha
            st.session_state["modo_edicao_idx"] = None

        elif acao == "✏️ Editar":
            st.session_state["modo_edicao_idx"] = idx_linha
            st.session_state["modo_edicao_proto_original"] = str(proto_item).strip()
            st.session_state["modo_exclusao_idx"] = None

            try:
                st.session_state["ag_data"] = datetime.strptime(data_item, "%d/%m/%Y").date()
            except Exception:
                st.session_state["ag_data"] = data_amanha_dt

            match_nome = re.match(r"^\((.*?)\)\s*(.*)", str(obs_item).strip(), flags=re.DOTALL)
            if match_nome:
                nome_extraido = re.sub(r"\s+AGENDOU$", "", match_nome.group(1).strip(), flags=re.IGNORECASE).strip()
                st.session_state["ag_nome"] = nome_extraido
                st.session_state["ag_obs"] = match_nome.group(2).strip()
            else:
                if not st.session_state.get("ag_nome"):
                    st.session_state["ag_nome"] = "MARCIA"
                st.session_state["ag_obs"] = str(obs_item).strip()

            st.session_state["ag_cliente"] = str(cli_item).strip()
            st.session_state["ag_protocolo"] = str(proto_item).strip()
            st.session_state["ag_cidade"] = cid_item if cid_item in lista_cidades_assertec else "Escolher..."
            st.session_state["aviso_topo_ag"] = ("info", f"✏️ Modo de Edição ativado para **{cli_item}**. Altere os dados no formulário abaixo e clique em **🔄 Relançar Agendamento**.")

    def cb_confirmar_exclusao(idx_linha, proto_item, cli_item):
        lista_mem = [dict(r) for r in obter_df_agendamentos()]
        if 0 <= idx_linha < len(lista_mem):
            lista_mem.pop(idx_linha)
            atualizar_memoria_e_backup(lista_mem)
        enviar_post_sheets({"acao": "excluir", "protocolo": str(proto_item).strip()})
        st.session_state["modo_exclusao_idx"] = None
        st.session_state["aviso_topo_ag"] = ("sucesso", f"🗑️ Agendamento de **{cli_item}** excluído com sucesso!")

    def cb_cancelar_exclusao():
        st.session_state["modo_exclusao_idx"] = None

    def cb_cancelar_edicao():
        st.session_state["modo_edicao_idx"] = None
        st.session_state["modo_edicao_proto_original"] = None
        st.session_state["ag_cliente"] = ""
        st.session_state["ag_protocolo"] = ""
        st.session_state["ag_obs"] = ""
        st.session_state["ag_cidade"] = "Escolher..."
        st.session_state["aviso_topo_ag"] = None
        st.session_state["erro_form_ag"] = None

    def cb_salvar_ou_relancar():
        data_escolhida = st.session_state.get("ag_data", data_amanha_dt)
        nome_bruto = st.session_state.get("ag_nome", "").strip().upper()
        nome_val = re.sub(r"\s+AGENDOU$", "", nome_bruto, flags=re.IGNORECASE).strip()
        cli_val = st.session_state.get("ag_cliente", "").strip().upper()
        proto_val = limpar_protocolo_str(st.session_state.get("ag_protocolo", ""))
        cid_val = st.session_state.get("ag_cidade", "Escolher...")
        obs_limpa = st.session_state.get("ag_obs", "").strip()

        if not nome_val or not cli_val or not proto_val or cid_val == "Escolher...":
            st.session_state["erro_form_ag"] = "⚠️ Preencha todos os campos obrigatórios (*): Nome, Cliente, Protocolo e Cidade."
            return

        st.session_state["erro_form_ag"] = None
        data_fmt = data_escolhida.strftime("%d/%m/%Y")
        prefixo_agendador = f"({nome_val} AGENDOU)"
        obs_final_planilha = f"{prefixo_agendador} {obs_limpa}" if obs_limpa else prefixo_agendador

        lista_mem = [dict(r) for r in obter_df_agendamentos()]
        idx_em_edicao = st.session_state.get("modo_edicao_idx")
        proto_antigo = st.session_state.get("modo_edicao_proto_original")

        novo_reg = {
            "DATA": data_fmt,
            "CLIENTE": cli_val,
            "PROTOCOLO": proto_val,
            "CIDADE": cid_val,
            "OBS": obs_final_planilha
        }

        if idx_em_edicao is not None:
            if 0 <= idx_em_edicao < len(lista_mem):
                lista_mem[idx_em_edicao] = novo_reg
            else:
                lista_mem.append(novo_reg)

            atualizar_memoria_e_backup(lista_mem)
            enviar_post_sheets({
                "acao": "editar",
                "protocolo_antigo": str(proto_antigo or proto_val).strip(),
                "data": data_fmt,
                "nome": data_fmt,
                "cliente": cli_val,
                "protocolo": proto_val,
                "cidade": cid_val,
                "obs": obs_final_planilha
            })
            st.session_state["modo_edicao_idx"] = None
            st.session_state["modo_edicao_proto_original"] = None
            st.session_state["aviso_topo_ag"] = ("sucesso", f"✅ Agendamento de **{cli_val}** atualizado e relançado para **{data_fmt}**!")
        else:
            ja_existe = any(str(r.get("PROTOCOLO", "")).strip() == proto_val for r in lista_mem)
            if ja_existe:
                st.session_state["aviso_topo_ag"] = ("aviso", f"⚠️ O protocolo **{proto_val}** já consta na lista! Use a opção '✏️ Editar' se quiser alterá-lo.")
                return

            lista_mem.append(novo_reg)
            atualizar_memoria_e_backup(lista_mem)

            enviar_post_sheets({
                "data": data_fmt,
                "nome": data_fmt,
                "cliente": cli_val,
                "protocolo": proto_val,
                "cidade": cid_val,
                "obs": obs_final_planilha
            })
            st.session_state["aviso_topo_ag"] = ("sucesso", f"✅ Agendamento de **{cli_val}** para o dia **{data_fmt}** salvo na planilha!")

        st.session_state["msg_agendamento_gerada"] = f"""📅 *AGENDAMENTO ({data_fmt})*
👤 *Agendado por:* {nome_val}
🏠 *Cliente:* {cli_val}
🔢 *Protocolo:* {proto_val}
🏙️ *Cidade:* {cid_val}
📝 *Obs:* {obs_limpa or 'Nenhuma'}"""

        st.session_state["ag_cliente"] = ""
        st.session_state["ag_protocolo"] = ""
        st.session_state["ag_obs"] = ""
        st.session_state["ag_cidade"] = "Escolher..."

    def cb_sincronizar_planilha():
        st.session_state["df_agendamentos_mem"] = buscar_agendamentos_remoto()
        st.session_state["modo_exclusao_idx"] = None
        st.session_state["modo_edicao_idx"] = None
        st.session_state["erro_form_ag"] = None
        st.session_state["aviso_topo_ag"] = ("sucesso", "🔄 Lista sincronizada com a Planilha Google!")

    # --- TOPO DA ABA DE AGENDAMENTOS ---
    col_tit_ag, col_sync_ag = st.columns([3, 1])
    with col_tit_ag:
        st.subheader(f"🔔 Agendamentos para Amanhã ({data_amanha_str})")
    with col_sync_ag:
        st.button("🔄 Atualizar", on_click=cb_sincronizar_planilha, key="btn_sync_sheets")

    if st.session_state.get("aviso_topo_ag"):
        tipo_av, texto_av = st.session_state["aviso_topo_ag"]
        if tipo_av == "sucesso":
            st.success(texto_av)
        elif tipo_av == "aviso":
            st.warning(texto_av)
        else:
            st.info(texto_av)

    lista_agend = obter_df_agendamentos()

    if lista_agend:
        itens_amanha = []
        itens_outros_dias = []
        for idx, row in enumerate(lista_agend):
            obs_str = str(row.get("OBS", ""))
            if "CONCLUÍDO" in obs_str.upper():
                continue
            data_real = extrair_data_da_linha(row)
            copia_row = dict(row)
            copia_row["DATA_REAL"] = data_real
            if data_real == data_amanha_str:
                itens_amanha.append((idx, copia_row))
            else:
                itens_outros_dias.append((idx, copia_row))

        def renderizar_lista_agendamentos(lista_pares, prefixo):
            for idx, row in lista_pares:
                proto_item = limpar_protocolo_str(row.get('PROTOCOLO', ''))
                cidade_item = str(row.get('CIDADE', 'N/D'))
                cliente_item = str(row.get('CLIENTE', 'N/D'))
                obs_item = str(row.get('OBS', '-'))
                data_item = str(row.get('DATA_REAL', 'N/D'))
                chave_select = f"acao_sel_{prefixo}_{idx}"

                if chave_select not in st.session_state:
                    st.session_state[chave_select] = "Ações..."

                cor_barra = "#2e7d32" if prefixo == "ama" else "#1976d2"
                bg_card = "rgba(46, 125, 50, 0.06)" if prefixo == "ama" else "rgba(25, 118, 210, 0.05)"

                with st.container():
                    col_info, col_menu = st.columns([2.8, 1.2])
                    with col_info:
                        st.markdown(
                            f"""
                            <div style="border-left: 5px solid {cor_barra}; background-color: {bg_card}; padding: 8px 12px; border-radius: 6px; margin-bottom: 6px;">
                                <strong>🗓️ Data: {html.escape(data_item)} &nbsp;|&nbsp; 🏙️ Cidade: {html.escape(cidade_item)}</strong><br>
                                👤 <strong>Cliente:</strong> {html.escape(cliente_item)} &nbsp;|&nbsp; 🔢 <strong>Protocolo:</strong> {html.escape(proto_item)}<br>
                                📝 <strong>Obs:</strong> {html.escape(obs_item)}
                            </div>
                            """,
                            unsafe_allow_html=True
                        )
                    with col_menu:
                        st.selectbox(
                            "Interagir",
                            ["Ações...", "✅ Concluído", "✏️ Editar", "🗑️ Excluir"],
                            key=chave_select,
                            label_visibility="collapsed",
                            on_change=cb_acao_card,
                            args=(chave_select, idx, proto_item, data_item, cliente_item, cidade_item, obs_item)
                        )

                if st.session_state.get("modo_exclusao_idx") == idx:
                    st.warning(f"⚠️ **Tem certeza** que deseja excluir o agendamento de **{cliente_item}** (Protocolo: {proto_item})?")
                    col_sim, col_nao, _ = st.columns([1, 1, 2])
                    with col_sim:
                        st.button(
                            "Sim",
                            key=f"btn_sim_exc_{prefixo}_{idx}",
                            on_click=cb_confirmar_exclusao,
                            args=(idx, proto_item, cliente_item),
                            type="primary"
                        )
                    with col_nao:
                        st.button(
                            "Não",
                            key=f"btn_nao_exc_{prefixo}_{idx}",
                            on_click=cb_cancelar_exclusao
                        )

        if itens_amanha:
            st.success(f"📌 Existem **{len(itens_amanha)}** atendimento(s) agendado(s) para amanhã ({data_amanha_str}):")
            renderizar_lista_agendamentos(itens_amanha, "ama")
        else:
            st.info(f"✅ Nenhum agendamento pendente para amanhã ({data_amanha_str}).")

        if itens_outros_dias:
            with st.expander(f"📅 Ver Agendamentos de Outras Datas ({len(itens_outros_dias)})", expanded=True):
                renderizar_lista_agendamentos(itens_outros_dias, "out")
    else:
        st.info("✅ Nenhum agendamento pendente registado até ao momento.")

    with st.expander("📊 Ver Tabela Completa da Planilha (Incluindo Concluídos)"):
        if lista_agend:
            linhas_html = ""
            for r in lista_agend:
                linhas_html += (
                    f"<tr>"
                    f"<td style='padding:6px; border-bottom:1px solid #ddd;'>{html.escape(str(r.get('DATA', '')))}</td>"
                    f"<td style='padding:6px; border-bottom:1px solid #ddd;'>{html.escape(str(r.get('CLIENTE', '')))}</td>"
                    f"<td style='padding:6px; border-bottom:1px solid #ddd;'>{html.escape(str(r.get('PROTOCOLO', '')))}</td>"
                    f"<td style='padding:6px; border-bottom:1px solid #ddd;'>{html.escape(str(r.get('CIDADE', '')))}</td>"
                    f"<td style='padding:6px; border-bottom:1px solid #ddd;'>{html.escape(str(r.get('OBS', '')))}</td>"
                    f"</tr>"
                )
            tabela_html = f"""
            <div style="overflow-x:auto;">
                <table style="width:100%; border-collapse:collapse; font-size:0.85rem;">
                    <thead>
                        <tr style="background-color:#f0f2f6; text-align:left;">
                            <th style="padding:6px; border-bottom:2px solid #ccc;">DATA</th>
                            <th style="padding:6px; border-bottom:2px solid #ccc;">CLIENTE</th>
                            <th style="padding:6px; border-bottom:2px solid #ccc;">PROTOCOLO</th>
                            <th style="padding:6px; border-bottom:2px solid #ccc;">CIDADE</th>
                            <th style="padding:6px; border-bottom:2px solid #ccc;">OBS</th>
                        </tr>
                    </thead>
                    <tbody>{linhas_html}</tbody>
                </table>
            </div>
            """
            st.markdown(tabela_html, unsafe_allow_html=True)
        else:
            st.write("Ainda não há agendamentos na planilha.")

    st.markdown("---")

    idx_editando = st.session_state.get("modo_edicao_idx")
    proto_orig = st.session_state.get("modo_edicao_proto_original")
    if idx_editando is not None:
        st.subheader(f"✏️ Editando Agendamento (Protocolo: {proto_orig})")
        st.markdown("Altere os campos desejados abaixo e clique em **🔄 Relançar Agendamento**.")
    else:
        st.subheader("📝 Novo Agendamento")
        st.markdown("Preencha os campos obrigatórios para gravar na planilha e gerar a mensagem para o grupo do WhatsApp.")

    if st.session_state.get("erro_form_ag"):
        st.error(st.session_state["erro_form_ag"])

    with st.form("form_agendamento", clear_on_submit=False):
        col_data, col_nome = st.columns(2)
        with col_data:
            st.date_input("Data do Agendamento *", format="DD/MM/YYYY", key="ag_data")
        with col_nome:
            st.text_input("Nome (Quem agendou) *", placeholder="EX: JOAO / MARCIA / MAIARA", key="ag_nome")

        st.text_input("Cliente (Agendamento) *", placeholder="EX: TIO NENI", key="ag_cliente")
        st.text_input("Protocolo (Agendamento) *", placeholder="EX: 2129988", key="ag_protocolo")
        st.selectbox("Cidade *", ["Escolher..."] + lista_cidades_assertec, key="ag_cidade")
        st.text_area("Observações (Turno, Endereço, Detalhes)", placeholder="EX: Manhã - Ligar antes...", key="ag_obs")

        if idx_editando is not None:
            col_btn_rel, col_btn_canc = st.columns(2)
            with col_btn_rel:
                st.form_submit_button(
                    "🔄 Relançar Agendamento",
                    on_click=cb_salvar_ou_relancar,
                    type="primary",
                    use_container_width=True
                )
            with col_btn_canc:
                st.form_submit_button(
                    "❌ Cancelar Edição",
                    on_click=cb_cancelar_edicao,
                    use_container_width=True
                )
        else:
            st.form_submit_button(
                "💾 Salvar na Planilha e Gerar Mensagem",
                on_click=cb_salvar_ou_relancar,
                type="primary",
                use_container_width=True
            )

    if "msg_agendamento_gerada" in st.session_state and st.session_state["msg_agendamento_gerada"]:
        st.markdown("### 📋 Mensagem Gerada:")
        st.code(st.session_state["msg_agendamento_gerada"], language="text")
        
        msg_ag_url = urllib.parse.quote(st.session_state["msg_agendamento_gerada"])
        link_wpp_ag = f"https://wa.me/?text={msg_ag_url}"
        st.link_button("📲 Enviar para o Grupo do WhatsApp", link_wpp_ag, use_container_width=True)