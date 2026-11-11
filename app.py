import os
import re
import io
import csv
import html
import json
import time
import urllib.parse
import urllib.request
import webbrowser
from datetime import datetime, timedelta

import streamlit as st
import pandas as pd

try:
    import openpyxl
except ImportError:
    openpyxl = None

try:
    import pyautogui
    import pyperclip
    IS_CLOUD = False
except Exception:
    pyautogui = None
    pyperclip = None
    IS_CLOUD = True

# ==============================================================================
# IMPORTAÇÃO DOS MÓDULOS DO SISTEMA E COFRE DE CREDENCIAIS
# ==============================================================================
from config import (
    URL_PLANILHA_AGENDAMENTOS_CSV,
    GOOGLE_SCRIPT_URL,
    ARQUIVO_AGENDAMENTOS_LOCAL,
    ARQUIVO_PRINT_PORTAL,
    COLUNAS_PADRAO_AG,
    ARQUIVO_PRINT_SF,
)
from modulo_assertec import (
    lista_cidades_assertec,
    MAPA_BATIMENTO_CURTO,
    LISTA_UNIFICADA_TELEMOVEL,
    LISTA_AGREGADOS_TELEMOVEL,
    lista_servicos_assertec_global,
    EQUIPAMENTOS_LISTA,
    limpar_protocolo_str,
    obter_nome_oficial_servico,
    servico_vai_para_amigo,
    mapear_para_assertec,
    salvar_na_fila_assertec,
)
from modulo_planilhas import registrar_servico_na_planilha_tecnico
from modulo_robos import (
    carregar_agenda_sf_cache,
    sincronizar_agenda_salesforce,
    executar_lancamento_amigo_seguro,
)

# ==============================================================================
# CONFIGURAÇÃO VISUAL DA PÁGINA
# ==============================================================================
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

import streamlit.components.v1 as components

# Bloqueia a abertura do teclado virtual no telemóvel ao tocar em caixas de seleção (selectbox)
components.html(
    """
    <script>
    const bloquearTecladoSelect = () => {
        const doc = window.parent.document;
        const inputsSelect = doc.querySelectorAll('div[data-baseweb="select"] input');
        inputsSelect.forEach(inp => {
            inp.setAttribute('inputmode', 'none');
            inp.setAttribute('readonly', 'readonly');
        });
    };
    bloquearTecladoSelect();
    const obs = new MutationObserver(bloquearTecladoSelect);
    obs.observe(window.parent.document.body, { childList: true, subtree: true });
    </script>
    """,
    height=0,
)

# ==============================================================================
# FUNÇÕES DE APOIO PARA AGENDAMENTOS (CSV & GOOGLE SHEETS)
# ==============================================================================
def normalizar_lista_agendamentos(lista_bruta):
    lista_limpa = []
    for row in lista_bruta:
        if not isinstance(row, dict):
            continue
        novo = {}
        for col in COLUNAS_PADRAO_AG:
            val = ""
            for k, v in row.items():
                if k and str(k).strip().upper() == col:
                    val = str(v or "").strip()
                    break
            if col == "PROTOCOLO":
                val = limpar_protocolo_str(val)
            novo[col] = val
        if any(novo.values()):
            lista_limpa.append(novo)
    return lista_limpa

def salvar_backup_csv(lista):
    try:
        with open(ARQUIVO_AGENDAMENTOS_LOCAL, mode="w", encoding="utf-8", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=COLUNAS_PADRAO_AG)
            writer.writeheader()
            for r in lista:
                writer.writerow({col: r.get(col, "") for col in COLUNAS_PADRAO_AG})
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

# ==============================================================================
# CRIAÇÃO DAS ABAS PRINCIPAIS
# ==============================================================================
if IS_CLOUD:
    aba_gerador, aba_agendamentos = st.tabs([
        "📱 Lançamento Único (Telemóvel)",
        "📅 Agendamentos"
    ])
    aba_gestao = None
else:
    aba_gestao, aba_gerador, aba_agendamentos = st.tabs([
        "💻 Gestão & Computador (Assertec & Planilhas)",
        "📱 Lançamento Único (Telemóvel)",
        "📅 Agendamentos"
    ])

# ==============================================================================
# ABA 1: GESTÃO & COMPUTADOR
# ==============================================================================
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
                st.session_state["formulario_escolhido"] = "Brasil TecPar / Amigo (Apenas Planilha)"
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
            raw_servico = re.sub(r"ATATIVAÇÃO", "ATIVAÇÃO", raw_servico, flags=re.IGNORECASE)
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
            "Cole a mensagem do WhatsApp aqui (Ex: ASSERTEC \\n MARCIO ...)...",
            height=240,
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
                            pyautogui.press("tab")
                            time.sleep(0.04)
                        pyautogui.press("tab")
                        pyautogui.press("tab")
                        pyautogui.press("tab")
                        pyautogui.press("enter")
                        time.sleep(0.2)
                        pyautogui.press("enter")
                        time.sleep(0.08)

                        pyautogui.press("tab")
                        time.sleep(0.04)
                        pyautogui.press("enter")
                        time.sleep(0.25)
                        pyautogui.write(tecnico_val)
                        time.sleep(0.1)
                        pyautogui.press("enter")
                        time.sleep(0.1)

                        pyautogui.press("tab")
                        pyautogui.press("tab")
                        pyautogui.press("space")
                        time.sleep(0.1)

                        for _ in range(3):
                            pyautogui.press("tab")
                            time.sleep(0.04)

                        pyperclip.copy(str(cliente_val))
                        pyautogui.hotkey("ctrl", "v")
                        time.sleep(0.08)

                        pyautogui.press("tab")
                        pyperclip.copy(str(protocolo_val))
                        pyautogui.hotkey("ctrl", "v")
                        time.sleep(0.08)

                        pyautogui.press("tab")
                        for _ in range(cidade_idx):
                            pyautogui.press("tab")
                            time.sleep(0.04)

                        pyautogui.press("space")
                        time.sleep(0.1)

                        passos_restantes = len(lista_cidades_assertec) - cidade_idx
                        for _ in range(passos_restantes):
                            pyautogui.press("tab")
                            time.sleep(0.04)

                        pyautogui.press("enter")
                        time.sleep(0.3)

                        pyautogui.press("home")
                        time.sleep(0.1)
                        for _ in range(servico_idx):
                            pyautogui.press("down")
                            time.sleep(0.03)

                        pyautogui.press("enter")
                        time.sleep(0.2)

                        pyautogui.press("tab")
                        time.sleep(0.1)
                        pyautogui.press("tab")
                        time.sleep(0.1)

                        pyperclip.copy(str(quantidade_val))
                        pyautogui.hotkey("ctrl", "v")
                        time.sleep(0.1)

                        pyautogui.press("tab")
                        time.sleep(0.1)

                        if observacao_val.strip():
                            pyperclip.copy(str(observacao_val))
                            pyautogui.hotkey("ctrl", "v")
                            time.sleep(0.1)

                        st.success("Automação Assertec concluída com sucesso! 🚀")

            st.markdown("")
            if st.button("📊 Lançar na Planilha Google do Técnico", key="btn_lancar_plan_aba1"):
                tec_msg_atual = st.session_state.get("tec_msg", "") or st.session_state.get("tec", "ANDERSON")
                qtd_num = int(re.sub(r"\D", "", str(st.session_state.get("qtd", "1"))) or 1)
                dados_manual = {
                    "data": datetime.now().strftime("%d/%m/%Y"),
                    "tecnico_exec": tec_msg_atual.upper(),
                    "cliente": st.session_state.get("cli", "").upper(),
                    "protocolo": limpar_protocolo_str(st.session_state.get("proto", "")),
                    "cidade": st.session_state.get("cidade", ""),
                    "obs": st.session_state.get("obs", ""),
                    "servicos": [{"nome": st.session_state.get("servico", ""), "cabo": qtd_num, "eh_principal": True}]
                }
                registrar_servico_na_planilha_tecnico(dados_manual)

        st.markdown("---")
        st.subheader("🔍 Consulta Global de Cliente / Protocolo (Assertec)")
        termo_pesquisa_global = st.text_input(
            "Consultar por Nome do Cliente ou Protocolo:",
            value=st.session_state.get("cli", ""),
            key="pesquisa_global_topo"
        )

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

# ==============================================================================
# ABA 2: LANÇAMENTO ÚNICO UNIFICADO (TELEMÓVEL)
# ==============================================================================
with aba_gerador:
    st.subheader("📱 Lançamento Único de Atendimento")
    st.caption("Legenda:  ⚡ = Amigo + Assertec   |   🛠️ = Só Assertec")

    with st.form("form_lancamento_unico_telemovel", clear_on_submit=False):
        col_t1, col_t2 = st.columns(2)
        with col_t1:
            tecnico_selecionado = st.selectbox("Técnico que executou *", ["Escolher...", "ANDERSON", "MARCIO", "BRUNO"], key="uni_tecnico")
        with col_t2:
            mercado_selecionado = st.selectbox("Mercado (Brasil TecPar) *", ["AMIGO - VAREJO", "AVATO - CORPORATIVO"], index=0, key="uni_mercado")

        col_c1, col_c2 = st.columns(2)
        with col_c1:
            cliente_gen = st.text_input("Nome do Cliente *", placeholder="EX: ELISETE ANTONIA DOS SANTOS", key="uni_cliente")
        with col_c2:
            protocolo_gen = st.text_input("Número do Protocolo *", placeholder="EX: 951753258", key="uni_protocolo")

        cidade_gen = st.selectbox("Cidade *", ["Escolher..."] + lista_cidades_assertec, key="uni_cidade")

        st.markdown("---")
        st.markdown("**1️⃣ Serviço Principal (Inclui Instalações e Serviços Avulsos - LPU fixa em 1)**")
        servico_principal_curto = st.selectbox("Serviço Realizado *", LISTA_UNIFICADA_TELEMOVEL, key="uni_servico_principal")
        metragem_gen = st.number_input("Metragem / Cabo Utilizado (m) *", min_value=1, value=1, step=1, key="uni_cabo")

        st.markdown("---")
        st.markdown("**2️⃣ Serviços Extras no Mesmo Protocolo (Apenas CTO e Serviços Agregados)**")

        col_b1, col_b2 = st.columns([2, 1])
        with col_b1:
            batimento_curto = st.selectbox(
                "Houve Batimento de CTO neste protocolo?",
                list(MAPA_BATIMENTO_CURTO.keys()),
                key="uni_batimento"
            )
        with col_b2:
            lpu_batimento = st.number_input("Qtd CTOs Batidas (LPU)", min_value=1, value=1, step=1, key="uni_lpu_bat")

        col_a1, col_a2 = st.columns([2, 1])
        with col_a1:
            agregado1_curto = st.selectbox("Serviço Agregado 1 (Opcional)", LISTA_AGREGADOS_TELEMOVEL, key="uni_agregado_1")
        with col_a2:
            lpu_agregado1 = st.number_input("Qtd Agregado 1 (LPU)", min_value=1, value=1, step=1, key="uni_lpu_agr_1")

        col_e1, col_e2 = st.columns([2, 1])
        with col_e1:
            agregado2_curto = st.selectbox("Serviço Agregado 2 (Opcional)", LISTA_AGREGADOS_TELEMOVEL, key="uni_agregado_2")
        with col_e2:
            lpu_agregado2 = st.number_input("Qtd Agregado 2 (LPU)", min_value=1, value=1, step=1, key="uni_lpu_agr_2")

        col_f1, col_f2 = st.columns([2, 1])
        with col_f1:
            agregado3_curto = st.selectbox("Serviço Agregado 3 (Opcional)", LISTA_AGREGADOS_TELEMOVEL, key="uni_agregado_3")
        with col_f2:
            lpu_agregado3 = st.number_input("Qtd Agregado 3 (LPU)", min_value=1, value=1, step=1, key="uni_lpu_agr_3")

        st.markdown("---")
        with st.expander("📦 Equipamentos Utilizados (Obrigatório apenas para serviços de ATIVAÇÃO)", expanded=True):
            st.caption("Marque apenas os equipamentos utilizados. Os restantes serão lançados automaticamente como 'Não se aplica'.")
            equipamentos_form = {}
            for eq_nome in EQUIPAMENTOS_LISTA:
                c_eq1, c_eq2, c_eq3 = st.columns([1.5, 1, 1.5])
                padrao_usou = True if eq_nome == "ONU" else False
                with c_eq1:
                    usou_eq = st.checkbox(f"{eq_nome}", value=padrao_usou, key=f"eq_use_{eq_nome}")
                with c_eq2:
                    qtd_eq = st.number_input(f"Qtd ({eq_nome})", min_value=1, value=1, step=1, key=f"eq_qtd_{eq_nome}", label_visibility="collapsed")
                with c_eq3:
                    tipo_eq = st.selectbox(f"Tipo ({eq_nome})", ["Novo", "Remanufaturado"], key=f"eq_tipo_{eq_nome}", label_visibility="collapsed")
                equipamentos_form[eq_nome] = {"usou": usou_eq, "qtd": qtd_eq, "tipo": tipo_eq}

        obs_gen = st.text_area("Observações Técnicas (Opcional)", placeholder="EX: Refeita configuração nos equipamentos...", key="uni_obs")

        clicou_lancar_unico = st.form_submit_button(
            "🚀 Lançar no Sistema e Gerar Mensagem WhatsApp",
            type="primary",
            use_container_width=True
        )

    if clicou_lancar_unico:
        t_ok = tecnico_selecionado != "Escolher..."
        c_ok = cidade_gen != "Escolher..."
        s_ok = servico_principal_curto != "Escolher..."
        cli_ok = cliente_gen.strip() != ""
        p_ok = protocolo_gen.strip() != ""

        if not (t_ok and c_ok and s_ok and cli_ok and p_ok):
            st.error("⚠️ Preencha todos os campos obrigatórios (*): Técnico, Cliente, Protocolo, Cidade e Serviço Principal.")
        else:
            servico_prin_oficial = obter_nome_oficial_servico(servico_principal_curto)
            batimento_oficial = obter_nome_oficial_servico(batimento_curto)
            agregado1_oficial = obter_nome_oficial_servico(agregado1_curto)
            agregado2_oficial = obter_nome_oficial_servico(agregado2_curto)
            agregado3_oficial = obter_nome_oficial_servico(agregado3_curto)

            vai_amigo = (
                servico_vai_para_amigo(servico_prin_oficial)
                or batimento_oficial != "Não"
                or agregado1_oficial != "Não"
                or agregado2_oficial != "Não"
                or agregado3_oficial != "Não"
            )
            metragem_valida = max(1, int(metragem_gen or 1))

            lista_servicos_os = [
                {"nome": servico_prin_oficial, "lpu": 1, "cabo": metragem_valida, "eh_principal": True}
            ]
            if batimento_oficial != "Não":
                lista_servicos_os.append({"nome": batimento_oficial, "lpu": max(1, int(lpu_batimento or 1)), "cabo": 1, "eh_principal": False})
            if agregado1_oficial != "Não":
                lista_servicos_os.append({"nome": agregado1_oficial, "lpu": max(1, int(lpu_agregado1 or 1)), "cabo": 1, "eh_principal": False})
            if agregado2_oficial != "Não":
                lista_servicos_os.append({"nome": agregado2_oficial, "lpu": max(1, int(lpu_agregado2 or 1)), "cabo": 1, "eh_principal": False})
            if agregado3_oficial != "Não":
                lista_servicos_os.append({"nome": agregado3_oficial, "lpu": max(1, int(lpu_agregado3 or 1)), "cabo": 1, "eh_principal": False})

            dados_os_atual = {
                "data": datetime.now().strftime("%d/%m/%Y"),
                "tecnico_exec": tecnico_selecionado,
                "tecnico_portal": "ANDERSON WOLFART",
                "mercado": mercado_selecionado,
                "cliente": cliente_gen.strip().upper(),
                "protocolo": limpar_protocolo_str(protocolo_gen),
                "cidade": cidade_gen,
                "servicos": lista_servicos_os,
                "equipamentos": equipamentos_form,
                "obs": obs_gen.strip()
            }

            salvar_na_fila_assertec(dados_os_atual)

            # 1º PASSO: Grava IMEDIATAMENTE na planilha do técnico e mostra a linha exata
            with st.spinner("📊 Gravando serviço na planilha do técnico..."):
                registrar_servico_na_planilha_tecnico(dados_os_atual)

            # 2º PASSO: Gera e exibe na hora a mensagem do WhatsApp
            if tecnico_selecionado == "ANDERSON":
                nome_tec_assertec = "Anderson"
            elif tecnico_selecionado == "BRUNO":
                nome_tec_assertec = "Bruno"
            else:
                nome_tec_assertec = "Márcio"

            blocos_mensagens = []
            for srv_item in lista_servicos_os:
                nome_srv_assertec = mapear_para_assertec(srv_item["nome"], mercado_selecionado)
                metragem_msg = srv_item["cabo"] if srv_item.get("eh_principal") else srv_item["lpu"]
                bloco = f"""ASSERTEC
{tecnico_selecionado}
Técnico: {nome_tec_assertec}
Cliente: {cliente_gen.strip().upper()}
Protocolo: {limpar_protocolo_str(protocolo_gen)}
Cidade: {cidade_gen}
Serviço: {nome_srv_assertec}
Metragem: {metragem_msg}
Obs: {obs_gen.strip()}"""
                blocos_mensagens.append(bloco)

            st.session_state["mensagem_gerada_final"] = "\n\n\n".join(blocos_mensagens)

            st.markdown("### 📋 Mensagem Gerada:")
            st.code(st.session_state["mensagem_gerada_final"], language="text")
            mensagem_url = urllib.parse.quote(st.session_state["mensagem_gerada_final"])
            link_wpp_ag = f"https://wa.me/?text={mensagem_url}"
            st.link_button("📲 Enviar diretamente para o WhatsApp", link_wpp_ag, use_container_width=True)

            # 3º PASSO: Avisa que o protocolo já foi enviado e pode fechar ou aguardar
            st.info("✅ Protocolo enviado! Já pode fechar a página, mas se preferir pode aguardar o relatório do lançamento abaixo.")

            # 4º PASSO: Executa o lançamento em 2º plano blindado no Portal Brasil TecPar (AMIGO)
            if vai_amigo:
                with st.spinner("🤖 Lançando em segundo plano no portal Brasil TecPar (AMIGO)..."):
                    ok_tecpar, msg_tecpar = executar_lancamento_amigo_seguro(dados_os_atual)
                if ok_tecpar:
                    st.success(f"🟢 SERVIÇO CONFIRMADO NO SISTEMA AMIGO! ({msg_tecpar})")
                else:
                    st.warning(msg_tecpar)
                if os.path.exists(ARQUIVO_PRINT_PORTAL):
                    with st.expander("📸 Ver Print da Tela do Portal Brasil TecPar (Conferência Visual)", expanded=not ok_tecpar):
                        st.image(ARQUIVO_PRINT_PORTAL, caption="Tela final do portal Brasil TecPar após a execução do robô")
            else:
                st.success("🟢 SERVIÇO EXCLUSIVO ASSERTEC CONCLUÍDO!")

# ==============================================================================
# ABA 3: CENTRAL DE AGENDAMENTOS (PC & SMARTPHONE)
# ==============================================================================
with aba_agendamentos:
    st.subheader("☁️ Agenda Salesforce (Sincronização & Mensagem Pronta)")
    col_sf_btn, col_sf_mfa = st.columns([2, 1.5])
    with col_sf_mfa:
        codigo_mfa_input = st.text_input(
            "Código Google Authenticator (6 dígitos - só se pedir)",
            max_chars=6,
            placeholder="Ex: 123456",
            key="sf_mfa_code"
        )
    with col_sf_btn:
        st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
        clicou_sync_sf = st.button("🔄 SINCRONIZAR SALESFORCE", type="primary", use_container_width=True, key="btn_sync_sf")

    if clicou_sync_sf:
        with st.spinner("🤖 Acessando Salesforce Field Service e buscando agenda..."):
            ok_sf, msg_sf, lista_sf_nova = sincronizar_agenda_salesforce(codigo_mfa_input)
        if ok_sf:
            st.success(msg_sf)
        else:
            st.warning(msg_sf)

    if os.path.exists(ARQUIVO_PRINT_SF):
        with st.expander("📸 Ver Print da Tela do Salesforce (Diagnóstico Visual)", expanded=False):
            st.image(ARQUIVO_PRINT_SF, caption="Tela capturada pelo robô no Salesforce")

    agenda_sf_atual = carregar_agenda_sf_cache()

    if agenda_sf_atual:
        st.markdown("**📋 Clientes na Agenda do Salesforce (Clique para selecionar):**")
        opcoes_sf = [
            f"{idx+1}. ⏳ Vence: {item.get('vencimento')} | 👤 {item.get('cliente')} | 📞 {item.get('telefone') or 'S/Tel'} | 🏙️ {item.get('cidade')}"
            for idx, item in enumerate(agenda_sf_atual)
        ]
        escolha_sf_idx = st.selectbox(
            "Selecione um cliente da lista sincronizada:",
            range(len(opcoes_sf)),
            format_func=lambda i: opcoes_sf[i],
            key="sel_cliente_sf"
        )

        cliente_sf_sel = agenda_sf_atual[escolha_sf_idx]

        with st.form("form_gerar_msg_salesforce", clear_on_submit=False):
            st.markdown(f"**✏️ Preparar Mensagem para: `{cliente_sf_sel.get('cliente')}`**")
            st.caption(f"Dados brutos capturados: {cliente_sf_sel.get('resumo_bruto', '')}")

            campo_livre_sf = st.text_area(
                "1. Mensagem / Observação da Agendadora *",
                placeholder="Ex: Turno da Manhã - Ligar 15 min antes de ir...",
                key="sf_campo_livre"
            )

            c_sf1, c_sf2 = st.columns(2)
            with c_sf1:
                tipo_proto_sf = st.text_input("2. Tipo de Protocolo", value=cliente_sf_sel.get("tipo_protocolo", ""), key="sf_tipo_p")
                nome_cli_sf = st.text_input("3. Nome do Cliente", value=cliente_sf_sel.get("cliente", ""), key="sf_nome_c")
                tel_cli_sf = st.text_input("4. Telefone do Cliente", value=cliente_sf_sel.get("telefone", ""), key="sf_tel_c")
            with c_sf2:
                num_proto_sf = st.text_input("5. Número do Protocolo", value=cliente_sf_sel.get("protocolo", ""), key="sf_num_p")
                etiqueta_sf = st.text_input("6. Etiqueta do Cliente", value=cliente_sf_sel.get("etiqueta", ""), key="sf_etiq_c")
                cidade_sf = st.text_input("7. Cidade", value=cliente_sf_sel.get("cidade", ""), key="sf_cid_c")

            endereco_sf = st.text_input(
                "8. Endereço do Cliente (Será convertido em Link do Google Maps)",
                value=cliente_sf_sel.get("endereco", ""),
                placeholder="Ex: Rua Independência, 120, Ivoti - RS",
                key="sf_end_c"
            )

            btn_gerar_msg_sf = st.form_submit_button("📲 Gerar Mensagem Pronta", type="primary", use_container_width=True)

        if btn_gerar_msg_sf:
            endereco_query = f"{endereco_sf.strip()} {cidade_sf.strip()} RS".strip()
            link_maps = (
                f"https://www.google.com/maps/search/?api=1&query={urllib.parse.quote(endereco_query)}"
                if endereco_sf.strip() else "Não informado"
            )

            msg_sf_pronta = f"""{campo_livre_sf.strip()}
📌 Tipo de Protocolo: {tipo_proto_sf.strip()}
👤 Nome do cliente: {nome_cli_sf.strip()}
📞 Telefone do cliente: {tel_cli_sf.strip()}
🔢 Número do Protocolo: {num_proto_sf.strip()}
🏷️ Etiqueta do cliente: {etiqueta_sf.strip()}
🏙️ Cidade: {cidade_sf.strip()}
📍 Endereço: {link_maps}"""

            st.session_state["msg_sf_final"] = msg_sf_pronta

        if st.session_state.get("msg_sf_final"):
            st.markdown("### 📋 Mensagem Pronta para o Grupo do Técnico:")
            st.code(st.session_state["msg_sf_final"], language="text")
            url_wpp_sf = f"https://wa.me/?text={urllib.parse.quote(st.session_state['msg_sf_final'])}"
            st.link_button("📲 Enviar no WhatsApp", url_wpp_sf, use_container_width=True)

    st.markdown("---")
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
                obs_atual = str(lista_mem[idx_linha].get("OBS", ""))
                if "CONCLUÍDO" not in obs_atual.upper():
                    lista_mem[idx_linha]["OBS"] = "✅ [CONCLUÍDO] " + obs_atual
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
            if "CONCLUÍDO" in obs_str.upper() or "CONCLUIDO" in obs_str.upper():
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
                proto_item = limpar_protocolo_str(row.get("PROTOCOLO", ""))
                cidade_item = str(row.get("CIDADE", "N/D"))
                cliente_item = str(row.get("CLIENTE", "N/D"))
                obs_item = str(row.get("OBS", "-"))
                data_item = str(row.get("DATA_REAL", "N/D"))
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

    def concluido_expirado_10_dias(row, data_hoje):
        obs_str = str(row.get("OBS", "")).upper()
        if "CONCLUÍDO" not in obs_str and "CONCLUIDO" not in obs_str:
            return False
        data_str = extrair_data_da_linha(row)
        try:
            data_dt = datetime.strptime(data_str, "%d/%m/%Y").date()
            if (data_hoje - data_dt).days > 10:
                return True
        except Exception:
            pass
        return False

    with st.expander("📊 Ver Tabela Completa da Planilha (Incluindo Concluídos)"):
        hoje_dt = datetime.now().date()
        lista_tabela = [
            r for r in lista_agend
            if not concluido_expirado_10_dias(r, hoje_dt)
        ]
        if lista_tabela:
            linhas_html = ""
            for r in lista_tabela:
                obs_texto = str(r.get("OBS", ""))
                data_exibicao = extrair_data_da_linha(r)
                eh_concluido = "CONCLUÍDO" in obs_texto.upper() or "CONCLUIDO" in obs_texto.upper()
                estilo_celula = (
                    "padding:6px; border-bottom:1px solid #c3e6cb; background-color:#d4edda !important; color:#155724 !important; font-weight:600;"
                    if eh_concluido else
                    "padding:6px; border-bottom:1px solid #ddd; background-color:#ffffff; color:#262730;"
                )
                linhas_html += (
                    f"<tr>"
                    f"<td style='{estilo_celula}'>{html.escape(str(data_exibicao))}</td>"
                    f"<td style='{estilo_celula}'>{html.escape(str(r.get('CLIENTE', '')))}</td>"
                    f"<td style='{estilo_celula}'>{html.escape(str(r.get('CIDADE', '')))}</td>"
                    f"<td style='{estilo_celula}'>{html.escape(obs_texto)}</td>"
                    f"</tr>"
                )
            tabela_html = f"""
            <div style="overflow-x:auto;">
                <table style="width:100%; border-collapse:collapse; font-size:0.85rem;">
                    <thead>
                        <tr style="background-color:#f0f2f6; color:#262730; text-align:left;">
                            <th style="padding:6px; border-bottom:2px solid #ccc;">DATA</th>
                            <th style="padding:6px; border-bottom:2px solid #ccc;">CLIENTE</th>
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

    if st.session_state.get("msg_agendamento_gerada"):
        st.markdown("### 📋 Mensagem Gerada:")
        st.code(st.session_state["msg_agendamento_gerada"], language="text")

        msg_ag_url = urllib.parse.quote(st.session_state["msg_agendamento_gerada"])
        link_wpp_ag = f"https://wa.me/?text={msg_ag_url}"
        st.link_button("📲 Enviar para o Grupo do WhatsApp", link_wpp_ag, use_container_width=True)