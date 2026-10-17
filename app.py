import streamlit as st
import pyautogui
import pyperclip
import time
import webbrowser
import pandas as pd

# Configuração da página
st.set_page_config(page_title="Sistema de Automação", layout="centered")

st.title("🛠️ Sistema de Automação de Processos")

# Menu de seleção de formulários
formulario_escolhido = st.sidebar.selectbox(
    "Selecione o Formulário",
    ["Escolher...", "Assertec / Fibrawer", "Brasil TecPar / Amigo"]
)

if formulario_escolhido == "Escolher...":
    st.info("👈 Selecione um formulário no menu lateral para começar.")

# --- FORMULÁRIO 1: ASSERTEC / FIBRAWER ---
elif formulario_escolhido == "Assertec / Fibrawer":
    st.subheader("📋 Preenchimento - Assertec / Fibrawer")
    
    # [Mantém aqui todo o código que já tinhas para a Assertec]
    cliente = st.text_input("Nome do cliente *", value="")
    # ...

# --- FORMULÁRIO 2: BRASIL TECPAR / AMIGO ---
elif formulario_escolhido == "Brasil TecPar / Amigo":
    st.subheader("📋 Preenchimento - Brasil TecPar / Amigo")

    lista_tecnicos_amigo = [
        "Escolher...", "MICHAEL DOUGLAS PANZENHAGEN", "JONAS SCHUTZ", "CASSIANO PALAVRO",
        "MAURICIO JONATAN DA CRUZ", "ANDERSON WOLFART MOLLERKE", "SAMUEL SCHNEIDER SCHOMMER", "HERIVELTON SILVEIRA NERY"
    ]
    tec_importado = st.session_state.get("tec", "")
    tec_idx_amigo = 0
    for idx, t in enumerate(lista_tecnicos_amigo):
        if t.lower() in tec_importado.lower() and tec_importado.strip():
            tec_idx_amigo = idx
            break
    tecnico_amigo = st.selectbox("Nome do Técnico / Assistente *", lista_tecnicos_amigo, index=tec_idx_amigo)

    lista_cidades_amigo = [
        "Escolher...", "ALVORADA", "CAXIAS DO SUL", "CAMPO BOM", "DOIS IRMÃOS", "ESTÂNCIA VELHA",
        "IVOTI", "NOVO HAMBURGO", "PAROBÉ", "PORTÃO", "SÃO LEOPOLDO", "SÃO SEBASTIÃO DO CAÍ", "TAQUARA"
    ]
    cidade_importada = st.session_state.get("cidade", "")
    cidade_idx_amigo = 0
    for idx, c in enumerate(lista_cidades_amigo):
        if c.lower() in cidade_importada.lower() and cidade_importada.strip():
            cidade_idx_amigo = idx
            break
    cidade_amigo = st.selectbox("Cidade *", lista_cidades_amigo, index=cidade_idx_amigo)

    lista_servicos_amigo = [
        "Escolher...",
        "ATIVAÇÃO DE SERVIÇO - ACESSO À INTERNET (0 ~ 300 MT)",
        "ATIVAÇÃO DE SERVIÇO - ACESSO À INTERNET (301 ~ 500MT)",
        "ATIVAÇÃO DE SERVIÇO - ACESSO À INTERNET (ACIMA DE 500MT)",
        "BATIMENTO DE CAIXA 1X16 PORTAS",
        "BATIMENTO DE CAIXA 1X8 PORTAS",
        "DESATIVAÇÃO DE SERVIÇO - PARCIAL",
        "DIÁRIAS (SOB DEMANDA - DUPLA)",
        "LANÇAMENTO DE FIBRA AVULSA ATÉ - FIBRA ATÉ 12V",
        "TROCA DE PONTO INTERNO (PONTO PRINCIPAL/WIFI/TV/FONE/CÂMERAS/PONTO EXTRA)",
        "VISTORIA TÉCNICA / VIABILIDADE TÉCNICA"
    ]
    servico_importado = st.session_state.get("servico", "")
    servico_idx_amigo = 0
    for idx, s in enumerate(lista_servicos_amigo):
        if s.lower() in servico_importado.lower() and servico_importado.strip():
            servico_idx_amigo = idx
            break
    servico_amigo = st.selectbox("Serviço Realizado *", lista_servicos_amigo, index=servico_idx_amigo)

    cliente = st.text_input("Nome do cliente *", value=st.session_state.get("cli", ""))
    protocolo = st.text_input("Protocolo *", value=st.session_state.get("proto", ""))
    quantidade_cabo = st.text_input("Quantidade de cabo utilizado (Metragem) *", value=st.session_state.get("qtd", "1"))
    observacao = st.text_area("Observações Técnicas", value=st.session_state.get("obs", ""))

    st.markdown("---")
    
    col_exec_amigo, col_conf_amigo = st.columns(2)

    with col_exec_amigo:
        btn_executar_amigo = st.button("🚀 Executar Automação Brasil TecPar", key="btn_exec_amigo")

    with col_conf_amigo:
        btn_conferir_amigo = st.button("🔍 Conferir na Planilha Amigo", key="btn_conf_amigo")

    if btn_executar_amigo:
        if tecnico_amigo == "Escolher..." or cidade_amigo == "Escolher..." or servico_amigo == "Escolher..." or not cliente.strip() or not protocolo.strip():
            st.error("Preencha todos os campos obrigatórios (*).")
        else:
            st.warning("⚡ A executar automação Brasil TecPar... Não mexa no rato nem no teclado.")
            
            webbrowser.open("https://docs.google.com/forms/d/e/1FAIpQLSfDT9GbrgCv5vcOi2fvZoeuPITItTWIQFWmS2vO_NY7gMEvQA/viewform")
            time.sleep(2.5)
            
            # --- PÁGINA 1 ---
            for _ in range(3):
                pyautogui.press('tab')
                time.sleep(0.04)
            pyautogui.press('enter')
            time.sleep(0.3)
            pyautogui.write("ASSERTEC SERVICOS E TECNOLOGIA LTDA")
            time.sleep(0.15)
            pyautogui.press('enter')
            time.sleep(0.15)
            
            pyautogui.press('tab')
            time.sleep(0.05)
            pyautogui.press('enter')
            time.sleep(0.3)
            pyautogui.write(tecnico_amigo)
            time.sleep(0.15)
            pyautogui.press('enter')
            time.sleep(0.15)
            
            pyautogui.press('tab')
            time.sleep(0.05)
            pyautogui.press('enter')
            time.sleep(1.2)
            
            # --- PÁGINA 2 ---
            for _ in range(3):
                pyautogui.press('tab')
                time.sleep(0.04)
            pyperclip.copy(str(protocolo))
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.15)
            
            pyautogui.press('tab')
            time.sleep(0.05)
            pyautogui.press('enter')
            time.sleep(0.3)
            pyautogui.write("AMIGO - VAREJO")
            time.sleep(0.15)
            pyautogui.press('enter')
            time.sleep(0.15)
            
            pyautogui.press('tab')
            time.sleep(0.05)
            pyperclip.copy(str(cliente))
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.15)
            
            pyautogui.press('tab')
            time.sleep(0.05)
            pyautogui.press('enter')
            time.sleep(0.3)
            pyautogui.write(cidade_amigo)
            time.sleep(0.15)
            pyautogui.press('enter')
            time.sleep(0.15)
            
            pyautogui.press('tab')
            time.sleep(0.1)
            pyautogui.press('tab')
            time.sleep(0.1)
            pyperclip.copy("0")
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.15)
            
            pyautogui.press('tab')
            time.sleep(0.05)
            pyautogui.press('tab')
            time.sleep(0.05)
            pyautogui.press('enter')
            time.sleep(1.2)
            
            # --- PÁGINA 3 ---
            for _ in range(3):
                pyautogui.press('tab')
                time.sleep(0.04)
            pyautogui.press('enter')
            time.sleep(0.3)
            pyautogui.write(servico_amigo)
            time.sleep(0.15)
            pyautogui.press('enter')
            time.sleep(0.15)
            
            pyautogui.press('tab')
            time.sleep(0.05)
            pyautogui.press('tab')
            time.sleep(0.05)
            pyautogui.press('enter')
            time.sleep(1.2)
            
            # --- PÁGINA 4 ---
            for _ in range(3):
                pyautogui.press('tab')
                time.sleep(0.04)
            pyperclip.copy(str(quantidade_cabo))
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.15)
            
            pyautogui.press('tab')
            time.sleep(0.05)
            pyautogui.press('tab')
            time.sleep(0.05)
            pyautogui.press('enter')
            time.sleep(1.2)
            
            # --- PÁGINA 5 ---
            for _ in range(3):
                pyautogui.press('tab')
                time.sleep(0.04)
            pyperclip.copy("1")
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.15)
            
            pyautogui.press('tab')
            time.sleep(0.05)
            pyautogui.press('tab')
            time.sleep(0.05)
            pyautogui.press('enter')
            time.sleep(1.2)
            
            # --- PÁGINA 6 ---
            for _ in range(3):
                pyautogui.press('tab')
                time.sleep(0.04)
            if observacao.strip():
                pyperclip.copy(str(observacao))
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(0.15)
            
            st.success("Automação Brasil TecPar concluída até às Observações! Podes conferir e enviar manualmente. 🚀")

    if btn_conferir_amigo:
        if not cliente.strip():
            st.warning("Insira o nome do cliente no campo 'Nome do cliente' acima para realizar a conferência.")
        else:
            with st.spinner("A abrir a planilha e a localizar o cliente no navegador..."):
                webbrowser.open("https://docs.google.com/spreadsheets/d/1jESKPwkNUo0ocIPOUCzKKeq8wav57Rdd_K9MF3WBSDk/edit?gid=1573921361#gid=1573921361")
                time.sleep(3.5)
                
                pyautogui.hotkey('ctrl', 'f')
                time.sleep(0.4)
                pyperclip.copy(str(cliente))
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(0.3)
                pyautogui.press('enter')
                
            st.success(f"✅ Planilha aberta com sucesso! O navegador procurou automaticamente por **'{cliente}'**.")
            st.info("💡 **Legenda de Cores na Planilha:** 🟢 **Verde** = Linha Aprovada | 🔴 **Vermelho** = Linha Negada.")