import os
import streamlit as st

def obter_segredo(chave, padrao=""):
    try:
        return st.secrets[chave]
    except Exception:
        return os.environ.get(chave, padrao)

# Credenciais e URLs carregadas do cofre (.streamlit/secrets.toml)
SHEET_ID_AGENDAMENTOS = obter_segredo("SHEET_ID_AGENDAMENTOS")
URL_PLANILHA_AGENDAMENTOS_CSV = f"https://docs.google.com/spreadsheets/d/{SHEET_ID_AGENDAMENTOS}/gviz/tq?tqx=out:csv"
GOOGLE_SCRIPT_URL = obter_segredo("GOOGLE_SCRIPT_URL")
GOOGLE_SCRIPT_PLANILHAS_URL = obter_segredo("GOOGLE_SCRIPT_PLANILHAS_URL")

TECPAR_PORTAL_URL = obter_segredo("TECPAR_PORTAL_URL")
TECPAR_USER = obter_segredo("TECPAR_USER")
TECPAR_PASS = obter_segredo("TECPAR_PASS")

SALESFORCE_URL = obter_segredo("SALESFORCE_URL")
SALESFORCE_USER = obter_segredo("SALESFORCE_USER")
SALESFORCE_USER_ALT = obter_segredo("SALESFORCE_USER_ALT")
SALESFORCE_PASS = obter_segredo("SALESFORCE_PASS")

# Arquivos de controle local e sessões
ARQUIVO_AGENDAMENTOS_LOCAL = "agendamentos_backup.csv"
ARQUIVO_FILA_ASSERTEC_LOCAL = "fila_assertec_excel.json"
ARQUIVO_PRINT_PORTAL = "ultimo_lancamento_amigo.png"
COLUNAS_PADRAO_AG = ["DATA", "CLIENTE", "PROTOCOLO", "CIDADE", "OBS"]

PASTA_SESSAO_SF = os.path.join(os.getcwd(), "sessao_salesforce")
ARQUIVO_PRINT_SF = "ultimo_print_salesforce.png"
ARQUIVO_CACHE_SF = "agenda_salesforce_cache.json"
ASSERTEC_PLANILHA_URL = obter_segredo("https://docs.google.com/spreadsheets/d/1MMQ8PAjr2GQfWxPnfYNfHSiiu9AUdE0pvfJikSGQgqU/edit?usp=sharing", "")
ASSERTEC_FORM_URL = obter_segredo("ASSERTEC_FORM_URL", "https://forms.gle/2MENn9xuKjj7pKVm6")
SENHA_FINANCEIRO = obter_segredo("SENHA_FINANCEIRO", "")