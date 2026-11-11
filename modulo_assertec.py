import os
import re
import json
from config import ARQUIVO_FILA_ASSERTEC_LOCAL

JUSTIFICATIVA_PADRAO_TECPAR = "Solicitado pelo cliente"
MSG_OBS_PADRAO_AMIGO = "Protocolo encerrado automaticamente pelo sistema autônomo FibraWer, duvidas contactar Anderson Wolfart (51) 999260016."

lista_cidades_assertec = [
    "Alto Feliz", "Bom Principio", "Caxias do Sul", "Feliz", "Harmonia", "Ivoti",
    "Linha Nova", "Pareci Novo", "Presidente Lucena", "São José do Hortêncio",
    "São Sebastião do Caí", "São Vendelino", "Tupandi", "Vale Real"
]

SERVICOS_PORTAL_TECPAR = [
    "ATIVAÇÃO DE SERVIÇO - ACESSO À INTERNET (0 ~ 300 MT)",
    "ATIVAÇÃO DE SERVIÇO - ACESSO À INTERNET (301 ~ 500MT)",
    "ATIVAÇÃO DE SERVIÇO - ACESSO À INTERNET (ACIMA DE 500MT)",
    "ATIVAÇÃO DE SERVIÇO - ACESSO À INTERNET (RÁDIO)",
    "ATIVAÇÃO DE SERVIÇO AGREGADO (CÂMERAS WIFI)",
    "ATIVAÇÃO DE SERVIÇO AGREGADO (CÂMERAS)",
    "ATIVAÇÃO DE SERVIÇO AGREGADO (WIFI/TV/FONE/PONTO EXTRA)",
    "ATIVAÇÃO DE SERVIÇO AGREGADO AVULSO (CÂMERAS WIFI)",
    "ATIVAÇÃO DE SERVIÇO AGREGADO AVULSO (WIFI/TV/FONE/PONTO EXTRA)",
    "BATIMENTO DE CAIXA 1X16 PORTAS",
    "BATIMENTO DE CAIXA 1X8 PORTAS",
    "DESATIVAÇÃO DE SERVIÇO",
    "DIÁRIAS (SOB DEMANDA - DUPLA)",
    "LANÇAMENTO DE FIBRA AVULSA ATÉ - FIBRA ATÉ 12V",
    "TROCA DE PONTO INTERNO (PONTO PRINCIPAL/WIFI/TV/FONE/CÂMERAS/PONTO EXTRA)",
    "VISTORIA TÉCNICA / VIABILIDADE TÉCNICA"
]

MAPA_SERVICOS_PRINCIPAIS_CURTOS = {
    # ⚡ Serviços AMIGO + ASSERTEC (Principais e Avulsos)
    "⚡ INSTALAÇÃO (0 A 300M)": "ATIVAÇÃO DE SERVIÇO - ACESSO À INTERNET (0 ~ 300 MT)",
    "⚡ INSTALAÇÃO (301 A 500M)": "ATIVAÇÃO DE SERVIÇO - ACESSO À INTERNET (301 ~ 500MT)",
    "⚡ INSTALAÇÃO (+ DE 500M)": "ATIVAÇÃO DE SERVIÇO - ACESSO À INTERNET (ACIMA DE 500MT)",
    "⚡ INSTALAÇÃO (VIA RÁDIO)": "ATIVAÇÃO DE SERVIÇO - ACESSO À INTERNET (RÁDIO)",
    "⚡ AVULSO PONTO EXTRA": "ATIVAÇÃO DE SERVIÇO AGREGADO AVULSO (WIFI/TV/FONE/PONTO EXTRA)",
    "⚡ AVULSO CÂMERA WIFI": "ATIVAÇÃO DE SERVIÇO AGREGADO AVULSO (CÂMERAS WIFI)",
    "⚡ BATIDA CTO 1X16": "BATIMENTO DE CAIXA 1X16 PORTAS",
    "⚡ BATIDA CTO 1X8": "BATIMENTO DE CAIXA 1X8 PORTAS",
    "⚡ DESATIVAÇÃO DE SERVIÇO": "DESATIVAÇÃO DE SERVIÇO",
    "⚡ DIÁRIA (DUPLA)": "DIÁRIAS (SOB DEMANDA - DUPLA)",
    "⚡ FIBRA AVULSA ATÉ 12V": "LANÇAMENTO DE FIBRA AVULSA ATÉ - FIBRA ATÉ 12V",
    "⚡ TROCA DE PONTO INTERNO": "TROCA DE PONTO INTERNO (PONTO PRINCIPAL/WIFI/TV/FONE/CÂMERAS/PONTO EXTRA)",
    "⚡ VISTORIA / VIABILIDADE": "VISTORIA TÉCNICA / VIABILIDADE TÉCNICA",

    # 🛠️ Serviços Exclusivos ASSERTEC
    "🛠️ MANUTENÇÃO / REPARO": "Atendimento Presencial - REPARO",
    "🛠️ ROMPIMENTO ATÉ 500M": "Atendimento Presencial - REPARO - ROMPIMENTO/LANÇAMENTO até 500 mt",
    "🛠️ ATENDIMENTO RNC": "Atendimento presencial - RNC",
    "🛠️ VISITA IMPRODUTIVA": "Visita improdutiva",
    "🛠️ DESATIVAÇÃO COMPLETA": "Desativações de serviços completa",
    "🛠️ TROCA DE ENDEREÇO": "Desativações de serviços parcial / troca de endereço / migração de tecnologia",
    "🛠️ KM ACIMA DE 50KM": "Km rodado acima de 50km",
    "🛠️ METRAGEM + DE 500M": "Metragem excedente (acima de 500 mt)",
    "🛠️ FUSÃO EM CAIXA": "Fusão em caixa",
    "🛠️ EMENDA CABO ROMPIDO": "Emenda de cabos rompidos na rede em infraestrutura própria",
    "🛠️ NOVA CTO (INSTALAÇÃO)": "Instalação de nova CTO",
    "🛠️ LANÇAMENTO CABO AS": "Lançamento de cabo AS",
    "🛠️ PLANTÃO - DUPLO": "PLANTAO - duplo",
    "🛠️ PLANTÃO - SIMPLES": "PLANTAO - simples",
    "🛠️ PLANTÃO - REPARO": "PLANTAO - Atendimento Presencial - REPARO",
    "🛠️ PLANTÃO - ROMPIMENTO": "PLANTAO - Atendimento Presencial - REPARO - ROMPIMENTO/LANÇAMENTO até 500 mt",
}

MAPA_BATIMENTO_CURTO = {
    "Não": "Não",
    "⚡ BATIDA CTO 1X16": "BATIMENTO DE CAIXA 1X16 PORTAS",
    "⚡ BATIDA CTO 1X8": "BATIMENTO DE CAIXA 1X8 PORTAS",
}

MAPA_AGREGADO_CURTO = {
    "Não": "Não",
    "⚡ AGREGADO PONTO EXTRA": "ATIVAÇÃO DE SERVIÇO AGREGADO (WIFI/TV/FONE/PONTO EXTRA)",
    "⚡ AGREGADO CÂMERA WIFI": "ATIVAÇÃO DE SERVIÇO AGREGADO (CÂMERAS WIFI)",
    "⚡ AGREGADO CÂMERAS CABO": "ATIVAÇÃO DE SERVIÇO AGREGADO (CÂMERAS)",
}

MAPA_SERVICOS_CURTOS = {
    **MAPA_SERVICOS_PRINCIPAIS_CURTOS,
    **{k: v for k, v in MAPA_AGREGADO_CURTO.items() if k != "Não"}
}

LISTA_UNIFICADA_TELEMOVEL = ["Escolher..."] + list(MAPA_SERVICOS_PRINCIPAIS_CURTOS.keys())
LISTA_AGREGADOS_TELEMOVEL = list(MAPA_AGREGADO_CURTO.keys())

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

EQUIPAMENTOS_LISTA = [
    "ONU",
    "Roteador",
    "Switch",
    "Câmera",
    "Setup box",
    "Fire Stick",
    "Aparelho de telefone",
    "ATA"
]

def limpar_protocolo_str(val):
    if val is None:
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s.strip()

def obter_nome_oficial_servico(rotulo_curto):
    if rotulo_curto in MAPA_SERVICOS_CURTOS:
        return MAPA_SERVICOS_CURTOS[rotulo_curto]
    if rotulo_curto in MAPA_BATIMENTO_CURTO:
        return MAPA_BATIMENTO_CURTO[rotulo_curto]
    if rotulo_curto in MAPA_AGREGADO_CURTO:
        return MAPA_AGREGADO_CURTO[rotulo_curto]
    texto = str(rotulo_curto or "").strip()
    texto = re.sub(r"^(⚡|🛠️)\s*(\[.*?\]\s*)?", "", texto)
    return texto.strip()

def servico_vai_para_amigo(servico_oficial):
    return str(servico_oficial or "").upper() in [s.upper() for s in SERVICOS_PORTAL_TECPAR]

def servico_exige_equipamentos(servico_oficial):
    return "ATIVAÇÃO" in str(servico_oficial or "").upper()

def mapear_para_assertec(servico_oficial, mercado="AMIGO - VAREJO"):
    s_up = str(servico_oficial or "").upper()
    eh_corp = "CORPORATIVO" in str(mercado or "").upper() or "AVATO" in str(mercado or "").upper()
    if eh_corp:
        for item in lista_servicos_assertec_global:
            if item.upper() == f"CORPORATIVO - {s_up}":
                return item
    mapa_direto = {
        "ATIVAÇÃO DE SERVIÇO - ACESSO À INTERNET (0 ~ 300 MT)": "Ativação de Serviço - Acesso à Internet (0 ~ 300 mt)",
        "ATIVAÇÃO DE SERVIÇO - ACESSO À INTERNET (301 ~ 500MT)": "Ativação de Serviço - Acesso à Internet (301 ~ 500 mt)",
        "ATIVAÇÃO DE SERVIÇO - ACESSO À INTERNET (ACIMA DE 500MT)": "Metragem excedente (acima de 500 mt)",
        "ATIVAÇÃO DE SERVIÇO - ACESSO À INTERNET (RÁDIO)": "Ativação de Serviço - Acesso à Internet (0 ~ 300 mt)",
        "ATIVAÇÃO DE SERVIÇO AGREGADO (CÂMERAS WIFI)": "Ativação de Serviço - Agregado (câmeras wifi)",
        "ATIVAÇÃO DE SERVIÇO AGREGADO (CÂMERAS)": "Ativação de Serviço - Agregado (câmeras)",
        "ATIVAÇÃO DE SERVIÇO AGREGADO (WIFI/TV/FONE/PONTO EXTRA)": "Ativação de Serviço - Agregado (wifi/tv/fone/ponto extra)",
        "ATIVAÇÃO DE SERVIÇO AGREGADO AVULSO (CÂMERAS WIFI)": "Ativação de Serviço - Avulso (câmeras wifi)",
        "ATIVAÇÃO DE SERVIÇO AGREGADO AVULSO (WIFI/TV/FONE/PONTO EXTRA)": "Ativação de Serviço - Agregado Avulso (wifi/tv/fone/ponto extra)",
        "BATIMENTO DE CAIXA 1X16 PORTAS": "Batimento de caixa 1x16 p",
        "BATIMENTO DE CAIXA 1X8 PORTAS": "Batimento de caixa 1x8p",
        "DESATIVAÇÃO DE SERVIÇO": "Desativações de serviços parcial / troca de endereço / migração de tecnologia",
        "DIÁRIAS (SOB DEMANDA - DUPLA)": "Diária de atendimentos (sob demanda - dupla)",
        "LANÇAMENTO DE FIBRA AVULSA ATÉ - FIBRA ATÉ 12V": "Lançamento de fibra avulsa até f.o 6v",
        "TROCA DE PONTO INTERNO (PONTO PRINCIPAL/WIFI/TV/FONE/CÂMERAS/PONTO EXTRA)": "Troca de ponto interno (ponto principal/wifi/tv/fone/câmeras/ponto extra)",
        "VISTORIA TÉCNICA / VIABILIDADE TÉCNICA": "Vistoria Técnica/ Viabilidade técnica",
    }
    return mapa_direto.get(s_up, servico_oficial)

def carregar_fila_assertec():
    if os.path.exists(ARQUIVO_FILA_ASSERTEC_LOCAL):
        try:
            with open(ARQUIVO_FILA_ASSERTEC_LOCAL, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []
    return []

def salvar_na_fila_assertec(dados_os):
    fila = carregar_fila_assertec()
    prot_novo = limpar_protocolo_str(dados_os.get("protocolo", ""))
    if prot_novo:
        fila = [item for item in fila if limpar_protocolo_str(item.get("protocolo", "")) != prot_novo]
    fila.append(dados_os)
    try:
        with open(ARQUIVO_FILA_ASSERTEC_LOCAL, "w", encoding="utf-8") as f:
            json.dump(fila, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"Erro ao salvar fila local: {e}")