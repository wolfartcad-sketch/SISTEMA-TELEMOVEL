import re
import requests
from datetime import datetime
import streamlit as st
from config import GOOGLE_SCRIPT_PLANILHAS_URL
from modulo_assertec import limpar_protocolo_str

# Cidades da base normal (sem sufixo FB). As demais recebem " FB" automaticamente.
CIDADES_BASE_NORMAL = {
    "ivoti",
    "são josé do hortêncio",
    "sao jose do hortencio",
    "linha nova",
    "são sebastião do caí",
    "sao sebastiao do cai",
    "presidente lucena",
}

def traduzir_servico_para_planilha(nome_servico, metragem_cabo=0, cidade=""):
    """
    Traduz o nome do serviço para um dos 33 nomes exatos da TABELA DE PREÇOS da planilha,
    garantindo que o =XLOOKUP(...) da coluna H encontre o valor correto.
    """
    s = str(nome_servico or "").strip().upper()
    cid_limpa = str(cidade or "").strip().lower()

    try:
        m = int(re.sub(r"\D", "", str(metragem_cabo)) or 0)
    except Exception:
        m = 0

    fora_da_base = bool(cid_limpa) and (cid_limpa not in CIDADES_BASE_NORMAL)
    eh_plantao_reparo = ("PLANTAO" in s or "PLANTÃO" in s) and ("REPARO" in s)
    ja_tem_fb = s.endswith(" FB")
    sufixo_fb = " FB" if (fora_da_base or eh_plantao_reparo or ja_tem_fb) else ""

    # 1. Retirada / Desativação completa (único item da tabela que nunca leva " FB")
    if "DESATIVAÇÕES DE SERVIÇOS COMPLETA" in s or "DESATIVAÇÃO DE SERVIÇO - COMPLETA" in s or s == "DESATIVAÇÃO DE SERVIÇO" or "RETIRADA" in s:
        return "RETIRADA"

    # 2. Movimentar (Troca de ponto interno)
    if "TROCA DE PONTO INTERNO" in s or "MOVIMENTAR" in s:
        base = "MOVIMENTAR"

    # 3. Batida de CTO 1/8 e 1/16
    elif "BATIMENTO" in s or "BATIDA" in s:
        base = "BATIDA DE CTO 1/16" if "16" in s else "BATIDA DE CTO 1/8"

    # 4. Recolhimento de equipamento no endereço antigo
    elif "RECOLHIMENTO" in s:
        base = "RECOLHIMENTO EQUIPAMENTO ENDEREÇO ANTIGO"

    # 5. Troca de endereço (normal e mais de 300m)
    elif "TROCA DE ENDEREÇO" in s or "PARCIAL" in s or "MIGRAÇÃO" in s:
        base = "TROCA DE ENDEREÇO MAIS DE 300M" if (m > 300 or "MAIS DE 300" in s) else "TROCA DE ENDEREÇO"

    # 6. Visita Improdutiva
    elif "IMPRODUTIVA" in s:
        base = "IMPRODUTIVA"

    # 7. Rompimento / Emenda / Fusão / Lançamento de fibra ou cabo AS
    elif "ROMPIMENTO" in s or "EMENDA" in s or "FUSÃO" in s or "FIBRA AVULSA" in s or "CABO AS" in s:
        base = "ROMPIMENTO"

    # 8. Manutenção / Reparo / RNC / Vistoria
    elif "REPARO" in s or "MANUTENÇÃO" in s or "MANUTENCAO" in s or "RNC" in s or "VISTORIA" in s:
        base = "MANUTENÇÃO"

    # 9. Câmeras (Avulso, Wifi Agregado e Cabo Agregado)
    elif "CÂMERA" in s or "CAMERA" in s:
        if "AVULSO" in s:
            base = "INSTALAÇÃO CAMERA AVULSO"
        elif "WIFI" in s:
            base = "INSTALAÇÃO CAMERA WIFI AGREGADO"
        else:
            base = "INSTALAÇÃO CAMERA AGREGADO"

    # 10. Wifi / TV / Fone / Ponto Extra (Avulso e Agregado - com barra invertida exata da tabela)
    elif "AGREGADO" in s or "AVULSO" in s or "PONTO EXTRA" in s:
        if "AVULSO" in s:
            base = r"INSTALAÇÃO WIFI\TV\FONE\PONTO EXTRA AVULSO"
        else:
            base = r"INSTALAÇÃO WIFI\TV\FONE\PONTO EXTRA AGREGADO"

    # 11. Instalação (normal até 300m e mais de 300m)
    else:
        if "301" in s or "500" in s or "MAIS DE 300" in s or "EXCEDENTE" in s or m > 300:
            base = "INSTALAÇÃO MAIS DE 300M"
        else:
            base = "INSTALAÇÃO"

    return f"{base}{sufixo_fb}"


def registrar_servico_na_planilha_tecnico(dados_os):
    tecnico = str(dados_os.get("tecnico_exec", "ANDERSON")).upper()
    if "ANDERSON" in tecnico:
        nome_planilha = "SERVIÇOS AUTOMATIZADOS ANDERSON"
    elif "MARCIO" in tecnico or "MÁRCIO" in tecnico:
        nome_planilha = "SERVIÇOS AUTOMATIZADOS MARCIO"
    elif "BRUNO" in tecnico:
        nome_planilha = "SERVIÇOS AUTOMATIZADOS BRUNO"
    else:
        st.warning(f"⚠️ Técnico não identificado para planilha: {tecnico}")
        return False

    protocolo_atual = limpar_protocolo_str(dados_os.get("protocolo", ""))
    if not protocolo_atual:
        st.warning("⚠️ Protocolo vazio. Não foi possível gravar na planilha.")
        return False

    data_str = str(dados_os.get("data", datetime.now().strftime("%d/%m/%Y")))[:10]
    dt = None
    for formato in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            dt = datetime.strptime(data_str, formato)
            break
        except ValueError:
            continue
    if not dt:
        dt = datetime.now()

    if dt.day >= 26:
        mes_ref = datetime(dt.year + 1, 1, 1) if dt.month == 12 else datetime(dt.year, dt.month + 1, 1)
        ini_mes = dt.month
        fim_mes = dt.month + 1 if dt.month < 12 else 1
    else:
        mes_ref = dt
        ini_mes = dt.month - 1 if dt.month > 1 else 12
        fim_mes = dt.month

    meses_siglas = {
        1: "Jan", 2: "Fev", 3: "Mar", 4: "Abr", 5: "Mai", 6: "Jun",
        7: "Jul", 8: "Ago", 9: "Set", 10: "Out", 11: "Nov", 12: "Dez"
    }
    nome_aba = f"26.{ini_mes:02d} a 25.{fim_mes:02d} - {meses_siglas[mes_ref.month]}.{str(mes_ref.year)[-2:]}"

    cidade_os = dados_os.get("cidade", "")
    lista_srv_bruta = dados_os.get("servicos", [])
    lista_srv = []

    for s in lista_srv_bruta:
        qtd_cabo = s.get("cabo", 0) if s.get("eh_principal", True) else s.get("lpu", 1)
        nome_plan = traduzir_servico_para_planilha(s.get("nome", ""), qtd_cabo, cidade_os)
        lista_srv.append({"nome": nome_plan, "cabo": qtd_cabo})

    if not lista_srv and dados_os.get("servico"):
        qtd_cabo = dados_os.get("cabo", 0)
        nome_plan = traduzir_servico_para_planilha(dados_os.get("servico", ""), qtd_cabo, cidade_os)
        lista_srv = [{"nome": nome_plan, "cabo": qtd_cabo}]

    payload = {
        "nome_planilha": nome_planilha,
        "nome_aba": nome_aba,
        "data_iso": dt.strftime("%Y-%m-%d"),
        "cliente": str(dados_os.get("cliente", "")).upper(),
        "protocolo": protocolo_atual,
        "cidade": cidade_os,
        "obs": dados_os.get("obs", ""),
        "servicos": lista_srv
    }

    try:
        resp = requests.post(GOOGLE_SCRIPT_PLANILHAS_URL, json=payload, timeout=20)
        dados_resp = resp.json()

        if dados_resp.get("ok"):
            aba_usada = dados_resp.get("aba", nome_aba)
            for item in dados_resp.get("resultados", []):
                srv_txt = item.get("servico", "")
                lin_txt = item.get("linha", "?")
                if item.get("ja_existia"):
                    st.info(
                        f'ℹ️ SERVIÇO "{srv_txt}" JÁ ESTAVA GRAVADO NA LINHA "{lin_txt}" '
                        f'DA PLANILHA {nome_planilha} (Aba: {aba_usada})'
                    )
                else:
                    st.success(
                        f'🟢 SERVIÇO "{srv_txt}" GRAVADO NA LINHA "{lin_txt}" '
                        f'DA PLANILHA {nome_planilha} (Aba: {aba_usada})'
                    )
            return True
        else:
            st.error(f"⚠️ Erro retornado pela planilha: {dados_resp.get('msg') or dados_resp}")
            return False
    except Exception as e:
        st.error(f"⚠️ Falha ao processar resposta da planilha: {e}")
        return False