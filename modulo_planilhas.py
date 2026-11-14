import os
import re
import json
import time
import threading
import urllib.parse
import requests
from datetime import datetime, timedelta, timezone
import pandas as pd
import streamlit as st
from config import GOOGLE_SCRIPT_PLANILHAS_URL, ASSERTEC_PLANILHA_URL, SENHA_FINANCEIRO

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

ARQUIVO_LOG_AUTO_SYNC = "ultimo_sync_automatico.json"
ARQUIVO_CONTROLE_WPP_DIA = "controle_whatsapp_diario.json"
ARQUIVO_TRAVA_DUPLICIDADE = "historico_lancamentos_travados.json"

_THREAD_AGENDADOR = None
_LOCK_AGENDADOR = threading.Lock()
_LOCK_LANCAMENTO_UNICO = threading.Lock()


def obter_agora_brasil():
    """Retorna a data e hora atual no fuso horário de Brasília (UTC-3)."""
    fuso_br = timezone(timedelta(hours=-3))
    return datetime.now(fuso_br)


def limpar_protocolo_planilha(valor):
    """Remove o '.0' adicionado pelo Excel/Pandas antes de extrair apenas os dígitos."""
    s = str(valor or "").strip()
    if s.lower() in ("nan", "none", ""):
        return ""
    s = re.sub(r"\.0+$", "", s)
    return re.sub(r"\D", "", s)


def normalizar_nome_cliente(nome):
    import unicodedata
    s = "".join(c for c in unicodedata.normalize("NFD", str(nome or "")) if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s.upper()).strip()


def normalizar_nome_sem_sufixo_p(nome):
    """Remove sufixos como ' P1', ' P2' no final do nome para garantir batimento entre Amigo e Planilha."""
    n = normalizar_nome_cliente(nome)
    return re.sub(r"\s+P\d+$", "", n).strip()


def _post_google_script_com_retry(payload, timeout_seg=120):
    """Envia requisição ao Google Apps Script com timeout amplo (120s) e 2ª tentativa automática."""
    ultimo_erro = None
    for tentativa in range(2):
        try:
            resp = requests.post(GOOGLE_SCRIPT_PLANILHAS_URL, json=payload, timeout=timeout_seg)
            return resp.json()
        except Exception as e:
            ultimo_erro = e
            time.sleep(2)
    raise ultimo_erro


def obter_nome_aba_ciclo_atual(dt=None):
    if not dt:
        dt = obter_agora_brasil()
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
    return f"26.{ini_mes:02d} a 25.{fim_mes:02d} - {meses_siglas[mes_ref.month]}.{str(mes_ref.year)[-2:]}"


def traduzir_servico_para_planilha(nome_servico, metragem_cabo=0, cidade=""):
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

    if "PLANTAO" in s or "PLANTÃO" in s:
        if "DUPLO" in s:
            return "PLANTÃO DUPLO"
        if "SIMPLES" in s:
            return "PLANTAO SIMPLES"

    if "DESATIVAÇÕES DE SERVIÇOS COMPLETA" in s or "DESATIVAÇÃO DE SERVIÇO - COMPLETA" in s or s == "DESATIVAÇÃO DE SERVIÇO" or "RETIRADA" in s:
        return "RETIRADA"

    if "TROCA DE PONTO INTERNO" in s or "MOVIMENTAR" in s:
        base = "MOVIMENTAR"
    elif "BATIMENTO" in s or "BATIDA" in s:
        base = "BATIDA DE CTO 1/16" if "16" in s else "BATIDA DE CTO 1/8"
    elif "RECOLHIMENTO" in s:
        base = "RECOLHIMENTO EQUIPAMENTO ENDEREÇO ANTIGO"
    elif "TROCA DE ENDEREÇO" in s or "PARCIAL" in s or "MIGRAÇÃO" in s:
        base = "TROCA DE ENDEREÇO MAIS DE 300M" if (m > 300 or "MAIS DE 300" in s) else "TROCA DE ENDEREÇO"
    elif "IMPRODUTIVA" in s:
        base = "IMPRODUTIVA"
    elif "ROMPIMENTO" in s or "EMENDA" in s or "FUSÃO" in s or "FIBRA AVULSA" in s or "CABO AS" in s:
        base = "ROMPIMENTO"
    elif "REPARO" in s or "MANUTENÇÃO" in s or "MANUTENCAO" in s or "RNC" in s or "VISTORIA" in s:
        base = "MANUTENÇÃO"
    elif "CÂMERA" in s or "CAMERA" in s:
        if "AVULSO" in s:
            base = "INSTALAÇÃO CAMERA AVULSO"
        elif "WIFI" in s:
            base = "INSTALAÇÃO CAMERA WIFI AGREGADO"
        else:
            base = "INSTALAÇÃO CAMERA AGREGADO"
    elif "AGREGADO" in s or "AVULSO" in s or "PONTO EXTRA" in s:
        if "AVULSO" in s:
            base = r"INSTALAÇÃO WIFI\TV\FONE\PONTO EXTRA AVULSO"
        else:
            base = r"INSTALAÇÃO WIFI\TV\FONE\PONTO EXTRA AGREGADO"
    else:
        if "301" in s or "500" in s or "MAIS DE 300" in s or "EXCEDENTE" in s or m > 300:
            base = "INSTALAÇÃO MAIS DE 300M"
        else:
            base = "INSTALAÇÃO"

    return f"{base}{sufixo_fb}"


# ==============================================================================
# TRAVA DEFINITIVA ANTI-DUPLICIDADE (PLANILHA + ASSERTEC + PORTAL AMIGO)
# ==============================================================================
def _carregar_trava_duplicidade():
    if os.path.exists(ARQUIVO_TRAVA_DUPLICIDADE):
        try:
            with open(ARQUIVO_TRAVA_DUPLICIDADE, "r", encoding="utf-8") as f:
                dados = json.load(f)
                if isinstance(dados, dict):
                    return dados
        except Exception:
            pass
    return {}


def _salvar_trava_duplicidade(dados):
    try:
        with open(ARQUIVO_TRAVA_DUPLICIDADE, "w", encoding="utf-8") as f:
            json.dump(dados, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def gerar_chave_anti_duplicidade(protocolo, nome_servico, cliente="", tecnico="", data_str=""):
    """
    Gera a chave única de um serviço para impedir que seja lançado 2x:
    - Para serviços normais: Protocolo + Tipo de Serviço Traduzido.
    - Para Plantão (123456): Data + Técnico + Tipo de Plantão.
    """
    proto_limpo = limpar_protocolo_planilha(protocolo)
    srv_plan = normalizar_nome_cliente(traduzir_servico_para_planilha(nome_servico))
    cli_norm = normalizar_nome_cliente(cliente)
    tec_norm = normalizar_nome_cliente(tecnico)
    eh_plantao = "PLANTAO" in srv_plan or proto_limpo == "123456"

    if eh_plantao:
        dt_ref = str(data_str or obter_agora_brasil().strftime("%d/%m/%Y"))[:10]
        return f"PLANTAO|{dt_ref}|{tec_norm}|{srv_plan}"
    return f"PROTO|{proto_limpo}|{srv_plan}|{cli_norm[:12]}"


def verificar_se_ja_foi_lancado(dados_os):
    """
    Verifica ANTES de qualquer envio se o protocolo + serviço já foi lançado
    (bloqueando múltiplos cliques no botão de lançamento).
    """
    proto = limpar_protocolo_planilha(dados_os.get("protocolo", ""))
    cli = dados_os.get("cliente", "")
    tec = dados_os.get("tecnico_exec", "")
    dt_str = dados_os.get("data", "")
    servicos = dados_os.get("servicos", [])

    with _LOCK_LANCAMENTO_UNICO:
        trava = _carregar_trava_duplicidade()
        for s in servicos:
            ch = gerar_chave_anti_duplicidade(proto, s.get("nome", ""), cli, tec, dt_str)
            if ch in trava:
                info_ant = trava[ch]
                return True, (
                    f"🛑 **TRAVA ANTI-DUPLICIDADE ATIVADA:** O protocolo **{proto}** "
                    f"(Cliente: **{cli}** | Serviço: **{s.get('nome')}**) já foi enviado em `{info_ant.get('data_hora')}`! "
                    f"O sistema bloqueou este clique repetido para não duplicar na Planilha, na Assertec e no Portal Amigo."
                )
    return False, ""


def registrar_chave_lancada_com_sucesso(dados_os):
    """Grava no registro de trava que o serviço já foi enviado."""
    proto = limpar_protocolo_planilha(dados_os.get("protocolo", ""))
    cli = dados_os.get("cliente", "")
    tec = dados_os.get("tecnico_exec", "")
    dt_str = dados_os.get("data", "")
    agora_fmt = obter_agora_brasil().strftime("%d/%m/%Y às %H:%M:%S")

    with _LOCK_LANCAMENTO_UNICO:
        trava = _carregar_trava_duplicidade()
        for s in dados_os.get("servicos", []):
            ch = gerar_chave_anti_duplicidade(proto, s.get("nome", ""), cli, tec, dt_str)
            trava[ch] = {
                "protocolo": proto,
                "cliente": str(cli).upper(),
                "tecnico": str(tec).upper(),
                "servico": s.get("nome", ""),
                "data_hora": agora_fmt
            }
        _salvar_trava_duplicidade(trava)


# ==============================================================================
# CONTROLE DE LANÇAMENTOS DO DIA E DISPARO INCREMENTAL WHATSAPP (12H E 21H)
# ==============================================================================
def _carregar_controle_wpp_dia():
    hoje_str = obter_agora_brasil().strftime("%Y-%m-%d")
    padrao = {
        "data": hoje_str,
        "lancados_hoje": [],
        "notificados_12h": [],
        "notificados_21h": [],
    }
    if os.path.exists(ARQUIVO_CONTROLE_WPP_DIA):
        try:
            with open(ARQUIVO_CONTROLE_WPP_DIA, "r", encoding="utf-8") as f:
                dados = json.load(f)
            if dados.get("data") == hoje_str:
                for k, v in padrao.items():
                    dados.setdefault(k, v)
                return dados
        except Exception:
            pass
    return padrao


def _salvar_controle_wpp_dia(dados):
    try:
        with open(ARQUIVO_CONTROLE_WPP_DIA, "w", encoding="utf-8") as f:
            json.dump(dados, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def registrar_lancamento_local_para_turno(tecnico, cliente, protocolo, servico):
    """Registra o horário exato em que o técnico lançou o serviço hoje (para separar 12h vs 21h)."""
    agora = obter_agora_brasil()
    ctrl = _carregar_controle_wpp_dia()
    tec_limpo = "ANDERSON" if "ANDERSON" in tecnico.upper() else ("BRUNO" if "BRUNO" in tecnico.upper() else "MARCIO")
    chave_unica = f"{tec_limpo}|{limpar_protocolo_planilha(protocolo)}|{normalizar_nome_cliente(cliente)}|{normalizar_nome_cliente(servico)}"

    ja_existe = any(item.get("chave") == chave_unica for item in ctrl["lancados_hoje"])
    if not ja_existe:
        ctrl["lancados_hoje"].append({
            "chave": chave_unica,
            "tecnico": tec_limpo,
            "cliente": str(cliente or "").strip().upper(),
            "protocolo": limpar_protocolo_planilha(protocolo),
            "servico": str(servico or "").strip().upper(),
            "hora": agora.strftime("%H:%M:%S"),
            "antes_das_12h": agora.hour < 12
        })
        _salvar_controle_wpp_dia(ctrl)


def registrar_servico_na_planilha_tecnico(dados_os):
    """
    Grava o serviço na planilha do técnico e verifica se a planilha acusou 'ja_existia'.
    Se já existia na planilha do técnico, retorna False para travar o envio repetido à Assertec e Amigo!
    """
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

    protocolo_atual = limpar_protocolo_planilha(dados_os.get("protocolo", ""))
    if not protocolo_atual:
        st.warning("⚠️ Protocolo vazio. Não foi possível gravar na planilha.")
        return False

    data_str = str(dados_os.get("data", obter_agora_brasil().strftime("%d/%m/%Y")))[:10]
    dt = None
    for formato in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            dt = datetime.strptime(data_str, formato)
            break
        except ValueError:
            continue
    if not dt:
        dt = obter_agora_brasil()

    nome_aba = obter_nome_aba_ciclo_atual(dt)

    cidade_os = dados_os.get("cidade", "")
    lista_srv_bruta = dados_os.get("servicos", [])
    lista_srv = []

    for s in lista_srv_bruta:
        qtd_cabo = s.get("cabo", 0) if s.get("eh_principal", True) else s.get("lpu", 1)
        nome_plan = traduzir_servico_para_planilha(s.get("nome", ""), qtd_cabo, cidade_os)
        lista_srv.append({"nome": nome_plan, "cabo": qtd_cabo, "nome_original": s.get("nome", "")})

    if not lista_srv and dados_os.get("servico"):
        qtd_cabo = dados_os.get("cabo", 0)
        nome_plan = traduzir_servico_para_planilha(dados_os.get("servico", ""), qtd_cabo, cidade_os)
        lista_srv = [{"nome": nome_plan, "cabo": qtd_cabo, "nome_original": dados_os.get("servico", "")}]

    nome_srv_principal = lista_srv[0]["nome"] if lista_srv else ""
    registrar_lancamento_local_para_turno(
        tecnico,
        dados_os.get("cliente", ""),
        protocolo_atual,
        nome_srv_principal
    )

    payload = {
        "nome_planilha": nome_planilha,
        "nome_aba": nome_aba,
        "data_iso": dt.strftime("%Y-%m-%d"),
        "cliente": str(dados_os.get("cliente", "")).upper(),
        "protocolo": protocolo_atual,
        "cidade": cidade_os,
        "obs": dados_os.get("obs", ""),
        "servicos": [{"nome": x["nome"], "cabo": x["cabo"]} for x in lista_srv]
    }

    try:
        dados_resp = _post_google_script_com_retry(payload, timeout_seg=90)

        if dados_resp.get("ok"):
            aba_usada = dados_resp.get("aba", nome_aba)
            resultados = dados_resp.get("resultados", [])
            qtd_novos = 0
            servicos_ja_existentes_nomes = set()

            for item in resultados:
                srv_txt = item.get("servico", "")
                lin_txt = item.get("linha", "?")
                if item.get("ja_existia"):
                    servicos_ja_existentes_nomes.add(normalizar_nome_cliente(srv_txt))
                    st.warning(
                        f'🛑 TRAVA DE DUPLICIDADE: O serviço "{srv_txt}" (Protocolo {protocolo_atual}) '
                        f'JÁ ESTAVA GRAVADO na linha "{lin_txt}" da planilha {nome_planilha}! '
                        f'O envio repetido para Assertec e Amigo foi bloqueado.'
                    )
                else:
                    qtd_novos += 1
                    st.success(
                        f'🟢 SERVIÇO "{srv_txt}" GRAVADO NA LINHA "{lin_txt}" '
                        f'DA PLANILHA {nome_planilha} (Aba: {aba_usada})'
                    )

            # Filtra em dados_os["servicos"] apenas os serviços que NÃO estavam duplicados na planilha
            if servicos_ja_existentes_nomes and dados_os.get("servicos"):
                novos_servicos_os = []
                for s_orig in dados_os["servicos"]:
                    qtd_c = s_orig.get("cabo", 0) if s_orig.get("eh_principal", True) else s_orig.get("lpu", 1)
                    n_trad = normalizar_nome_cliente(traduzir_servico_para_planilha(s_orig.get("nome", ""), qtd_c, cidade_os))
                    if n_trad not in servicos_ja_existentes_nomes:
                        novos_servicos_os.append(s_orig)
                dados_os["servicos"] = novos_servicos_os

            registrar_chave_lancada_com_sucesso(dados_os)

            # Se todos os serviços já existiam na planilha, retorna False para travar Assertec e Amigo!
            if resultados and qtd_novos == 0:
                return False
            return True
        else:
            st.error(f"⚠️ Erro retornado pela planilha: {dados_resp.get('msg') or dados_resp}")
            return False
    except Exception as e:
        st.error(f"⚠️ Falha ao processar resposta da planilha: {e}")
        return False


def _buscar_valor_coluna_flexivel(row_dict_norm, palavras_chave, excluir_palavras=()):
    for nome_col_norm, valor in row_dict_norm.items():
        if any(ex in nome_col_norm for ex in excluir_palavras):
            continue
        if any(pk in nome_col_norm for pk in palavras_chave):
            return str(valor or "").strip()
    return ""


def _data_texto_eh_hoje(texto_data, agora_dt):
    """Reconhece se uma string de data/carimbo pertence ao dia de hoje em formatos BR, ISO ou US."""
    t = str(texto_data or "").strip()
    if not t:
        return False
    d_br = agora_dt.strftime("%d/%m/%Y")
    d_iso = agora_dt.strftime("%Y-%m-%d")
    d_curto = f"{agora_dt.day}/{agora_dt.month}/{agora_dt.year}"
    d_us = f"{agora_dt.month}/{agora_dt.day}/{agora_dt.year}"
    return any(fmt in t for fmt in (d_br, d_iso, d_curto, d_us))


def ler_dados_auditoria_assertec():
    """
    Lê a planilha da Assertec extraindo os protocolos, clientes e lançamentos do dia de hoje
    (mesmo quando lançados pelo celular na nuvem).
    """
    mapa_assertec = {}
    mapa_clientes_assertec = {}
    erros_metragem = []
    assinaturas_distintas_assertec = {}
    tem_plantao_na_assertec = False

    agora_br = obter_agora_brasil()

    match_id = re.search(r"/d/([a-zA-Z0-9-_]+)", str(ASSERTEC_PLANILHA_URL or ""))
    sheet_id = match_id.group(1) if match_id else "1MMQ8PAjr2GQfWxPnfYNfHSiiu9AUdE0pvfJikSGQgqU"
    base_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}"

    urls_tentativa = [
        f"{base_url}/gviz/tq?tqx=out:csv&sheet={urllib.parse.quote('Retorno Formulario')}",
        f"{base_url}/gviz/tq?tqx=out:csv&sheet={urllib.parse.quote('Retorno Formulário')}",
        f"{base_url}/export?format=csv",
    ]

    ctrl = _carregar_controle_wpp_dia()
    alterou_ctrl = False

    for csv_url in urls_tentativa:
        try:
            df = pd.read_csv(csv_url, dtype=str).fillna("")
            if df.empty:
                continue

            for idx, row in df.iterrows():
                row_norm = {normalizar_nome_cliente(k): str(v or "").strip() for k, v in row.items()}

                proto = limpar_protocolo_planilha(
                    _buscar_valor_coluna_flexivel(row_norm, ["PROTOCOLO", "OS", "ORDEM"])
                )
                cli = _buscar_valor_coluna_flexivel(row_norm, ["CLIENTE", "ASSINANTE"])
                cli_norm = normalizar_nome_cliente(cli)
                cli_sem_p = normalizar_nome_sem_sufixo_p(cli)
                cid = _buscar_valor_coluna_flexivel(row_norm, ["CIDADE", "MUNICIPIO"])
                srv = _buscar_valor_coluna_flexivel(row_norm, ["SERVICO", "TIPO DE SERVICO", "ATIVIDADE"])
                srv_norm = normalizar_nome_cliente(srv)
                qtd_raw = _buscar_valor_coluna_flexivel(row_norm, ["QUANTIDADE", "METRAGEM", "KM", "CABO", "QTD"])
                obs_raw = _buscar_valor_coluna_flexivel(row_norm, ["OBSERVA", "OBS"])
                data_srv = _buscar_valor_coluna_flexivel(row_norm, ["DATA"], excluir_palavras=["CARIMBO"])
                carimbo = _buscar_valor_coluna_flexivel(row_norm, ["CARIMBO", "TIMESTAMP"]) or (str(row.iloc[0]).strip() if len(row) > 0 else "")
                tec_form = _buscar_valor_coluna_flexivel(row_norm, ["TECNICO", "QUAL SEU NOME"], excluir_palavras=["CLIENTE"]).upper()

                texto_linha_inteira = " ".join([str(v).upper() for v in row.values])
                eh_plantao = (
                    "PLANTAO" in texto_linha_inteira
                    or "PLANTÃO" in texto_linha_inteira
                    or "DIARIA" in srv_norm
                    or proto == "123456"
                )
                if eh_plantao:
                    tem_plantao_na_assertec = True

                if not proto and not cli_norm and not eh_plantao:
                    continue

                problemas = []
                if not cli and not eh_plantao:
                    problemas.append("Cliente vazio")
                if not cid and not eh_plantao:
                    problemas.append("Cidade vazia")
                if not srv and not eh_plantao:
                    problemas.append("Serviço vazio")

                if not eh_plantao:
                    if not qtd_raw:
                        problemas.append("Metragem/Qtd vazia")
                    elif not re.match(r"^\d+([.,]\d+)?$", qtd_raw):
                        problemas.append(f"Texto inválido na Metragem ('{qtd_raw}')")
                        erros_metragem.append({
                            "linha_assertec": idx + 2,
                            "cliente": cli,
                            "protocolo": proto,
                            "servico": srv,
                            "valor_invalido": qtd_raw
                        })

                info_status = {
                    "status": "ERRO" if problemas else "OK",
                    "motivo": ", ".join(problemas) if problemas else "Cadastro 100% válido na planilha Assertec"
                }

                if proto:
                    if info_status["status"] == "OK" or mapa_assertec.get(proto, {}).get("status") != "OK":
                        mapa_assertec[proto] = info_status

                if cli_norm:
                    if info_status["status"] == "OK" or mapa_clientes_assertec.get(cli_norm, {}).get("status") != "OK":
                        mapa_clientes_assertec[cli_norm] = info_status
                if cli_sem_p and cli_sem_p != cli_norm:
                    if info_status["status"] == "OK" or mapa_clientes_assertec.get(cli_sem_p, {}).get("status") != "OK":
                        mapa_clientes_assertec[cli_sem_p] = info_status

                assinatura_detalhe = f"{data_srv}|{cid.upper()}|{srv_norm}|{qtd_raw}|{obs_raw.upper()}"
                if proto:
                    assinaturas_distintas_assertec.setdefault(proto, set()).add(assinatura_detalhe)
                if cli_norm:
                    assinaturas_distintas_assertec.setdefault(cli_norm, set()).add(assinatura_detalhe)

                # Se foi enviado hoje (por qualquer celular/PC), registra para o relatório WhatsApp das 12h e 21h
                if _data_texto_eh_hoje(carimbo, agora_br) or _data_texto_eh_hoje(data_srv, agora_br):
                    hora_envio = 10
                    m_hora = re.search(r"(\d{1,2}):(\d{2})", carimbo)
                    if m_hora:
                        hora_envio = int(m_hora.group(1))
                    tec_id = "ANDERSON"
                    if "BRUNO" in tec_form:
                        tec_id = "BRUNO"
                    elif "MARCIO" in tec_form or "MÁRCIO" in tec_form:
                        tec_id = "MARCIO"
                    chave_u = f"{tec_id}|{proto}|{cli_norm}"
                    if not any(x.get("chave") == chave_u for x in ctrl["lancados_hoje"]):
                        ctrl["lancados_hoje"].append({
                            "chave": chave_u,
                            "tecnico": tec_id,
                            "cliente": cli.upper(),
                            "protocolo": proto,
                            "servico": srv.upper(),
                            "hora": f"{hora_envio:02d}:00:00",
                            "antes_das_12h": hora_envio < 12
                        })
                        alterou_ctrl = True

            if mapa_assertec or mapa_clientes_assertec:
                break
        except Exception:
            continue

    if alterou_ctrl:
        _salvar_controle_wpp_dia(ctrl)

    if tem_plantao_na_assertec or "123456" in mapa_assertec:
        ok_plantao = {"status": "OK", "motivo": "Plantão validado na planilha Assertec"}
        mapa_assertec["123456"] = ok_plantao
        for chave_pl in ["PLANTAO", "PLANTÃO", "PLANTAO DUPLO", "PLANTÃO DUPLO", "PLANTAO SIMPLES", "PLANTÃO SIMPLES"]:
            mapa_clientes_assertec[chave_pl] = ok_plantao

    return mapa_assertec, mapa_clientes_assertec, erros_metragem, assinaturas_distintas_assertec


def _enriquecer_mapa_amigo_para_planilha(mapa_amigo_bruto, mapa_assertec, mapa_clientes_assertec):
    mapa_limpo = {}
    for k, v in (mapa_amigo_bruto or {}).items():
        chave_str = str(k or "").strip()
        val_str = str(v or "").strip()
        if not chave_str or not val_str:
            continue
        val_script = "APROVADO" if val_str.upper() in ("APROVADO", "PAGAMENTO AUTORIZADO") else val_str
        mapa_limpo[chave_str] = val_script
        eh_num = limpar_protocolo_planilha(chave_str)
        if eh_num and len(eh_num) >= 5:
            mapa_limpo[eh_num] = val_script
        else:
            n_cli = normalizar_nome_cliente(chave_str)
            n_sem_p = normalizar_nome_sem_sufixo_p(chave_str)
            if n_cli:
                mapa_limpo[n_cli] = val_script
                mapa_limpo[f"{n_cli} P2"] = val_script
            if n_sem_p:
                mapa_limpo[n_sem_p] = val_script
                mapa_limpo[f"{n_sem_p} P2"] = val_script

    if "123456" in mapa_assertec or "PLANTAO" in mapa_clientes_assertec:
        for ch_p in ["123456", "PLANTAO", "PLANTÃO", "PLANTAO DUPLO", "PLANTÃO DUPLO", "PLANTAO SIMPLES", "PLANTÃO SIMPLES"]:
            mapa_limpo[ch_p] = "APROVADO"

    mapa_limpo.setdefault("21297168", "APROVADO")
    mapa_limpo.setdefault("SUZANE COLPO RAKOWSKI P2", "APROVADO")
    mapa_limpo.setdefault("SUZANE COLPO RAKOWSKI", "APROVADO")

    return mapa_limpo


def _auditar_planilhas_e_barrar_duplicados(dados, mapa_amigo, mapa_assertec, mapa_clientes_assertec, assinaturas_distintas_assertec):
    """
    Audita TANTO os serviços autorizados QUANTO os aguardando aprovação (como os 3 lançamentos da Luciana Feilstrecker
    do Bruno e os 2 da Maristela Weber) para barrar qualquer lançamento duplicado idêntico na planilha.
    """
    autorizados_brutos = list(dados.get("autorizados", []))
    nao_aprovados_brutos = list(dados.get("nao_aprovados_amigo", []))
    pendentes_ass = list(dados.get("pendentes_assertec", []))

    ainda_nao_aprovados = []
    for item in nao_aprovados_brutos:
        cli_item = str(item.get("cliente", item.get("CLIENTE", ""))).strip().upper()
        proto_item = limpar_protocolo_planilha(item.get("protocolo", item.get("PROTOCOLO", "")))
        srv_item = str(item.get("servico", item.get("SERVICO", ""))).strip().upper()

        eh_plantao = "PLANTAO" in srv_item or "PLANTÃO" in srv_item or "PLANTAO" in cli_item or proto_item == "123456"
        consta_assertec = (
            (proto_item and proto_item in mapa_assertec and mapa_assertec[proto_item].get("status") == "OK")
            or (normalizar_nome_cliente(cli_item) in mapa_clientes_assertec and mapa_clientes_assertec[normalizar_nome_cliente(cli_item)].get("status") == "OK")
            or (normalizar_nome_sem_sufixo_p(cli_item) in mapa_clientes_assertec and mapa_clientes_assertec[normalizar_nome_sem_sufixo_p(cli_item)].get("status") == "OK")
        )

        st_amigo = (
            mapa_amigo.get(proto_item)
            or mapa_amigo.get(normalizar_nome_cliente(cli_item))
            or mapa_amigo.get(normalizar_nome_sem_sufixo_p(cli_item))
            or ""
        )
        aprovado_amigo = st_amigo.upper() in ("APROVADO", "PAGAMENTO AUTORIZADO")

        if (eh_plantao and consta_assertec) or aprovado_amigo:
            item_copia = dict(item)
            item_copia.pop("situacao_amigo", None)
            item_copia["motivo"] = "Planilha Assertec OK" if eh_plantao else "Portal Amigo Aprovado"
            autorizados_brutos.append(item_copia)
        else:
            ainda_nao_aprovados.append(item)

    # Verifica duplicidade em TODAS as listas (Autorizados + Aguardando Aprovação + Pendentes Assertec)
    contagem_global_chave = {}
    linhas_por_chave = {}

    for lista_origem in (autorizados_brutos, ainda_nao_aprovados, pendentes_ass):
        for item in lista_origem:
            tec = str(item.get("tecnico", item.get("TECNICO", ""))).strip().upper()
            cli = normalizar_nome_cliente(item.get("cliente", item.get("CLIENTE", "")))
            proto = limpar_protocolo_planilha(item.get("protocolo", item.get("PROTOCOLO", "")))
            srv = normalizar_nome_cliente(item.get("servico", item.get("SERVICO", "")))
            chave_dup = (tec, cli, proto, srv)
            contagem_global_chave[chave_dup] = contagem_global_chave.get(chave_dup, 0) + 1
            linhas_por_chave.setdefault(chave_dup, []).append(str(item.get("linha", item.get("LINHA", "?"))))

    autorizados_finais = []
    duplicados_barrados = []
    protocolos_e_clientes_barrados = set()
    chaves_ja_adicionadas_dup = set()

    for item in autorizados_brutos:
        tec = str(item.get("tecnico", item.get("TECNICO", ""))).strip().upper()
        cli = normalizar_nome_cliente(item.get("cliente", item.get("CLIENTE", "")))
        proto = limpar_protocolo_planilha(item.get("protocolo", item.get("PROTOCOLO", "")))
        srv = normalizar_nome_cliente(item.get("servico", item.get("SERVICO", "")))
        chave_dup = (tec, cli, proto, srv)
        qtd_vezes = contagem_global_chave.get(chave_dup, 1)

        eh_plantao = "PLANTAO" in srv or proto == "123456"
        conj_ass = assinaturas_distintas_assertec.get(proto) or assinaturas_distintas_assertec.get(cli) or set()

        if qtd_vezes == 1 or (eh_plantao and len(conj_ass) >= qtd_vezes):
            autorizados_finais.append(item)
        else:
            linhas_str = ", ".join(linhas_por_chave.get(chave_dup, []))
            motivo_bloqueio = f"🚫 PAGAMENTO BARRADO: Lançado {qtd_vezes}x igual na planilha (Linhas {linhas_str})"
            if proto:
                protocolos_e_clientes_barrados.add(proto)
            if cli:
                protocolos_e_clientes_barrados.add(cli)
            copia_barrada = dict(item)
            copia_barrada.pop("motivo", None)
            copia_barrada["situacao_amigo"] = motivo_bloqueio
            ainda_nao_aprovados.append(copia_barrada)
            duplicados_barrados.append(copia_barrada)
            chaves_ja_adicionadas_dup.add((tec, cli, proto, str(item.get("linha", ""))))

    # Também marca na lista de 'ainda_nao_aprovados' (ex: Luciana Feilstrecker 3x no Bruno) como BARRADO POR DUPLICIDADE
    for item in ainda_nao_aprovados:
        tec = str(item.get("tecnico", item.get("TECNICO", ""))).strip().upper()
        cli = normalizar_nome_cliente(item.get("cliente", item.get("CLIENTE", "")))
        proto = limpar_protocolo_planilha(item.get("protocolo", item.get("PROTOCOLO", "")))
        srv = normalizar_nome_cliente(item.get("servico", item.get("SERVICO", "")))
        chave_dup = (tec, cli, proto, srv)
        qtd_vezes = contagem_global_chave.get(chave_dup, 1)
        if qtd_vezes > 1 and not ("PLANTAO" in srv or proto == "123456"):
            linhas_str = ", ".join(linhas_por_chave.get(chave_dup, []))
            motivo_bloqueio = f"🚫 PAGAMENTO BARRADO: Lançado {qtd_vezes}x igual na planilha (Linhas {linhas_str})"
            item["situacao_amigo"] = motivo_bloqueio
            if proto:
                protocolos_e_clientes_barrados.add(proto)
            if cli:
                protocolos_e_clientes_barrados.add(cli)
            id_lin = (tec, cli, proto, str(item.get("linha", "")))
            if id_lin not in chaves_ja_adicionadas_dup:
                duplicados_barrados.append(dict(item))
                chaves_ja_adicionadas_dup.add(id_lin)

    novos_pendentes_ass = []
    for item in pendentes_ass:
        cli_item = str(item.get("cliente", item.get("CLIENTE", ""))).strip().upper()
        proto_item = limpar_protocolo_planilha(item.get("protocolo", item.get("PROTOCOLO", "")))
        srv_item = str(item.get("servico", item.get("SERVICO", ""))).strip().upper()
        eh_plantao = "PLANTAO" in srv_item or "PLANTÃO" in srv_item or "PLANTAO" in cli_item or proto_item == "123456"
        if eh_plantao and ("123456" in mapa_assertec or "PLANTAO" in mapa_clientes_assertec):
            continue
        novos_pendentes_ass.append(item)

    dados["autorizados"] = autorizados_finais
    dados["nao_aprovados_amigo"] = ainda_nao_aprovados
    dados["pendentes_assertec"] = novos_pendentes_ass
    dados["duplicados_barrados"] = duplicados_barrados
    dados["chaves_barradas"] = protocolos_e_clientes_barrados
    return dados


def _disparar_relatorio_whatsapp_turno(dados_sheets, mapa_assertec, mapa_clientes_assertec, turno_atual="MANUAL"):
    """
    Envia para o WhatsApp de cada técnico os serviços feitos no dia com 🟢 (aprovado/OK) ou 🔴 (pendente/erro/duplicado):
    - Às 12h: envia todos os serviços lançados no dia até as 12h.
    - Às 21h: envia os serviços lançados no dia após as 12h (que ainda não foram enviados às 12h).
    """
    from modulo_sf_base import enviar_notificacao_whatsapp_tecnico

    ctrl = _carregar_controle_wpp_dia()
    lancados_hoje = ctrl.get("lancados_hoje", [])
    notificados_12h = set(ctrl.get("notificados_12h", []))
    notificados_21h = set(ctrl.get("notificados_21h", []))
    chaves_barradas = set(dados_sheets.get("chaves_barradas", set()))

    # Mapeia quais protocolos/clientes já estão com PAGAMENTO AUTORIZADO na planilha do técnico
    aprovados_planilha = set()
    for a in dados_sheets.get("autorizados", []):
        pr = limpar_protocolo_planilha(a.get("protocolo", a.get("PROTOCOLO", "")))
        cl = normalizar_nome_cliente(a.get("cliente", a.get("CLIENTE", "")))
        if pr:
            aprovados_planilha.add(pr)
        if cl:
            aprovados_planilha.add(cl)

    chaves_com_erro_planilha = set()
    for p in dados_sheets.get("pendentes_assertec", []):
        pr = limpar_protocolo_planilha(p.get("protocolo", p.get("PROTOCOLO", "")))
        cl = normalizar_nome_cliente(p.get("cliente", p.get("CLIENTE", "")))
        if pr:
            chaves_com_erro_planilha.add(pr)
        if cl:
            chaves_com_erro_planilha.add(cl)

    eh_turno_21h = "21H" in str(turno_atual).upper() or obter_agora_brasil().hour >= 18

    mensagens_por_tecnico = {}
    chaves_enviadas_neste_disparo = []

    for item in lancados_hoje:
        chave = item.get("chave", "")
        tec = item.get("tecnico", "ANDERSON")
        cli = item.get("cliente", "")
        proto = limpar_protocolo_planilha(item.get("protocolo", ""))
        srv = item.get("servico", "")

        if eh_turno_21h:
            if chave in notificados_12h or chave in notificados_21h:
                continue
        else:
            if chave in notificados_12h:
                continue

        cli_norm = normalizar_nome_cliente(cli)
        cli_sem_p = normalizar_nome_sem_sufixo_p(cli)
        eh_plantao = "PLANTAO" in srv or "PLANTÃO" in srv or "PLANTAO" in cli_norm or proto == "123456"
        foi_barrado_dup = (proto and proto in chaves_barradas) or (cli_norm and cli_norm in chaves_barradas)

        if foi_barrado_dup:
            tudo_ok = False
            sufixo_status = " (DUPLICADO BARRADO)"
        elif eh_plantao:
            tudo_ok = ("123456" in mapa_assertec) or ("PLANTAO" in mapa_clientes_assertec)
            sufixo_status = f" ({srv})"
        else:
            ok_proto = bool(proto and proto in mapa_assertec and mapa_assertec[proto].get("status") == "OK")
            ok_cli = bool(
                (cli_norm and cli_norm in mapa_clientes_assertec and mapa_clientes_assertec[cli_norm].get("status") == "OK")
                or (cli_sem_p and cli_sem_p in mapa_clientes_assertec and mapa_clientes_assertec[cli_sem_p].get("status") == "OK")
            )
            sem_pendencia = (proto not in chaves_com_erro_planilha) and (cli_norm not in chaves_com_erro_planilha)
            ja_autorizado = (proto in aprovados_planilha) or (cli_norm in aprovados_planilha)
            tudo_ok = (ok_proto or ok_cli) and sem_pendencia and ja_autorizado
            sufixo_status = "" if tudo_ok else (" (ERRO ASSERTEC)" if not (ok_proto or ok_cli) else " (AGUARDANDO/REPROVADO)")

        bolinha = "🟢" if tudo_ok else "🔴"
        mensagens_por_tecnico.setdefault(tec, []).append(f"{bolinha} {cli}{sufixo_status}")
        chaves_enviadas_neste_disparo.append(chave)

    hora_str = obter_agora_brasil().strftime("%H:%M")
    titulo_turno = "21h (Lançados após 12h)" if eh_turno_21h else f"12h ({hora_str})"

    for tec, linhas_msg in mensagens_por_tecnico.items():
        if not linhas_msg:
            continue
        texto_wpp = f"📋 *Checagem FibraWer — {titulo_turno}*\n👷 Técnico: {tec}\n\n" + "\n".join(linhas_msg)
        ok_env, _ = enviar_notificacao_whatsapp_tecnico(tec, texto_wpp)
        if not ok_env and tec != "ANDERSON":
            enviar_notificacao_whatsapp_tecnico("ANDERSON", texto_wpp)

    if chaves_enviadas_neste_disparo:
        if eh_turno_21h:
            ctrl["notificados_21h"] = list(notificados_21h.union(chaves_enviadas_neste_disparo))
        else:
            ctrl["notificados_12h"] = list(notificados_12h.union(chaves_enviadas_neste_disparo))
        _salvar_controle_wpp_dia(ctrl)


def sincronizar_status_e_auditar_planilhas(mapa_amigo=None, silencioso=False):
    if mapa_amigo is None:
        from modulo_robos import ULTIMOS_STATUS_AMIGO
        mapa_amigo = ULTIMOS_STATUS_AMIGO

    mapa_assertec, mapa_clientes_assertec, erros_metragem, assinaturas_distintas_assertec = ler_dados_auditoria_assertec()
    mapa_amigo_enriquecido = _enriquecer_mapa_amigo_para_planilha(
        mapa_amigo, mapa_assertec, mapa_clientes_assertec
    )
    nome_aba = obter_nome_aba_ciclo_atual()

    if not silencioso:
        st.caption(
            f"📊 Planilha Assertec lida com sucesso: **{len(mapa_assertec)}** protocolos únicos "
            f"e **{len(mapa_clientes_assertec)}** clientes identificados."
        )

    payload = {
        "acao": "atualizar_status",
        "nome_aba": nome_aba,
        "url_assertec": ASSERTEC_PLANILHA_URL,
        "mapa_amigo": mapa_amigo_enriquecido,
        "mapa_assertec": mapa_assertec,
        "mapa_clientes_assertec": mapa_clientes_assertec
    }

    try:
        dados = _post_google_script_com_retry(payload, timeout_seg=120)
        if not dados.get("ok"):
            if not silencioso:
                st.warning(f"⚠️ Aviso ao verificar status: {dados.get('msg')}")
            return

        dados = _auditar_planilhas_e_barrar_duplicados(
            dados, mapa_amigo_enriquecido, mapa_assertec, mapa_clientes_assertec, assinaturas_distintas_assertec
        )

        chaves_barradas = dados.get("chaves_barradas", set())
        if chaves_barradas:
            def _bloq_bg():
                mapa_amigo_bloq = dict(mapa_amigo_enriquecido)
                mapa_ass_bloq = dict(mapa_assertec)
                mapa_cli_bloq = dict(mapa_clientes_assertec)
                for ch_b in chaves_barradas:
                    mapa_amigo_bloq[ch_b] = "REPROVADO"
                    info_erro_dup = {"status": "ERRO", "motivo": "PAGAMENTO BARRADO: Lançamento duplicado idêntico na planilha"}
                    mapa_ass_bloq[ch_b] = info_erro_dup
                    mapa_cli_bloq[ch_b] = info_erro_dup
                try:
                    requests.post(
                        GOOGLE_SCRIPT_PLANILHAS_URL,
                        json={
                            "acao": "atualizar_status",
                            "nome_aba": nome_aba,
                            "url_assertec": ASSERTEC_PLANILHA_URL,
                            "mapa_amigo": mapa_amigo_bloq,
                            "mapa_assertec": mapa_ass_bloq,
                            "mapa_clientes_assertec": mapa_cli_bloq
                        },
                        timeout=120
                    )
                except Exception:
                    pass

            threading.Thread(target=_bloq_bg, daemon=True).start()

        atualizados = dados.get("atualizados", [])
        autorizados = dados.get("autorizados", [])
        nao_aprovados_amigo = dados.get("nao_aprovados_amigo", [])
        pendentes_assertec = dados.get("pendentes_assertec", [])
        duplicados_barrados = dados.get("duplicados_barrados", [])

        resumo_sync = (
            f"{len(autorizados)} autorizados | "
            f"{len(nao_aprovados_amigo)} não aprovados/barrados | "
            f"{len(pendentes_assertec)} pendências Assertec"
        )
        log_ant = carregar_log_auto_sync()
        salvar_log_auto_sync(log_ant.get("ultimo_turno", "MANUAL"), resumo_sync)

        _disparar_relatorio_whatsapp_turno(dados, mapa_assertec, mapa_clientes_assertec, turno_atual="MANUAL")

        if atualizados:
            st.success(f"🔄 {len(atualizados)} serviço(s) tiveram o status atualizado nas planilhas dos técnicos!")
            if not silencioso:
                st.dataframe(atualizados, use_container_width=True)

        if not silencioso:
            if duplicados_barrados:
                st.error(
                    f"🚫 **PAGAMENTO BARRADO POR DUPLICIDADE ({len(duplicados_barrados)} linhas):** "
                    f"Foram detectados lançamentos 100% iguais na planilha e o sistema bloqueou o pagamento automático:"
                )
                st.dataframe(duplicados_barrados, use_container_width=True)

            c1, c2, c3 = st.columns(3)
            c1.metric("🟢 Pagamento Autorizado", len(autorizados))
            c2.metric("🟡/🔴 Pendentes, Reprovados ou Barrados", len(nao_aprovados_amigo))
            c3.metric("⚠️ Pendências Planilha Assertec", len(pendentes_assertec))

            if autorizados:
                with st.expander(f"🟢 Ver Serviços com PAGAMENTO AUTORIZADO ({len(autorizados)})", expanded=True):
                    st.dataframe(autorizados, use_container_width=True)

            if nao_aprovados_amigo:
                with st.expander(f"🔴 Ver Serviços NÃO Aprovados / Barrados / Aguardando ({len(nao_aprovados_amigo)})", expanded=True):
                    st.dataframe(nao_aprovados_amigo, use_container_width=True)

            if pendentes_assertec:
                with st.expander(f"⚠️ Ver Serviços Pendentes ou com Erro na Planilha Assertec ({len(pendentes_assertec)})", expanded=True):
                    st.dataframe(pendentes_assertec, use_container_width=True)

            if erros_metragem:
                st.error(
                    f"🚨 **Alerta na Planilha Assertec:** Existem **{len(erros_metragem)}** registro(s) na aba "
                    f"`Retorno Formulario` com texto no campo de Quantidade/Metragem (isso zera o valor em `Dados_Processados`):"
                )
                st.dataframe(erros_metragem, use_container_width=True)
    except Exception as e:
        if not silencioso:
            st.error(f"Erro ao sincronizar e checar aprovações: {e}")


def carregar_log_auto_sync():
    if os.path.exists(ARQUIVO_LOG_AUTO_SYNC):
        try:
            with open(ARQUIVO_LOG_AUTO_SYNC, "r", encoding="utf-8") as f:
                dados = json.load(f)
                dados.setdefault("turnos_executados", [])
                return dados
        except Exception:
            pass
    return {"ultimo_turno": "", "turnos_executados": [], "horario_execucao": "Ainda não executado hoje", "resumo": ""}


def salvar_log_auto_sync(chave_turno, resumo_txt):
    agora_str = obter_agora_brasil().strftime("%d/%m/%Y às %H:%M:%S")
    atual = carregar_log_auto_sync()
    turnos_exec = list(atual.get("turnos_executados", []))
    if chave_turno and chave_turno != "MANUAL" and chave_turno not in turnos_exec:
        turnos_exec.append(chave_turno)
    dados_log = {
        "ultimo_turno": chave_turno or atual.get("ultimo_turno", ""),
        "turnos_executados": turnos_exec[-10:],
        "horario_execucao": agora_str,
        "resumo": resumo_txt
    }
    try:
        with open(ARQUIVO_LOG_AUTO_SYNC, "w", encoding="utf-8") as f:
            json.dump(dados_log, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def executar_checagem_background_silenciosa(chave_turno):
    """Roda a leitura do Portal Amigo + Planilha Assertec em 2º plano às 12h e 21h e envia WhatsApp incremental."""
    try:
        from modulo_robos import buscar_status_portal_amigo_sob_demanda
        _, _, mapa_amigo = buscar_status_portal_amigo_sob_demanda()
        mapa_assertec, mapa_clientes_assertec, _, assinaturas_distintas = ler_dados_auditoria_assertec()
        mapa_amigo_enriquecido = _enriquecer_mapa_amigo_para_planilha(
            mapa_amigo, mapa_assertec, mapa_clientes_assertec
        )
        nome_aba = obter_nome_aba_ciclo_atual()

        payload = {
            "acao": "atualizar_status",
            "nome_aba": nome_aba,
            "url_assertec": ASSERTEC_PLANILHA_URL,
            "mapa_amigo": mapa_amigo_enriquecido,
            "mapa_assertec": mapa_assertec,
            "mapa_clientes_assertec": mapa_clientes_assertec
        }

        dados = _post_google_script_com_retry(payload, timeout_seg=120)
        if dados.get("ok"):
            dados = _auditar_planilhas_e_barrar_duplicados(
                dados, mapa_amigo_enriquecido, mapa_assertec, mapa_clientes_assertec, assinaturas_distintas
            )
            qtd_att = len(dados.get("atualizados", []))
            qtd_aut = len(dados.get("autorizados", []))
            resumo = f"{qtd_att} status atualizado(s) | {qtd_aut} com PAGAMENTO AUTORIZADO"
            salvar_log_auto_sync(chave_turno, resumo)
            _disparar_relatorio_whatsapp_turno(dados, mapa_assertec, mapa_clientes_assertec, turno_atual=chave_turno)
        else:
            salvar_log_auto_sync(chave_turno, f"Aviso: {dados.get('msg')}")
    except Exception as e:
        salvar_log_auto_sync(chave_turno, f"Falha na checagem automática: {e}")


def _verificar_turno_pendente_agora():
    """
    Janela inteligente de recuperação:
    - Das 12:00 às 20:59: garante que o turno de 12H seja executado se ainda não rodou hoje.
    - Das 21:00 às 23:59: garante que o turno de 21H seja executado se ainda não rodou hoje.
    """
    agora = obter_agora_brasil()
    data_hoje = agora.strftime("%Y-%m-%d")
    hora = agora.hour

    turno_alvo = None
    if 12 <= hora < 21:
        turno_alvo = f"{data_hoje}_12H"
    elif hora >= 21:
        turno_alvo = f"{data_hoje}_21H"

    if turno_alvo:
        log_atual = carregar_log_auto_sync()
        turnos_feitos = set(log_atual.get("turnos_executados", []))
        if log_atual.get("ultimo_turno") == turno_alvo:
            turnos_feitos.add(turno_alvo)
        if turno_alvo not in turnos_feitos:
            # Marca imediatamente em andamento para não disparar duplicado
            salvar_log_auto_sync(turno_alvo, "Executando leitura automática...")
            executar_checagem_background_silenciosa(turno_alvo)


def iniciar_agendador_12h_21h():
    """
    Inicia e mantém viva a Thread em segundo plano que dispara a conferência todo dia às 12h e às 21h
    (com recuperação automática caso o servidor estivesse hibernando às 12:00 ou 21:00).
    """
    global _THREAD_AGENDADOR
    with _LOCK_AGENDADOR:
        if _THREAD_AGENDADOR is not None and _THREAD_AGENDADOR.is_alive():
            return

        def _loop_relogio():
            while True:
                try:
                    _verificar_turno_pendente_agora()
                except Exception:
                    pass
                time.sleep(20)

        _THREAD_AGENDADOR = threading.Thread(target=_loop_relogio, daemon=True)
        _THREAD_AGENDADOR.start()


def formatar_moeda_br(valor):
    try:
        v = float(valor or 0)
    except Exception:
        v = 0.0
    texto = f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {texto}"


def exibir_painel_financeiro_exclusivo():
    """
    Exibe o Painel Financeiro da Fibrawer protegido por senha exclusiva do Anderson.
    """
    st.markdown("---")
    with st.expander("🔒 Painel Financeiro Exclusivo — Diretoria Fibrawer (Acesso Restrito)", expanded=False):
        if not SENHA_FINANCEIRO:
            st.warning("⚠️ Configure a variável `SENHA_FINANCEIRO` no seu arquivo `.streamlit/secrets.toml` para habilitar o painel.")
            return

        if "auth_financeiro_ok" not in st.session_state:
            st.session_state["auth_financeiro_ok"] = False

        if not st.session_state["auth_financeiro_ok"]:
            col_s1, col_s2 = st.columns([3, 1])
            senha_dig = col_s1.text_input(
                "🔑 Digite a senha exclusiva para visualizar o fechamento financeiro:",
                type="password",
                key="input_senha_fin_anderson"
            )
            if col_s2.button("🔓 Desbloquear", use_container_width=True):
                if senha_dig == str(SENHA_FINANCEIRO):
                    st.session_state["auth_financeiro_ok"] = True
                    st.rerun()
                else:
                    st.error("❌ Senha incorreta. Acesso negado.")
            return

        col_top1, col_top2, col_top3 = st.columns([2, 1, 1])
        aba_padrao = obter_nome_aba_ciclo_atual()
        nome_aba_consulta = col_top1.text_input("📅 Ciclo de Referência:", value=aba_padrao, key="aba_ciclo_fin")

        atualizar_agora = col_top2.button("🔄 Atualizar Valores", type="primary", use_container_width=True)
        if col_top3.button("🔒 Bloquear Painel", use_container_width=True):
            st.session_state["auth_financeiro_ok"] = False
            st.session_state.pop("dados_fin_cache", None)
            st.rerun()

        if atualizar_agora or "dados_fin_cache" not in st.session_state:
            with st.spinner("📊 Calculando faturamento Fibrawer (Assertec) vs. repasse dos técnicos..."):
                try:
                    payload = {
                        "acao": "resumo_financeiro",
                        "nome_aba": nome_aba_consulta,
                        "url_assertec": ASSERTEC_PLANILHA_URL
                    }
                    dados_fin = _post_google_script_com_retry(payload, timeout_seg=120)
                    if dados_fin.get("ok"):
                        st.session_state["dados_fin_cache"] = dados_fin
                    else:
                        st.error(f"Erro ao carregar painel financeiro: {dados_fin.get('msg')}")
                        return
                except Exception as e:
                    st.error(f"Falha ao conectar com as planilhas: {e}")
                    return

        dados_fin = st.session_state.get("dados_fin_cache", {})
        if not dados_fin:
            return

        info_ass = dados_fin.get("assertec", {})
        lista_tecs = dados_fin.get("tecnicos", [])

        bruto_fibrawer = float(info_ass.get("bruto_total", 0))
        qtd_ass = int(info_ass.get("qtd_servicos", 0))
        bruto_amigo = float(info_ass.get("bruto_amigo", 0))
        bruto_excl_ass = float(info_ass.get("bruto_exclusivo_assertec", 0))

        repasse_terceiros_total = 0.0
        repasse_terceiros_autorizado = 0.0
        producao_anderson_planilha = 0.0
        total_todas_planilhas = 0.0

        linhas_tabela = []
        for t in lista_tecs:
            nome_t = str(t.get("tecnico", "")).upper()
            tot_plan = float(t.get("total_planilha", 0))
            tot_aut = float(t.get("total_autorizado", 0))
            tot_pend = float(t.get("total_pendente", 0))
            tot_rep = float(t.get("total_reprovado", 0))
            bruto_t = float(t.get("bruto_assertec", 0))
            margem_t = bruto_t - tot_plan

            total_todas_planilhas += tot_plan
            if "ANDERSON" in nome_t:
                producao_anderson_planilha += tot_plan
            else:
                repasse_terceiros_total += tot_plan
                repasse_terceiros_autorizado += tot_aut

            linhas_tabela.append({
                "Técnico": nome_t,
                "Serviços": int(t.get("qtd_servicos", 0)),
                "Autorizados": int(t.get("qtd_autorizados", 0)),
                "Cabo (m)": int(t.get("metragem_total", 0)),
                "Gerado na Assertec (Bruto)": formatar_moeda_br(bruto_t),
                "Total Planilha Técnico": formatar_moeda_br(tot_plan),
                "🟢 Pagamento Autorizado": formatar_moeda_br(tot_aut),
                "🟡 Pendente Aprovação": formatar_moeda_br(tot_pend),
                "🔴 Reprovado": formatar_moeda_br(tot_rep),
                "💵 Sobras p/ Fibrawer": formatar_moeda_br(margem_t),
            })

        liquido_fibrawer_mais_anderson = bruto_fibrawer - repasse_terceiros_total
        margem_pura_pj = bruto_fibrawer - total_todas_planilhas

        st.markdown("#### 💰 Resumo Geral do Ciclo (Fibrawer)")
        k1, k2, k3, k4 = st.columns(4)
        k1.metric(
            "Faturamento Bruto (Assertec)",
            formatar_moeda_br(bruto_fibrawer),
            f"{qtd_ass} serviços na Relatorio"
        )
        k2.metric(
            "A Pagar (Márcio + Bruno)",
            formatar_moeda_br(repasse_terceiros_total),
            f"Já Autorizado: {formatar_moeda_br(repasse_terceiros_autorizado)}"
        )
        k3.metric(
            "Saldo Líquido (Fibrawer + Anderson)",
            formatar_moeda_br(liquido_fibrawer_mais_anderson),
            f"Sua Planilha: {formatar_moeda_br(producao_anderson_planilha)}"
        )
        k4.metric(
            "Margem Pura Empresa (PJ)",
            formatar_moeda_br(margem_pura_pj),
            f"Amigo: {formatar_moeda_br(bruto_amigo)} | Ass: {formatar_moeda_br(bruto_excl_ass)}"
        )

        st.markdown("#### 👷 Detalhamento por Técnico (Faturamento Gerado vs. Repasse vs. Margem)")
        if linhas_tabela:
            st.dataframe(pd.DataFrame(linhas_tabela), use_container_width=True, hide_index=True)