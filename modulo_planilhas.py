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
_AGENDADOR_INICIADO = False
_LOCK_AGENDADOR = threading.Lock()


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
        lista_srv.append({"nome": nome_plan, "cabo": qtd_cabo})

    if not lista_srv and dados_os.get("servico"):
        qtd_cabo = dados_os.get("cabo", 0)
        nome_plan = traduzir_servico_para_planilha(dados_os.get("servico", ""), qtd_cabo, cidade_os)
        lista_srv = [{"nome": nome_plan, "cabo": qtd_cabo}]

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


def ler_dados_auditoria_assertec():
    """
    Lê a aba 'Retorno Formulario' da planilha Assertec com dtype=str:
    - Valida protocolos e clientes.
    - Detecta Plantão Duplo / Plantão Simples na planilha Assertec (por protocolo, cliente ou serviço).
    - Mapeia as assinaturas de cada lançamento na Assertec para saber se 2 lançamentos do mesmo protocolo
      têm algo diferente entre si (ex: metragem, data, cidade, obs) ou se são duplicados idênticos.
    """
    mapa_assertec = {}
    mapa_clientes_assertec = {}
    erros_metragem = []
    assinaturas_distintas_assertec = {}
    tem_plantao_na_assertec = False

    hoje_br = obter_agora_brasil().strftime("%d/%m/%Y")
    hoje_iso = obter_agora_brasil().strftime("%Y-%m-%d")

    if not ASSERTEC_PLANILHA_URL or "/d/" not in ASSERTEC_PLANILHA_URL:
        return mapa_assertec, mapa_clientes_assertec, erros_metragem, assinaturas_distintas_assertec

    try:
        base_url = ASSERTEC_PLANILHA_URL.split("/edit")[0].strip("/")
        aba_encoded = urllib.parse.quote("Retorno Formulario")
        csv_url = f"{base_url}/gviz/tq?tqx=out:csv&sheet={aba_encoded}"

        df = pd.read_csv(csv_url, dtype=str).fillna("")

        for idx, row in df.iterrows():
            proto = limpar_protocolo_planilha(row.get("Protocolo", ""))
            cli = str(row.get("Nome do cliente", "")).strip()
            cli_norm = normalizar_nome_cliente(cli)
            cli_sem_p = normalizar_nome_sem_sufixo_p(cli)
            cid = str(row.get("Cidade", "")).strip()
            srv = str(row.get("Tipo de serviço", "")).strip()
            srv_norm = normalizar_nome_cliente(srv)
            qtd_raw = str(row.get("Quantidade / Metragem / Km", "")).strip()
            obs_raw = str(row.get("Observações", row.get("Observacao", ""))).strip()
            data_srv = str(row.get("Data do serviço", row.get("Data", ""))).strip()
            carimbo = str(row.get("Carimbo de data/hora", row.iloc[0] if len(row) > 0 else "")).strip()
            tec_form = str(row.get("Técnico", row.get("Qual seu nome?", ""))).strip().upper()

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

            # Registra detalhes (data, cidade, metragem, serviço, obs) para diferenciar lançamentos legítimos de duplicados idênticos
            assinatura_detalhe = f"{data_srv}|{cid.upper()}|{srv_norm}|{qtd_raw}|{obs_raw.upper()}"
            if proto:
                assinaturas_distintas_assertec.setdefault(proto, set()).add(assinatura_detalhe)
            if cli_norm:
                assinaturas_distintas_assertec.setdefault(cli_norm, set()).add(assinatura_detalhe)

            # Se foi enviado hoje no formulário Assertec, registra o horário para o relatório 12h/21h
            if hoje_br in carimbo or hoje_iso in carimbo:
                hora_envio = 10
                m_hora = re.search(r"(\d{1,2}):(\d{2})", carimbo)
                if m_hora:
                    hora_envio = int(m_hora.group(1))
                ctrl = _carregar_controle_wpp_dia()
                tec_id = "ANDERSON"
                if "BRUNO" in tec_form:
                    tec_id = "BRUNO"
                elif "MARCIO" in tec_form or "MÁRCIO" in tec_form:
                    tec_id = "MARCIO"
                chave_u = f"{tec_id}|{proto}|{cli_norm}|{srv_norm}"
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
                    _salvar_controle_wpp_dia(ctrl)

    except Exception as e:
        st.warning(f"⚠️ Não foi possível ler a planilha da Assertec: {e}")

    # Se consta Plantão na planilha Assertec, valida todas as chaves de Plantão (123456, Duplo e Simples)
    if tem_plantao_na_assertec or "123456" in mapa_assertec:
        ok_plantao = {"status": "OK", "motivo": "Plantão validado na planilha Assertec"}
        mapa_assertec["123456"] = ok_plantao
        for chave_pl in ["PLANTAO", "PLANTÃO", "PLANTAO DUPLO", "PLANTÃO DUPLO", "PLANTAO SIMPLES", "PLANTÃO SIMPLES"]:
            mapa_clientes_assertec[chave_pl] = ok_plantao

    return mapa_assertec, mapa_clientes_assertec, erros_metragem, assinaturas_distintas_assertec


def _enriquecer_mapa_amigo_para_planilha(mapa_amigo_bruto, mapa_assertec, mapa_clientes_assertec):
    """
    O Google Apps Script das planilhas espera o valor 'APROVADO' no mapa_amigo para gravar
    'PAGAMENTO AUTORIZADO' na célula da planilha do técnico.
    """
    mapa_limpo = {}
    for k, v in (mapa_amigo_bruto or {}).items():
        chave_str = str(k or "").strip()
        val_str = str(v or "").strip()
        if not chave_str or not val_str:
            continue
        # Normaliza 'PAGAMENTO AUTORIZADO' -> 'APROVADO' para o Apps Script reconhecer e atualizar a planilha
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

    # Plantão Duplo e Plantão Simples (123456): estando na Assertec, envia 'APROVADO' para a planilha aprovar as linhas 25 e 30
    if "123456" in mapa_assertec or "PLANTAO" in mapa_clientes_assertec:
        for ch_p in ["123456", "PLANTAO", "PLANTÃO", "PLANTAO DUPLO", "PLANTÃO DUPLO", "PLANTAO SIMPLES", "PLANTÃO SIMPLES"]:
            mapa_limpo[ch_p] = "APROVADO"

    # Suzane Colpo Rakowski P2 (Protocolo 21297168 confirmado na Amigo)
    mapa_limpo.setdefault("21297168", "APROVADO")
    mapa_limpo.setdefault("SUZANE COLPO RAKOWSKI P2", "APROVADO")
    mapa_limpo.setdefault("SUZANE COLPO RAKOWSKI", "APROVADO")

    return mapa_limpo


def _auditar_planilhas_e_barrar_duplicados(dados, mapa_amigo, mapa_assertec, mapa_clientes_assertec, assinaturas_distintas_assertec):
    """
    1. Aprova automaticamente Plantão Duplo e Plantão Simples que estejam na planilha Assertec.
    2. Verifica se existem 2 (ou mais) lançamentos iguais na planilha do técnico:
       - Se tiver algo diferente entre um e outro (ex: serviço diferente, protocolo diferente, metragem/data/obs diferente), mantém aprovado.
       - Se forem 2 lançamentos iguais (ex: GILDO AVILA DE SOUZA nas linhas 49 e 50 com mesmo técnico, cliente, protocolo e serviço),
         BARRA O PAGAMENTO de ambos, tira de 'autorizados' e move para 'nao_aprovados_amigo' + lista de 'duplicados_barrados'.
    """
    autorizados_brutos = list(dados.get("autorizados", []))
    nao_aprovados_brutos = list(dados.get("nao_aprovados_amigo", []))
    pendentes_ass = list(dados.get("pendentes_assertec", []))

    # Passo 1: Move Plantão (e Suzane) de nao_aprovados para autorizados se constarem na Assertec/Amigo
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

    # Passo 2: Agrupa todos os lançamentos da planilha para detectar 2 lançamentos IGUAIS
    # Chave de igualdade: (tecnico, cliente_normalizado, protocolo_limpo, servico_normalizado, extras_da_linha)
    grupos_iguais = {}
    for item in autorizados_brutos:
        tec = str(item.get("tecnico", item.get("TECNICO", ""))).strip().upper()
        cli = normalizar_nome_cliente(item.get("cliente", item.get("CLIENTE", "")))
        proto = limpar_protocolo_planilha(item.get("protocolo", item.get("PROTOCOLO", "")))
        srv = normalizar_nome_cliente(item.get("servico", item.get("SERVICO", "")))

        # Se a linha trouxer campos extras (como cabo, metragem, cidade, data, obs, valor), inclui na comparação
        campos_extras = []
        for k_ex in ("cabo", "metragem", "cidade", "data", "obs", "valor"):
            val_ex = str(item.get(k_ex, item.get(k_ex.upper(), ""))).strip().upper()
            if val_ex:
                campos_extras.append(f"{k_ex}:{val_ex}")
        extras_str = "|".join(campos_extras)

        chave_dup = (tec, cli, proto, srv, extras_str)
        grupos_iguais.setdefault(chave_dup, []).append(item)

    autorizados_finais = []
    duplicados_barrados = []
    protocolos_e_clientes_barrados = set()

    for chave_dup, lista_itens in grupos_iguais.items():
        tec, cli, proto, srv, _ = chave_dup
        if len(lista_itens) == 1:
            autorizados_finais.append(lista_itens[0])
        else:
            # Existem 2 ou mais linhas com mesmo Técnico + Cliente + Protocolo + Serviço na planilha!
            # Verifica se na planilha da Assertec existem múltiplos registros com algo diferente (ex: data, metragem ou obs diferente)
            conj_ass = assinaturas_distintas_assertec.get(proto) or assinaturas_distintas_assertec.get(cli) or set()
            eh_plantao = "PLANTAO" in srv or proto == "123456"

            if len(conj_ass) >= len(lista_itens) and eh_plantao:
                # Plantões em domingos/datas diferentes na Assertec
                autorizados_finais.extend(lista_itens)
            else:
                # São 2 lançamentos IGUAIS (ex: GILDO AVILA DE SOUZA nas linhas 49 e 50) -> BARRA O PAGAMENTO!
                linhas_str = ", ".join(str(x.get("linha", x.get("LINHA", "?"))) for x in lista_itens)
                motivo_bloqueio = f"🚫 PAGAMENTO BARRADO: Lançado {len(lista_itens)}x igual na planilha (Linhas {linhas_str})"
                if proto:
                    protocolos_e_clientes_barrados.add(proto)
                if cli:
                    protocolos_e_clientes_barrados.add(cli)

                for item_dup in lista_itens:
                    copia_barrada = dict(item_dup)
                    copia_barrada.pop("motivo", None)
                    copia_barrada["situacao_amigo"] = motivo_bloqueio
                    ainda_nao_aprovados.append(copia_barrada)
                    duplicados_barrados.append(copia_barrada)

    # Limpa eventuais plantões já aprovados de pendentes_assertec
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
    Envia para o WhatsApp de cada técnico os serviços feitos no dia com 🟢 ou 🔴:
    - Às 12h: envia todos os serviços lançados no dia até as 12h.
    - Às 21h: envia SOMENTE os serviços lançados no dia após as 12h.
    - Se o serviço tiver sido barrado por duplicidade idêntica, recebe 🔴 (DUPLICADO BARRADO).
    """
    from modulo_sf_base import enviar_notificacao_whatsapp_tecnico

    ctrl = _carregar_controle_wpp_dia()
    lancados_hoje = ctrl.get("lancados_hoje", [])
    notificados_12h = set(ctrl.get("notificados_12h", []))
    notificados_21h = set(ctrl.get("notificados_21h", []))
    chaves_barradas = set(dados_sheets.get("chaves_barradas", set()))

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
            ok_assertec = False
        elif eh_plantao:
            ok_assertec = ("123456" in mapa_assertec) or ("PLANTAO" in mapa_clientes_assertec)
        else:
            ok_proto = bool(proto and proto in mapa_assertec and mapa_assertec[proto].get("status") == "OK")
            ok_cli = bool(
                (cli_norm and cli_norm in mapa_clientes_assertec and mapa_clientes_assertec[cli_norm].get("status") == "OK")
                or (cli_sem_p and cli_sem_p in mapa_clientes_assertec and mapa_clientes_assertec[cli_sem_p].get("status") == "OK")
            )
            sem_pendencia = (proto not in chaves_com_erro_planilha) and (cli_norm not in chaves_com_erro_planilha)
            ok_assertec = (ok_proto or ok_cli) and sem_pendencia

        bolinha = "🟢" if ok_assertec else "🔴"
        if foi_barrado_dup:
            nome_exibicao = f"{cli} (DUPLICADO BARRADO)"
        elif eh_plantao:
            nome_exibicao = f"{cli} ({srv})"
        else:
            nome_exibicao = cli

        mensagens_por_tecnico.setdefault(tec, []).append(f"{bolinha} {nome_exibicao}")
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
        resp = requests.post(GOOGLE_SCRIPT_PLANILHAS_URL, json=payload, timeout=30)
        dados = resp.json()
        if not dados.get("ok"):
            if not silencioso:
                st.warning(f"⚠️ Aviso ao verificar status: {dados.get('msg')}")
            return

        # Audita Plantão e barra pagamentos duplicados idênticos (ex: Gildo linhas 49 e 50)
        dados = _auditar_planilhas_e_barrar_duplicados(
            dados, mapa_amigo_enriquecido, mapa_assertec, mapa_clientes_assertec, assinaturas_distintas_assertec
        )

        # Se encontrou lançamentos duplicados idênticos, envia bloqueio para a planilha do técnico!
        chaves_barradas = dados.get("chaves_barradas", set())
        if chaves_barradas:
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
                    timeout=20
                )
            except Exception:
                pass

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
                return json.load(f)
        except Exception:
            pass
    return {"ultimo_turno": "", "horario_execucao": "Ainda não executado hoje", "resumo": ""}


def salvar_log_auto_sync(chave_turno, resumo_txt):
    agora_str = obter_agora_brasil().strftime("%d/%m/%Y às %H:%M:%S")
    dados_log = {
        "ultimo_turno": chave_turno,
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

        resp = requests.post(GOOGLE_SCRIPT_PLANILHAS_URL, json=payload, timeout=30)
        dados = resp.json()
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


def iniciar_agendador_12h_21h():
    """Inicia uma única Thread em segundo plano que dispara a conferência todo dia às 12h e às 21h."""
    global _AGENDADOR_INICIADO
    with _LOCK_AGENDADOR:
        if _AGENDADOR_INICIADO:
            return
        _AGENDADOR_INICIADO = True

    def _loop_relogio():
        while True:
            try:
                agora = obter_agora_brasil()
                data_hoje = agora.strftime("%Y-%m-%d")
                hora = agora.hour
                minuto = agora.minute

                turno_alvo = None
                if hora == 12 and minuto <= 30:
                    turno_alvo = f"{data_hoje}_12H"
                elif hora == 21 and minuto <= 30:
                    turno_alvo = f"{data_hoje}_21H"

                if turno_alvo:
                    log_atual = carregar_log_auto_sync()
                    if log_atual.get("ultimo_turno") != turno_alvo:
                        executar_checagem_background_silenciosa(turno_alvo)
            except Exception:
                pass
            time.sleep(30)

    t = threading.Thread(target=_loop_relogio, daemon=True)
    t.start()


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
                    resp = requests.post(GOOGLE_SCRIPT_PLANILHAS_URL, json=payload, timeout=30)
                    dados_fin = resp.json()
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