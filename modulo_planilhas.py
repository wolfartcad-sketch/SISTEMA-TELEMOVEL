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
_AGENDADOR_INICIADO = False
_LOCK_AGENDADOR = threading.Lock()


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


def obter_nome_aba_ciclo_atual(dt=None):
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
    Lê a aba 'Retorno Formulario' da planilha Assertec com dtype=str (para nunca adicionar '.0' ao protocolo)
    e cria dois mapas de validação: por Protocolo e por Nome do Cliente.
    """
    mapa_assertec = {}
    mapa_clientes_assertec = {}
    erros_metragem = []

    if not ASSERTEC_PLANILHA_URL or "/d/" not in ASSERTEC_PLANILHA_URL:
        return mapa_assertec, mapa_clientes_assertec, erros_metragem

    try:
        base_url = ASSERTEC_PLANILHA_URL.split("/edit")[0].strip("/")
        aba_encoded = urllib.parse.quote("Retorno Formulario")
        csv_url = f"{base_url}/gviz/tq?tqx=out:csv&sheet={aba_encoded}"

        df = pd.read_csv(csv_url, dtype=str).fillna("")

        for idx, row in df.iterrows():
            proto = limpar_protocolo_planilha(row.get("Protocolo", ""))
            cli = str(row.get("Nome do cliente", "")).strip()
            cli_norm = normalizar_nome_cliente(cli)
            cid = str(row.get("Cidade", "")).strip()
            srv = str(row.get("Tipo de serviço", "")).strip()
            qtd_raw = str(row.get("Quantidade / Metragem / Km", "")).strip()

            if not proto and not cli_norm:
                continue

            problemas = []
            if not cli:
                problemas.append("Cliente vazio")
            if not cid:
                problemas.append("Cidade vazia")
            if not srv:
                problemas.append("Serviço vazio")

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

    except Exception as e:
        st.warning(f"⚠️ Não foi possível ler a planilha da Assertec: {e}")

    return mapa_assertec, mapa_clientes_assertec, erros_metragem


def sincronizar_status_e_auditar_planilhas(mapa_amigo=None, silencioso=False):
    if mapa_amigo is None:
        from modulo_robos import ULTIMOS_STATUS_AMIGO
        mapa_amigo = ULTIMOS_STATUS_AMIGO

    mapa_assertec, mapa_clientes_assertec, erros_metragem = ler_dados_auditoria_assertec()
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
        "mapa_amigo": mapa_amigo or {},
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

        atualizados = dados.get("atualizados", [])
        autorizados = dados.get("autorizados", [])
        nao_aprovados_amigo = dados.get("nao_aprovados_amigo", [])
        pendentes_assertec = dados.get("pendentes_assertec", [])

        if atualizados:
            st.success(f"🔄 {len(atualizados)} serviço(s) tiveram o status atualizado nas planilhas dos técnicos!")
            if not silencioso:
                st.dataframe(atualizados, use_container_width=True)

        if not silencioso:
            c1, c2, c3 = st.columns(3)
            c1.metric("🟢 Pagamento Autorizado", len(autorizados))
            c2.metric("🟡/🔴 Pendentes ou Reprovados (Amigo)", len(nao_aprovados_amigo))
            c3.metric("⚠️ Pendências Planilha Assertec", len(pendentes_assertec))

            if autorizados:
                with st.expander(f"🟢 Ver Serviços com PAGAMENTO AUTORIZADO ({len(autorizados)})", expanded=True):
                    st.dataframe(autorizados, use_container_width=True)

            if nao_aprovados_amigo:
                with st.expander(f"🔴 Ver Serviços NÃO Aprovados / Aguardando no Sistema Amigo ({len(nao_aprovados_amigo)})", expanded=True):
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


def obter_agora_brasil():
    """Retorna a data e hora atual no fuso horário de Brasília (UTC-3)."""
    fuso_br = timezone(timedelta(hours=-3))
    return datetime.now(fuso_br)


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
    """Roda a leitura do Portal Amigo + Planilha Assertec em 2º plano às 12h e 21h sem precisar de clique."""
    try:
        from modulo_robos import buscar_status_portal_amigo_sob_demanda
        _, _, mapa_amigo = buscar_status_portal_amigo_sob_demanda()
        mapa_assertec, mapa_clientes_assertec, _ = ler_dados_auditoria_assertec()
        nome_aba = obter_nome_aba_ciclo_atual()

        payload = {
            "acao": "atualizar_status",
            "nome_aba": nome_aba,
            "mapa_amigo": mapa_amigo or {},
            "mapa_assertec": mapa_assertec,
            "mapa_clientes_assertec": mapa_clientes_assertec
        }

        resp = requests.post(GOOGLE_SCRIPT_PLANILHAS_URL, json=payload, timeout=30)
        dados = resp.json()
        if dados.get("ok"):
            qtd_att = len(dados.get("atualizados", []))
            qtd_aut = len(dados.get("autorizados", []))
            resumo = f"{qtd_att} status atualizado(s) | {qtd_aut} com PAGAMENTO AUTORIZADO"
            salvar_log_auto_sync(chave_turno, resumo)
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
                if hora == 12 and minuto <= 20:
                    turno_alvo = f"{data_hoje}_12H"
                elif hora == 21 and minuto <= 20:
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

        # Separa os valores do Anderson (proprietário) e dos técnicos externos (Márcio + Bruno)
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

        # Saldo que fica para a Fibrawer + Anderson (descontando pagamentos de Márcio e Bruno)
        liquido_fibrawer_mais_anderson = bruto_fibrawer - repasse_terceiros_total
        # Margem pura da PJ (se descontar também a tabela do próprio Anderson)
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