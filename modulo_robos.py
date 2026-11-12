import os
import re
import sys
import json
import time
import threading
import subprocess
from datetime import datetime

from config import (
    TECPAR_PORTAL_URL,
    TECPAR_USER,
    TECPAR_PASS,
    ARQUIVO_PRINT_PORTAL,
    SALESFORCE_URL,
    SALESFORCE_USER,
    SALESFORCE_USER_ALT,
    SALESFORCE_PASS,
    PASTA_SESSAO_SF,
    ARQUIVO_PRINT_SF,
    ARQUIVO_CACHE_SF,
    ASSERTEC_FORM_URL,
)
from modulo_assertec import (
    JUSTIFICATIVA_PADRAO_TECPAR,
    MSG_OBS_PADRAO_AMIGO,
    lista_cidades_assertec,
    EQUIPAMENTOS_LISTA,
    servico_vai_para_amigo,
    servico_exige_equipamentos,
    limpar_protocolo_str, mapear_para_assertec,
)

# Controle global de instalação do navegador (seguro para rodar em 2º plano sem travar threads)
_CHROMIUM_VERIFICADO = False
_LOCK_ROBO_AMIGO = threading.Lock()

def garantir_instalacao_chromium():
    global _CHROMIUM_VERIFICADO
    if _CHROMIUM_VERIFICADO:
        return True
    try:
        subprocess.run(
            [sys.executable, "-m", "playwright", "install", "chromium"],
            check=False,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )
        _CHROMIUM_VERIFICADO = True
        return True
    except Exception:
        return False

def iniciar_browser_playwright(p):
    opcoes_args = ["--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu"]
    for caminho_sis in ["/usr/bin/chromium", "/usr/bin/chromium-browser", "/usr/bin/google-chrome"]:
        if os.path.exists(caminho_sis):
            try:
                return p.chromium.launch(headless=True, executable_path=caminho_sis, args=opcoes_args)
            except Exception:
                pass
    try:
        return p.chromium.launch(headless=True, args=opcoes_args)
    except Exception:
        subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], check=False)
        return p.chromium.launch(headless=True, args=opcoes_args)

# ==============================================================================
# ROBÔ 1: SALESFORCE (FIELD SERVICE)
# ==============================================================================
def carregar_agenda_sf_cache():
    if os.path.exists(ARQUIVO_CACHE_SF):
        try:
            with open(ARQUIVO_CACHE_SF, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return []

def salvar_agenda_sf_cache(lista):
    try:
        with open(ARQUIVO_CACHE_SF, "w", encoding="utf-8") as f:
            json.dump(lista, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def sincronizar_agenda_salesforce(codigo_mfa=""):
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        return False, "Biblioteca Playwright não instalada neste ambiente.", []

    garantir_instalacao_chromium()
    os.makedirs(PASTA_SESSAO_SF, exist_ok=True)

    try:
        with sync_playwright() as p:
            opcoes_args = ["--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu"]
            caminho_exec = None
            for c_sis in ["/usr/bin/chromium", "/usr/bin/chromium-browser", "/usr/bin/google-chrome"]:
                if os.path.exists(c_sis):
                    caminho_exec = c_sis
                    break

            kwargs_ctx = {
                "user_data_dir": PASTA_SESSAO_SF,
                "headless": True,
                "viewport": {"width": 1440, "height": 1080},
                "args": opcoes_args
            }
            if caminho_exec:
                kwargs_ctx["executable_path"] = caminho_exec

            context = p.chromium.launch_persistent_context(**kwargs_ctx)
            page = context.pages[0] if context.pages else context.new_page()
            page.goto(SALESFORCE_URL, timeout=60000)
            page.wait_for_timeout(3000)

            inp_user = page.locator("input#username, input[name='username']").first
            if inp_user.is_visible(timeout=5000):
                inp_user.fill(SALESFORCE_USER)
                inp_pass = page.locator("input#password, input[name='pw']").first
                inp_pass.fill(SALESFORCE_PASS)
                page.locator("input#Login, input[name='Login'], button#Login").first.click()
                page.wait_for_timeout(3500)

                erro_login = page.locator("#error, .loginError").first
                if erro_login.is_visible(timeout=1500):
                    inp_user.fill(SALESFORCE_USER_ALT)
                    inp_pass.fill(SALESFORCE_PASS)
                    page.locator("input#Login, input[name='Login'], button#Login").first.click()
                    page.wait_for_timeout(3500)

            inp_mfa = page.locator("input#tc, input#emc, input[name='tc'], input[name='emc'], input[autocomplete='one-time-code']").first
            if inp_mfa.is_visible(timeout=4000):
                if not str(codigo_mfa).strip():
                    page.screenshot(path=ARQUIVO_PRINT_SF, full_page=True)
                    context.close()
                    return False, "🔑 O Salesforce pediu o código do Google Authenticator. Digite os 6 dígitos no campo ao lado e clique em Sincronizar novamente.", []
                inp_mfa.fill(str(codigo_mfa).strip())
                btn_verif = page.locator("input#save, input[name='save'], button[type='submit']").first
                btn_verif.click()
                page.wait_for_timeout(5000)

            page.wait_for_timeout(8000)
            page.screenshot(path=ARQUIVO_PRINT_SF, full_page=True)

            js_extract = """
            () => {
                const resultados = [];
                const linhas = Array.from(document.querySelectorAll('tr, .ServiceAppointment, .listItem, [role="row"]'));
                for (const el of linhas) {
                    const txt = (el.innerText || '').trim();
                    if (!txt || txt.length < 15) continue;
                    if (txt.includes('SA-') || /\\d{6,}/.test(txt)) {
                        resultados.push(txt.replace(/\\n+/g, ' | '));
                    }
                }
                return resultados.slice(0, 50);
            }
            """
            brutos = []
            for frame in page.frames:
                try:
                    res_f = frame.evaluate(js_extract)
                    if res_f:
                        brutos.extend(res_f)
                except Exception:
                    pass

            context.close()

            lista_clientes = []
            vistos = set()
            for idx_b, linha_str in enumerate(brutos):
                if linha_str in vistos:
                    continue
                vistos.add(linha_str)
                partes = [p.strip() for p in linha_str.split("|") if p.strip()]
                proto_match = re.search(r"(SA-\d+|\b\d{6,10}\b)", linha_str)
                tel_match = re.search(r"(\(?\d{2}\)?\s*9?\d{4}[-\s]?\d{4})", linha_str)
                venc_match = re.search(r"(\d{2}/\d{2}/\d{4}(?:\s+\d{2}:\d{2})?)", linha_str)

                cid_achada = "N/D"
                for c_item in lista_cidades_assertec:
                    if c_item.upper() in linha_str.upper():
                        cid_achada = c_item
                        break

                lista_clientes.append({
                    "id": f"sf_{idx_b}",
                    "vencimento": venc_match.group(1) if venc_match else "Ver linha",
                    "cliente": partes[1] if len(partes) > 1 else partes[0][:40],
                    "telefone": tel_match.group(1) if tel_match else "",
                    "protocolo": proto_match.group(1) if proto_match else "",
                    "tipo_protocolo": "Manutenção / Instalação",
                    "etiqueta": "",
                    "cidade": cid_achada,
                    "endereco": "",
                    "resumo_bruto": linha_str
                })

            if lista_clientes:
                salvar_agenda_sf_cache(lista_clientes)
                return True, f"✅ {len(lista_clientes)} cliente(s) sincronizado(s) do Salesforce!", lista_clientes
            else:
                return False, "⚠️ Conectado ao Salesforce! Confira o print abaixo para ajustarmos as colunas exatas da sua tela.", []
    except Exception as e:
        return False, f"Erro ao sincronizar Salesforce: {e}", []

# ==============================================================================
# ROBÔ 2: PORTAL BRASIL TECPAR (AMIGO / AVATO) - RÁPIDO E SEM FALSO ERRO
# ==============================================================================
def obter_frame_interno_gas(page, timeout_ms=25000):
    inicio = time.time()
    while (time.time() - inicio) * 1000 < timeout_ms:
        f_user = page.frame(name="userHtmlFrame")
        if f_user:
            try:
                if f_user.locator("input, button").count() > 0:
                    return f_user
            except Exception:
                pass

        for frame in reversed(page.frames):
            try:
                if frame.locator("input, button").count() >= 2:
                    return frame
            except Exception:
                pass
        page.wait_for_timeout(150)
    return page.frames[-1]

def selecionar_opcao_select_flexivel(select_locator, texto_alvo):
    def norm(s):
        return re.sub(r"\s+", "", str(s or "").upper())

    alvo_n = norm(texto_alvo)
    if not alvo_n:
        return False

    opcoes = select_locator.locator("option")
    qtd = opcoes.count()
    for idx in range(qtd):
        txt_op = opcoes.nth(idx).inner_text() or opcoes.nth(idx).text_content() or ""
        val_op = opcoes.nth(idx).get_attribute("value") or ""
        tn = norm(txt_op)
        if not tn or "SELECIONE" in tn or tn == "TODOS" or tn == "TODAS":
            continue
        if tn == alvo_n or alvo_n in tn or tn in alvo_n:
            if val_op:
                select_locator.select_option(value=val_op)
            else:
                select_locator.select_option(index=idx)
            select_locator.dispatch_event("change")
            return True
    return False

def marcar_input_por_rotulo_isolado(target_frame, rotulo_exato, attr_marca):
    js_mark = """
    ([rotulo, marca]) => {
        const norm = (s) => (s || '').normalize('NFD').replace(/[\\u0300-\\u036f]/g, '').toUpperCase().trim();
        const alvo = norm(rotulo);
        const antigos = document.querySelectorAll(`[data-robo-alvo="${marca}"]`);
        antigos.forEach(el => el.removeAttribute('data-robo-alvo'));

        const todos = Array.from(document.querySelectorAll('label, span, div, p'));
        for (const el of todos) {
            if (el.children.length > 0) continue;
            const t = norm(el.innerText || el.textContent);
            if (t === alvo || (alvo.length >= 10 && t.includes(alvo))) {
                let p = el.parentElement;
                for (let i = 0; i < 3 && p; i++) {
                    const inps = p.querySelectorAll('input:not([type="hidden"]):not([disabled])');
                    if (inps.length === 1) {
                        inps[0].setAttribute('data-robo-alvo', marca);
                        return true;
                    }
                    if (inps.length > 1) break;
                    p = p.parentElement;
                }
            }
        }
        return false;
    }
    """
    try:
        return target_frame.evaluate(js_mark, [rotulo_exato, attr_marca])
    except Exception:
        return False

def preencher_justificativas_visiveis(target_frame, texto_justificativa=JUSTIFICATIVA_PADRAO_TECPAR):
    js_just = """
    (msg) => {
        const norm = (s) => (s || '').normalize('NFD').replace(/[\\u0300-\\u036f]/g, '').toUpperCase().trim();
        const setNativeValue = (el, val) => {
            const proto = el.tagName === 'TEXTAREA' ? window.HTMLTextAreaElement.prototype : window.HTMLInputElement.prototype;
            const setter = Object.getOwnPropertyDescriptor(proto, 'value')?.set;
            if (setter) {
                setter.call(el, val);
            } else {
                el.value = val;
            }
            el.dispatchEvent(new Event('focus', { bubbles: true }));
            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true, key: 'e' }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
            el.dispatchEvent(new Event('blur', { bubbles: true }));
        };

        let qtdPreenchida = 0;
        const campos = Array.from(document.querySelectorAll('input:not([type="hidden"]):not([type="number"]):not([type="date"]):not([type="password"]):not([type="checkbox"]):not([type="radio"]), textarea'));

        for (const el of campos) {
            if (el.disabled || el.readOnly) continue;
            const r = el.getBoundingClientRect();
            if (r.width <= 0 || r.height <= 0) continue;

            const ph = norm(el.getAttribute('placeholder') || '');
            const idName = norm((el.id || '') + ' ' + (el.name || '') + ' ' + (el.className || ''));

            let textoBloco = '';
            let p = el.parentElement;
            for (let i = 0; i < 2 && p; i++) {
                const t = norm(p.innerText || p.textContent);
                if (t.length < 400) {
                    textoBloco += ' ' + t;
                }
                p = p.parentElement;
            }

            const ehCampoJustificativa = (
                ph.includes('JUSTIFIQUE') ||
                ph.includes('CONTRATOU') ||
                ph.includes('PONTOS EXTRAS') ||
                ph.includes('MAIS DE UM') ||
                idName.includes('JUST') ||
                textoBloco.includes('JUSTIFICATIVA OBRIGATORIA') ||
                textoBloco.includes('JUSTIFIQUE POR QUE')
            );

            if (ehCampoJustificativa) {
                el.setAttribute('data-robo-just', 'sim');
                if (!el.value || el.value.trim().length < 10) {
                    setNativeValue(el, msg);
                    qtdPreenchida++;
                }
            }
        }
        return qtdPreenchida;
    }
    """
    try:
        target_frame.evaluate(js_just, texto_justificativa)
    except Exception:
        pass

    try:
        locs = target_frame.locator("[data-robo-just='sim']:visible")
        for i in range(locs.count()):
            campo = locs.nth(i)
            val_atual = campo.input_value() or ""
            if len(val_atual.strip()) < 10:
                campo.click()
                campo.fill(texto_justificativa)
                campo.dispatch_event("input")
                campo.dispatch_event("keyup")
                campo.dispatch_event("change")
                campo.press("Tab")
    except Exception:
        pass

def executar_lancamento_portal_tecpar(dados_os):
    servicos_amigo = [s for s in dados_os.get("servicos", []) if servico_vai_para_amigo(s["nome"])]
    if not servicos_amigo:
        return True, "Serviço exclusivo Assertec (não requer lançamento no portal Amigo)."

    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        return False, "Biblioteca Playwright não instalada neste ambiente."

    garantir_instalacao_chromium()

    with _LOCK_ROBO_AMIGO:
        try:
            with sync_playwright() as p:
                browser = iniciar_browser_playwright(p)
                context = browser.new_context(viewport={"width": 1366, "height": 2200})
                page = context.new_page()
                page.goto(TECPAR_PORTAL_URL, timeout=60000, wait_until="domcontentloaded")

                target_frame = obter_frame_interno_gas(page, timeout_ms=20000)

                # 1. LOGIN RÁPIDO
                campo_pass_vis = target_frame.locator("input[type='password']:visible").first
                try:
                    campo_pass_vis.wait_for(state="visible", timeout=5000)
                    campo_user_vis = target_frame.locator(
                        "input:visible:not([type='password']):not([type='hidden']):not([type='checkbox']):not([type='radio'])"
                    ).first
                    campo_user_vis.fill(TECPAR_USER)
                    campo_pass_vis.fill(TECPAR_PASS)
                    target_frame.locator("button:visible", has_text=re.compile(r"^\s*Entrar\s*$", re.I)).first.click()
                except Exception:
                    pass

                btn_lancar_ativ = None
                inicio_painel = time.time()
                while (time.time() - inicio_painel) < 22:
                    target_frame = obter_frame_interno_gas(page, timeout_ms=2500)
                    cand = target_frame.locator("button:visible", has_text=re.compile("Lançar atividade", re.I)).first
                    try:
                        if cand.is_visible():
                            btn_lancar_ativ = cand
                            break
                    except Exception:
                        pass
                    page.wait_for_timeout(150)

                if btn_lancar_ativ is None:
                    page.screenshot(path=ARQUIVO_PRINT_PORTAL, full_page=True)
                    browser.close()
                    return False, "❌ O portal não abriu a tela de lançamento após o login."

                page.wait_for_timeout(150)

                # 2. PREENCHIMENTO ÁGIL DOS CAMPOS PRINCIPAIS
                try:
                    inp_data_ativ = target_frame.locator("input[type='date']:visible").last
                    if inp_data_ativ.is_visible() and not inp_data_ativ.input_value():
                        inp_data_ativ.fill(datetime.now().strftime("%Y-%m-%d"))
                        inp_data_ativ.dispatch_event("change")
                except Exception:
                    pass

                proto_str = str(dados_os["protocolo"]).strip()
                loc_proto = target_frame.locator("input[placeholder*='12554094']:visible").first
                if not loc_proto.is_visible():
                    marcar_input_por_rotulo_isolado(target_frame, "PROTOCOLO", "proto")
                    loc_proto = target_frame.locator("input[data-robo-alvo='proto']:visible").first
                loc_proto.click()
                loc_proto.fill(proto_str)
                loc_proto.press("Tab")
                page.wait_for_timeout(80)

                mercado_escolhido = dados_os.get("mercado", "AMIGO - VAREJO")
                selects_vis = target_frame.locator("select:visible")
                for i in range(selects_vis.count()):
                    sel = selects_vis.nth(i)
                    txt_opcoes = sel.text_content() or ""
                    if "AMIGO - VAREJO" in txt_opcoes and "Todos" not in txt_opcoes:
                        selecionar_opcao_select_flexivel(sel, mercado_escolhido)
                        break
                page.wait_for_timeout(120)

                loc_cli = target_frame.locator("input[placeholder*='NOME DO CLIENTE' i]:visible").first
                if not loc_cli.is_visible():
                    marcar_input_por_rotulo_isolado(target_frame, "NOME DO CLIENTE", "cliente")
                    loc_cli = target_frame.locator("input[data-robo-alvo='cliente']:visible").first
                loc_cli.click()
                loc_cli.fill(dados_os["cliente"].upper())
                loc_cli.press("Tab")
                page.wait_for_timeout(80)

                preencheu_cid = False
                if marcar_input_por_rotulo_isolado(target_frame, "CIDADE", "cidade"):
                    loc_cid = target_frame.locator("input[data-robo-alvo='cidade']:visible").first
                    if loc_cid.is_visible():
                        loc_cid.click()
                        loc_cid.fill(dados_os["cidade"].upper())
                        loc_cid.press("Tab")
                        preencheu_cid = True
                if not preencheu_cid:
                    loc_cid_fb = target_frame.locator("input[value='FELIZ']:visible").first
                    if loc_cid_fb.is_visible():
                        loc_cid_fb.click()
                        loc_cid_fb.fill(dados_os["cidade"].upper())
                        loc_cid_fb.press("Tab")
                page.wait_for_timeout(150)

                obs_base = dados_os.get("obs", "")
                obs_final = f"{obs_base} | {MSG_OBS_PADRAO_AMIGO}" if obs_base else MSG_OBS_PADRAO_AMIGO

                loc_obs = target_frame.locator("textarea:visible").first
                if loc_obs.is_visible():
                    loc_obs.fill(obs_final)
                    loc_obs.dispatch_event("change")

                # 3. ADIÇÃO DE SERVIÇOS OTIMIZADA
                for idx, srv in enumerate(servicos_amigo):
                    selects_vis = target_frame.locator("select:visible")
                    selecionou_srv = False
                    for i in range(selects_vis.count()):
                        sel = selects_vis.nth(i)
                        txt_opcoes = (sel.text_content() or "").upper()
                        if "ATIVAÇÃO" in txt_opcoes or "BATIMENTO" in txt_opcoes or "VISTORIA" in txt_opcoes:
                            selecionou_srv = selecionar_opcao_select_flexivel(sel, srv["nome"])
                            break

                    if not selecionou_srv:
                        page.screenshot(path=ARQUIVO_PRINT_PORTAL, full_page=True)
                        browser.close()
                        return False, f"❌ Serviço '{srv['nome']}' não encontrado no portal."

                    page.wait_for_timeout(200)

                    lpu_val = str(max(1, int(srv.get("lpu", 1) or 1)))
                    inp_lpu = target_frame.locator("input[type='number']:visible:not([disabled])").first
                    try:
                        if inp_lpu.is_visible() and inp_lpu.is_enabled():
                            inp_lpu.fill(lpu_val)
                            inp_lpu.dispatch_event("input")
                            inp_lpu.dispatch_event("change")
                            inp_lpu.press("Tab")
                    except Exception:
                        pass

                    if int(lpu_val) > 1:
                        page.wait_for_timeout(150)
                        preencher_justificativas_visiveis(target_frame)

                    cabo_num = max(1, int(re.sub(r"\D", "", str(srv.get("cabo", 1))) or 1))
                    inp_cabo = target_frame.locator(
                        "input[placeholder*='Obrigatório' i]:visible, input[placeholder*='Se houver' i]:visible"
                    ).first
                    try:
                        if inp_cabo.is_visible() and inp_cabo.is_enabled():
                            inp_cabo.fill(str(cabo_num))
                            inp_cabo.dispatch_event("input")
                            inp_cabo.dispatch_event("change")
                            inp_cabo.press("Tab")
                    except Exception:
                        pass

                    page.wait_for_timeout(120)
                    preencher_justificativas_visiveis(target_frame)

                    if idx < len(servicos_amigo) - 1:
                        btn_add_srv = target_frame.locator("button:visible", has_text=re.compile("Adicionar serviço", re.I)).first
                        btn_add_srv.click()
                        page.wait_for_timeout(350)

                preencher_justificativas_visiveis(target_frame)

                # 4. DISPARO DO BOTÃO "Lançar atividade" E MODAL "Não, lançar agora"
                btn_lancar_ativ = target_frame.locator("button:visible", has_text=re.compile("Lançar atividade", re.I)).first
                try:
                    if btn_lancar_ativ.is_disabled():
                        target_frame.evaluate("""
                            () => {
                                const b = Array.from(document.querySelectorAll('button')).find(x => (x.innerText || '').includes('Lançar atividade'));
                                if (b) { b.disabled = false; b.click(); }
                            }
                        """)
                    else:
                        btn_lancar_ativ.click()
                except Exception:
                    pass

                btn_nao_lancar = target_frame.locator("button", has_text=re.compile("Não, lançar agora", re.I)).first
                try:
                    btn_nao_lancar.wait_for(state="visible", timeout=5000)
                    page.wait_for_timeout(100)
                    btn_nao_lancar.click()
                except Exception:
                    pass

                # 5. MODAL DE EQUIPAMENTOS (SEM ESPERA REPETIDA DE 5 SEGUNDOS)
                tem_ativacao = any(servico_exige_equipamentos(s["nome"]) for s in servicos_amigo)
                if tem_ativacao:
                    equipamentos_usados = dados_os.get("equipamentos", {})
                    btn_confirmar_eq = target_frame.locator("button", has_text=re.compile("Confirmar e lançar", re.I)).first
                    try:
                        btn_confirmar_eq.wait_for(state="visible", timeout=4500)
                    except Exception:
                        pass

                    for tentativa_eq in range(2):
                        if not btn_confirmar_eq.is_visible():
                            break

                        page.wait_for_timeout(150)
                        btns_nao_aplica = target_frame.locator("button:visible", has_text=re.compile(r"^\s*Não se aplica\s*$", re.I))
                        btns_novo = target_frame.locator("button:visible", has_text=re.compile(r"^\s*Novo\s*$", re.I))
                        btns_reman = target_frame.locator("button:visible", has_text=re.compile(r"^\s*Remanufaturado\s*$", re.I))

                        for idx_eq, nome_eq in enumerate(EQUIPAMENTOS_LISTA):
                            info_eq = equipamentos_usados.get(nome_eq, {"usou": False, "qtd": 0, "tipo": "Novo"})
                            if info_eq.get("usou") and int(info_eq.get("qtd", 0)) > 0:
                                try:
                                    if btns_novo.count() > idx_eq:
                                        btn_n = btns_novo.nth(idx_eq)
                                        inp_linha = btn_n.locator("xpath=ancestor::div[.//input][1]//input").first
                                        if inp_linha.is_visible():
                                            inp_linha.fill(str(info_eq["qtd"]))
                                            inp_linha.dispatch_event("input")
                                            inp_linha.dispatch_event("change")
                                    if str(info_eq.get("tipo", "Novo")).lower().startswith("reman"):
                                        if btns_reman.count() > idx_eq:
                                            btns_reman.nth(idx_eq).click()
                                    else:
                                        if btns_novo.count() > idx_eq:
                                            btns_novo.nth(idx_eq).click()
                                except Exception:
                                    pass
                            else:
                                try:
                                    if btns_nao_aplica.count() > idx_eq:
                                        btns_nao_aplica.nth(idx_eq).click()
                                except Exception:
                                    pass

                        page.wait_for_timeout(150)
                        preencher_justificativas_visiveis(target_frame)
                        page.wait_for_timeout(100)

                        if btn_confirmar_eq.is_disabled():
                            target_frame.evaluate("() => { const b = document.getElementById('eqOk'); if (b) { b.disabled = false; b.click(); } }")
                        else:
                            btn_confirmar_eq.click()

                        # Assim que o modal fechar, sai imediatamente sem perder 5 segundos
                        try:
                            btn_confirmar_eq.wait_for(state="hidden", timeout=3500)
                            break
                        except Exception:
                            pass

                # 6. CONFIRMAÇÃO INTELIGENTE (EVITA FALSO ERRO QUANDO O PORTAL DEMORA A ATUALIZAR A TABELA)
                gravou_com_sucesso = False
                inicio_verif = time.time()
                while (time.time() - inicio_verif) < 20:
                    try:
                        modal_aberto = target_frame.locator(
                            "button:visible", has_text=re.compile("Não, lançar agora|Confirmar e lançar", re.I)
                        ).count() > 0

                        if not modal_aberto:
                            texto_pagina = (target_frame.locator("body").inner_text(timeout=1500) or "").upper()
                            parte_atividades = texto_pagina.split("MINHAS ATIVIDADES")[-1] if "MINHAS ATIVIDADES" in texto_pagina else ""

                            # Critério 1: Apareceu na lista Minhas Atividades
                            na_lista = (proto_str in parte_atividades) and (
                                "AGUARDANDO APROVAÇÃO" in parte_atividades or "LAN-" in parte_atividades or dados_os["cliente"].upper()[:10] in parte_atividades
                            )

                            # Critério 2: O portal limpou o campo de protocolo após gravar com sucesso e não há mensagem de erro
                            val_campo_proto = ""
                            try:
                                val_campo_proto = (loc_proto.input_value(timeout=500) or "").strip()
                            except Exception:
                                pass
                            formulario_limpo = (val_campo_proto == "") and (proto_str in texto_pagina or "SUCESSO" in texto_pagina or "LAN-" in texto_pagina)

                            if na_lista or formulario_limpo:
                                gravou_com_sucesso = True
                                break
                    except Exception:
                        pass
                    page.wait_for_timeout(300)

                page.wait_for_timeout(300)
                target_frame.evaluate("window.scrollTo(0, 0)")
                page.screenshot(path=ARQUIVO_PRINT_PORTAL, full_page=True)
                browser.close()

                if gravou_com_sucesso:
                    return True, f"Protocolo {proto_str} ({dados_os['cliente']}) confirmado no Portal Amigo!"
                else:
                    return False, f"⚠️ Protocolo {proto_str} enviado, mas o portal demorou a atualizar a tabela. Confira o print abaixo."
        except Exception as e:
            return False, f"Erro no protocolo {dados_os.get('protocolo')}: {e}"

def executar_lancamento_amigo_seguro(dados_os):
    """
    Dispara o robô em uma Thread independente (daemon=False).
    Se o técnico fechar a aba do telemóvel durante o processo, a Thread continua
    executando em segundo plano até concluir 100% do lançamento no portal.
    Se o técnico permanecer na página, retorna o resultado normalmente para a tela.
    """
    resultado = {"ok": False, "msg": "Em processamento..."}

    def _worker():
        ok, msg = executar_lancamento_portal_tecpar(dados_os)
        resultado["ok"] = ok
        resultado["msg"] = msg

    t = threading.Thread(target=_worker, daemon=False)
    t.start()

    while t.is_alive():
        time.sleep(0.3)

    return resultado["ok"], resultado["msg"]

ULTIMOS_STATUS_AMIGO = {}

def extrair_status_minhas_atividades(target_frame):
    """
    Varre a tabela 'Minhas Atividades' do Portal Brasil TecPar (Amigo)
    e retorna um dicionário {protocolo: status} reconhecendo 'PAGAMENTO AUTORIZADO'.
    """
    js_status = """
    () => {
        const resultado = {};
        const norm = (s) => (s || '').toUpperCase().trim();
        const linhas = Array.from(document.querySelectorAll('tr, .card, .list-group-item'));

        for (const el of linhas) {
            const txt = norm(el.innerText || '');
            if (!txt || txt.length < 15 || txt.length > 2000) continue;

            let statusAchado = '';
            if (txt.includes('AGUARDANDO APROVA') || txt.includes('EM ANALISE') || txt.includes('EM ANÁLISE')) {
                statusAchado = 'AGUARDANDO APROVAÇÃO';
            } else if (txt.includes('REPROVAD') || txt.includes('RECUSAD') || txt.includes('REJEITAD') || txt.includes('CANCELAD')) {
                statusAchado = 'REPROVADO';
            } else if (txt.includes('PAGAMENTO AUTORIZADO') || txt.includes('AUTORIZADO') || txt.includes('APROVAD')) {
                statusAchado = 'APROVADO';
            }

            if (!statusAchado) continue;

            // 1º Captura explícita no padrão "Protocolo 180573"
            const matchExplicito = txt.match(/PROTOCOLO\\s*:?\\s*(\\d{4,14})/g);
            if (matchExplicito) {
                for (const m of matchExplicito) {
                    const num = m.replace(/\\D/g, '');
                    if (num) resultado[num] = statusAchado;
                }
            }

            // 2º Captura números de protocolo de 5 a 12 dígitos na linha
            const todosNumeros = txt.match(/\\b\\d{5,12}\\b/g);
            if (todosNumeros) {
                for (const n of todosNumeros) {
                    resultado[n] = statusAchado;
                }
            }
        }
        return resultado;
    }
    """
    try:
        mapa = target_frame.evaluate(js_status)
        if isinstance(mapa, dict):
            ULTIMOS_STATUS_AMIGO.update(mapa)
            return mapa
    except Exception:
        pass
    return {}

    try:
        mapa = target_frame.evaluate(js_status)
        if isinstance(mapa, dict):
            ULTIMOS_STATUS_AMIGO.update(mapa)
            return mapa
    except Exception:
        pass
    return {}

def buscar_status_portal_amigo_sob_demanda():
    """Abre o Portal Amigo em 2º plano apenas para ler os status atuais de Minhas Atividades."""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        return False, "Biblioteca Playwright não instalada.", {}

    garantir_instalacao_chromium()

    with _LOCK_ROBO_AMIGO:
        try:
            with sync_playwright() as p:
                browser = iniciar_browser_playwright(p)
                context = browser.new_context(viewport={"width": 1366, "height": 2200})
                page = context.new_page()
                page.goto(TECPAR_PORTAL_URL, timeout=60000, wait_until="domcontentloaded")

                target_frame = obter_frame_interno_gas(page, timeout_ms=20000)

                campo_pass_vis = target_frame.locator("input[type='password']:visible").first
                try:
                    campo_pass_vis.wait_for(state="visible", timeout=5000)
                    campo_user_vis = target_frame.locator(
                        "input:visible:not([type='password']):not([type='hidden']):not([type='checkbox']):not([type='radio'])"
                    ).first
                    campo_user_vis.fill(TECPAR_USER)
                    campo_pass_vis.fill(TECPAR_PASS)
                    target_frame.locator("button:visible", has_text=re.compile(r"^\s*Entrar\s*$", re.I)).first.click()
                except Exception:
                    pass

                inicio_painel = time.time()
                while (time.time() - inicio_painel) < 22:
                    target_frame = obter_frame_interno_gas(page, timeout_ms=2500)
                    cand = target_frame.locator("button:visible", has_text=re.compile("Lançar atividade", re.I)).first
                    try:
                        if cand.is_visible():
                            break
                    except Exception:
                        pass
                    page.wait_for_timeout(200)

                page.wait_for_timeout(1500)
                mapa = extrair_status_minhas_atividades(target_frame)
                browser.close()
                return True, f"{len(mapa)} protocolo(s) lido(s) em Minhas Atividades.", mapa
        except Exception as e:
            return False, f"Erro ao consultar Portal Amigo: {e}", {}

        # ==============================================================================
# ==============================================================================
# ==============================================================================
# ROBÔ 3: FORMULÁRIO GOOGLE ASSERTEC (ENVIO DIRETO + PLAYWRIGHT COM SCROLL POPUP)
# ==============================================================================
def normalizar_numero_metragem_assertec(nome_servico_assertec, cabo_bruto, lpu_bruto, eh_principal=True):
    """
    Garante que o campo 'Quantidade / Metragem / Km' receba APENAS números inteiros puros (mínimo 1).
    """
    srv_up = str(nome_servico_assertec or "").upper()
    eh_multi = any(
        termo in srv_up
        for termo in ["AGREGADO", "AVULSO", "BATIMENTO", "CTO", "DIÁRIA", "DIARIA", "KM", "DESLOCAMENTO"]
    )

    if eh_multi or not eh_principal:
        num_str = re.sub(r"\D", "", str(lpu_bruto if not eh_principal else (cabo_bruto or lpu_bruto or 1)))
        val = int(num_str) if num_str else 1
        return str(max(1, val))
    else:
        num_str = re.sub(r"\D", "", str(cabo_bruto or 0))
        val = int(num_str) if num_str else 0
        return str(max(1, val))


def _sem_acento_upper(texto):
    import unicodedata
    s = "".join(c for c in unicodedata.normalize("NFD", str(texto or "")) if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s.upper()).strip()


def _escolher_opcao_exata_forms(valor_desejado, opcoes_disponiveis):
    """Compara o valor com a lista oficial de opções do Google Forms e devolve o texto idêntico."""
    if not opcoes_disponiveis:
        return str(valor_desejado or "").strip()

    alvo = _sem_acento_upper(valor_desejado)

    # 1. Correspondência exata (ignorando acentos e maiúsculas/minúsculas)
    for op in opcoes_disponiveis:
        if _sem_acento_upper(op) == alvo:
            return op

    # 2. Regra específica para Desativação / Retirada (singular ou plural)
    if "DESATIVA" in alvo or "RETIRADA" in alvo:
        for op in opcoes_disponiveis:
            op_n = _sem_acento_upper(op)
            if "DESATIVA" in op_n or "RETIRADA" in op_n:
                return op

    # 3. Correspondência por inclusão parcial
    for op in opcoes_disponiveis:
        op_norm = _sem_acento_upper(op)
        if len(alvo) >= 4 and (alvo in op_norm or op_norm in alvo):
            return op

    # 4. Correspondência por palavras em comum
    palavras_alvo = [w for w in re.split(r"\W+", alvo) if len(w) >= 3]
    melhor_op = None
    maior_score = 0
    for op in opcoes_disponiveis:
        op_norm = _sem_acento_upper(op)
        score = sum(1 for w in palavras_alvo if w in op_norm)
        if score > maior_score:
            maior_score = score
            melhor_op = op

    if melhor_op and maior_score >= 2:
        return melhor_op

    return str(valor_desejado or "").strip()


def preencher_item_google_forms(page, container, valor_alvo):
    val_str = str(valor_alvo or "").strip()
    if not val_str:
        return False

    alvo_norm = _sem_acento_upper(val_str)

    # 1. Campo de Data (input[type='date'] ou campos separados de Dia/Mês/Ano)
    inp_date = container.locator("input[type='date']")
    if inp_date.count() > 0:
        try:
            campo_dt = inp_date.first
            campo_dt.fill(val_str)
            campo_dt.dispatch_event("input")
            campo_dt.dispatch_event("change")
            campo_dt.evaluate("el => el.blur()")
            return True
        except Exception:
            pass

    inp_dia = container.locator("input[aria-label*='Dia'], input[aria-label*='Day']")
    inp_mes = container.locator("input[aria-label*='Mês'], input[aria-label*='Mes'], input[aria-label*='Month']")
    inp_ano = container.locator("input[aria-label*='Ano'], input[aria-label*='Year']")
    if inp_dia.count() > 0 and inp_mes.count() > 0 and inp_ano.count() > 0:
        try:
            partes = val_str.split("-")
            if len(partes) == 3:
                inp_ano.first.fill(partes[0])
                inp_mes.first.fill(partes[1])
                inp_dia.first.fill(partes[2])
                return True
        except Exception:
            pass

    # 2. Radio buttons (Técnico, Empresa, etc.)
    radios = container.locator("[role='radio']")
    if radios.count() > 0:
        textos_radios = []
        for i in range(radios.count()):
            r = radios.nth(i)
            dv = r.get_attribute("data-value") or r.get_attribute("aria-label") or r.inner_text()
            textos_radios.append(str(dv or "").strip())

        melhor_radio = _sem_acento_upper(_escolher_opcao_exata_forms(val_str, textos_radios))
        for i in range(radios.count()):
            r = radios.nth(i)
            dv = _sem_acento_upper(r.get_attribute("data-value") or r.get_attribute("aria-label") or r.inner_text())
            if dv == melhor_radio or dv == alvo_norm:
                r.scroll_into_view_if_needed()
                if r.get_attribute("aria-checked") != "true":
                    r.click(timeout=4000)
                page.wait_for_timeout(150)
                return True

    # 3. Dropdown Listbox (foca exclusivamente no menu popup aberto V68bde / exportSelectPopup)
    listbox = container.locator("[role='listbox']")
    if listbox.count() > 0:
        try:
            listbox.first.scroll_into_view_if_needed()
            listbox.first.click(timeout=4000)
            page.wait_for_timeout(500)

            popup_opcoes = container.locator("div[jsname='V68bde'] div[role='option'], .exportSelectPopup div[role='option']")
            if popup_opcoes.count() == 0:
                popup_opcoes = page.locator("div[jsname='V68bde']:visible div[role='option'], .exportSelectPopup:visible div[role='option']")

            lista_textos = []
            for i in range(popup_opcoes.count()):
                op = popup_opcoes.nth(i)
                txt = str(op.get_attribute("data-value") or op.inner_text() or "").strip()
                if txt and "ESCOLHER" not in _sem_acento_upper(txt) and "SELECIONE" not in _sem_acento_upper(txt):
                    lista_textos.append(txt)

            escolhido_norm = _sem_acento_upper(_escolher_opcao_exata_forms(val_str, lista_textos))

            for i in range(popup_opcoes.count()):
                op = popup_opcoes.nth(i)
                txt_op = _sem_acento_upper(op.get_attribute("data-value") or op.inner_text())
                if not txt_op or "ESCOLHER" in txt_op or "SELECIONE" in txt_op:
                    continue
                if txt_op == escolhido_norm or txt_op == alvo_norm:
                    op.evaluate("el => el.scrollIntoView({block: 'center'})")
                    page.wait_for_timeout(150)
                    op.click(timeout=4000)
                    page.wait_for_timeout(350)
                    return True

            page.keyboard.press("Escape")
        except Exception:
            try:
                page.keyboard.press("Escape")
            except Exception:
                pass

    # 4. Campo de Texto / Número / Observação
    inp_txt = container.locator("input[type='text'], input[type='number'], input:not([type]), textarea")
    if inp_txt.count() > 0:
        try:
            campo = inp_txt.first
            campo.scroll_into_view_if_needed()
            campo.click(timeout=3000)
            campo.fill(val_str)
            campo.dispatch_event("input")
            campo.dispatch_event("change")
            campo.evaluate("el => el.blur()")
            return True
        except Exception:
            pass

    return False


def enviar_formulario_assertec_playwright(dados_os):
    """
    Envia os serviços da OS diretamente para o Google Forms da Assertec em uma única requisição,
    validando que as perguntas do formulário ('role="listitem"') deram lugar à tela de confirmação.
    """
    import requests as req_http

    protocolo_limpo = re.sub(r"\.0+$", "", str(dados_os.get("protocolo", "")).strip())
    protocolo_limpo = re.sub(r"\D", "", protocolo_limpo)
    cliente_nome = str(dados_os.get("cliente", "")).strip().upper()
    cidade_nome = str(dados_os.get("cidade", "")).strip()
    obs_texto = str(dados_os.get("obs", "")).strip() or "Atendimento concluído com sucesso"
    mercado = dados_os.get("mercado", "AMIGO - VAREJO")

    tecnico_raw = str(dados_os.get("tecnico_exec", "ANDERSON")).upper()
    if "ANDERSON" in tecnico_raw:
        nome_tecnico = "Anderson"
    elif "BRUNO" in tecnico_raw:
        nome_tecnico = "Bruno"
    else:
        nome_tecnico = "Marcio"

    data_str = str(dados_os.get("data", datetime.now().strftime("%d/%m/%Y")))[:10]
    dt = None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            dt = datetime.strptime(data_str, fmt)
            break
        except ValueError:
            continue
    if not dt:
        dt = datetime.now()

    servicos_os = dados_os.get("servicos", [])
    if not servicos_os and dados_os.get("servico"):
        servicos_os = [{
            "nome": dados_os.get("servico"),
            "cabo": dados_os.get("cabo", 1),
            "lpu": 1,
            "eh_principal": True
        }]

    try:
        sess = req_http.Session()
        resp_get = sess.get(ASSERTEC_FORM_URL, timeout=15)
        html_form = resp_get.text
        url_post = resp_get.url.split("?")[0]
        if "/viewform" in url_post:
            url_post = url_post.replace("/viewform", "/formResponse")
        elif not url_post.endswith("/formResponse"):
            url_post = url_post.rstrip("/") + "/formResponse"

        match_fbzx = re.search(r'name="fbzx"\s+value="([^"]+)"', html_form)
        fbzx_val = match_fbzx.group(1) if match_fbzx else ""

        match_data = re.search(r"FB_PUBLIC_LOAD_DATA_\s*=\s*(\[.*?\]);\s*</script>", html_form, re.S)
        if not match_data:
            return False, "⚠️ Não foi possível ler a estrutura do Formulário Assertec."

        form_json = json.loads(match_data.group(1))
        perguntas = form_json[1][1] if (len(form_json) > 1 and form_json[1] and len(form_json[1]) > 1) else []

        campos_mapeados = []
        qtd_secoes = 0
        for item in perguntas:
            if not isinstance(item, list) or len(item) < 4:
                continue
            titulo_q = str(item[1] or "")
            tipo_q = item[3]
            if tipo_q == 8:
                qtd_secoes += 1
                continue
            if len(item) > 4 and isinstance(item[4], list) and len(item[4]) > 0:
                sub = item[4][0]
                entry_id = sub[0]
                opcoes_q = []
                if len(sub) > 1 and isinstance(sub[1], list):
                    opcoes_q = [str(opt[0]) for opt in sub[1] if isinstance(opt, list) and len(opt) > 0 and opt[0]]
                campos_mapeados.append({
                    "titulo": _sem_acento_upper(titulo_q),
                    "tipo": tipo_q,
                    "entry_id": entry_id,
                    "opcoes": opcoes_q
                })

        page_history = ",".join(str(i) for i in range(qtd_secoes + 1))
        enviados_http = 0
        erros = []

        for srv in servicos_os:
            nome_srv_assertec = mapear_para_assertec(srv.get("nome", ""), mercado)
            qtd_numerica = normalizar_numero_metragem_assertec(
                nome_srv_assertec,
                srv.get("cabo", 1),
                srv.get("lpu", 1),
                srv.get("eh_principal", True)
            )

            payload = {"fvv": "1", "pageHistory": page_history}
            if fbzx_val:
                payload["fbzx"] = fbzx_val

            for c in campos_mapeados:
                tit = c["titulo"]
                eid = c["entry_id"]
                opcoes = c["opcoes"]

                if "DATA" in tit:
                    payload[f"entry.{eid}_year"] = str(dt.year)
                    payload[f"entry.{eid}_month"] = str(dt.month)
                    payload[f"entry.{eid}_day"] = str(dt.day)
                elif "TECNICO" in tit:
                    payload[f"entry.{eid}"] = _escolher_opcao_exata_forms(nome_tecnico, opcoes)
                elif "EMPRESA" in tit:
                    payload[f"entry.{eid}"] = _escolher_opcao_exata_forms("Fibrawer", opcoes)
                elif "CLIENTE" in tit:
                    payload[f"entry.{eid}"] = cliente_nome
                elif "PROTOCOLO" in tit:
                    payload[f"entry.{eid}"] = protocolo_limpo
                elif "CIDADE" in tit:
                    payload[f"entry.{eid}"] = _escolher_opcao_exata_forms(cidade_nome, opcoes)
                elif "SERVICO" in tit:
                    payload[f"entry.{eid}"] = _escolher_opcao_exata_forms(nome_srv_assertec, opcoes)
                elif "QUANTIDADE" in tit or "METRAGEM" in tit or "KM" in tit:
                    payload[f"entry.{eid}"] = qtd_numerica
                elif "OBSERVA" in tit:
                    payload[f"entry.{eid}"] = obs_texto

            r1 = sess.post(url_post, data=payload, timeout=15)

            # Na página de resposta gravada com sucesso, os blocos de perguntas (role="listitem") desaparecem
            if r1.status_code == 200 and 'role="listitem"' not in r1.text and 'role="alert"' not in r1.text:
                enviados_http += 1
            else:
                erros.append(f"Recusado pelo formulário para {nome_srv_assertec}")

        if enviados_http == len(servicos_os):
            return True, f"🟢 {enviados_http} serviço(s) gravado(s) e confirmado(s) na Planilha Assertec (Fibrawer)!"
        else:
            return False, f"⚠️ Enviados {enviados_http}/{len(servicos_os)} na Assertec. Detalhes: {', '.join(erros)}"

    except Exception as e:
        return False, f"⚠️ Erro ao enviar para o Formulário Assertec: {e}"


def executar_lancamento_amigo_seguro(dados_os):
    """
    Dispara em segundo plano (Thread independente):
    1. O envio do Formulário Google da Assertec (gravando na planilha Fibrawer).
    2. O lançamento no Portal Brasil TecPar (Amigo), se aplicável.
    """
    resultado = {
        "ok": False,
        "msg": "Em processamento...",
        "ok_assertec": False,
        "msg_assertec": ""
    }

    def _worker():
        ok_ass, msg_ass = enviar_formulario_assertec_playwright(dados_os)
        resultado["ok_assertec"] = ok_ass
        resultado["msg_assertec"] = msg_ass

        ok_am, msg_am = executar_lancamento_portal_tecpar(dados_os)
        resultado["ok"] = ok_am
        resultado["msg"] = msg_am

    t = threading.Thread(target=_worker, daemon=False)
    t.start()

    while t.is_alive():
        time.sleep(0.2)

    return resultado["ok"], resultado["msg"], resultado["ok_assertec"], resultado["msg_assertec"]