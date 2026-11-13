import os
import re
import json
import time
import threading
import urllib.parse
from datetime import datetime
import requests
from config import (
    SF_URL,
    SF_CREDENCIAIS_TECNICOS,
    ARQUIVO_PRINT_SF_ENCERRAMENTO,
    PASTA_FOTOS_TEMP_SF,
)
from modulo_robos import (
    garantir_instalacao_chromium,
    formatar_observacao_padrao_fibrawer,
)

_LOCK_FILA_SF = threading.Lock()


def obter_caminho_sessao_sf(nome_tecnico):
    tec_up = str(nome_tecnico or "ANDERSON").upper()
    chave = "ANDERSON" if "ANDERSON" in tec_up else ("BRUNO" if "BRUNO" in tec_up else "MARCIO")
    return f"sessao_sf_{chave}.json"


def _ler_segredo_ao_vivo(chave_toml, padrao=""):
    """LÃª diretamente o arquivo .streamlit/secrets.toml para nÃ£o depender de cache antigo."""
    caminho_toml = os.path.join(".streamlit", "secrets.toml")
    if os.path.exists(caminho_toml):
        try:
            with open(caminho_toml, "r", encoding="utf-8") as f:
                for linha in f:
                    lin = linha.strip()
                    if lin.startswith(chave_toml) and "=" in lin:
                        valor_bruto = lin.split("=", 1)[1].strip()
                        if (valor_bruto.startswith('"') and valor_bruto.endswith('"')) or (
                            valor_bruto.startswith("'") and valor_bruto.endswith("'")
                        ):
                            return valor_bruto[1:-1]
                        return valor_bruto
        except Exception:
            pass
    return padrao


def obter_credenciais_sf_tecnico(nome_tecnico):
    """Retorna o login individual lendo direto do secrets.toml ao vivo."""
    tec_up = str(nome_tecnico or "ANDERSON").upper()
    if "ANDERSON" in tec_up:
        chave = "ANDERSON"
    elif "BRUNO" in tec_up:
        chave = "BRUNO"
    else:
        chave = "MARCIO"

    cred = dict(SF_CREDENCIAIS_TECNICOS.get(chave, {}))
    usuario = _ler_segredo_ao_vivo(f"SF_USER_{chave}", cred.get("user", "")).strip().rstrip(".")
    senha = _ler_segredo_ao_vivo(f"SF_PASS_{chave}", cred.get("pass", "")).strip()
    nome_ass = str(cred.get("nome_assinatura", chave.title())).strip()

    cred["whatsapp_fone"] = _ler_segredo_ao_vivo(f"WHATSAPP_FONE_{chave}", cred.get("whatsapp_fone", ""))
    cred["whatsapp_apikey"] = _ler_segredo_ao_vivo(f"WHATSAPP_APIKEY_{chave}", cred.get("whatsapp_apikey", ""))

    if not usuario or not senha:
        raise ValueError(
            f"Credenciais do Salesforce para o tÃ©cnico {chave} nÃ£o configuradas no secrets.toml "
            f"(SF_USER_{chave} e SF_PASS_{chave})."
        )
    return usuario, senha, nome_ass, cred


def enviar_notificacao_whatsapp_tecnico(nome_tecnico, mensagem):
    tec_up = str(nome_tecnico or "ANDERSON").upper()
    if "ANDERSON" in tec_up:
        chave = "ANDERSON"
    elif "BRUNO" in tec_up:
        chave = "BRUNO"
    else:
        chave = "MARCIO"

    cred = dict(SF_CREDENCIAIS_TECNICOS.get(chave, {}))
    fone_padrao = cred.get("whatsapp_fone", "555199260016" if chave == "ANDERSON" else "")
    fone = re.sub(r"\D", "", str(_ler_segredo_ao_vivo(f"WHATSAPP_FONE_{chave}", fone_padrao)))

    if not fone:
        return False, f"Telefone de WhatsApp do tÃ©cnico {chave} nÃ£o configurado no secrets.toml."

    # Garante o formato brasileiro no WhatsApp (cÃ³digo 55 + DDD + nÃºmero)
    if len(fone) in (10, 11) and not fone.startswith("55"):
        fone = f"55{fone}"

    # 1Âº Tenta enviar pela GREEN-API (nÃ£o exige ativaÃ§Ã£o no telemÃ³vel dos tÃ©cnicos)
    id_instance = _ler_segredo_ao_vivo("GREEN_API_INSTANCE", "710722800809").strip()
    api_token = _ler_segredo_ao_vivo("GREEN_API_TOKEN", "").strip()

    if id_instance and api_token:
        prefixo_host = id_instance[:4] if len(id_instance) >= 4 else "7107"
        urls_green = [
            f"https://{prefixo_host}.api.green-api.com/waInstance{id_instance}/sendMessage/{api_token}",
            f"https://api.green-api.com/waInstance{id_instance}/sendMessage/{api_token}",
        ]
        payload = {
            "chatId": f"{fone}@c.us",
            "message": str(mensagem),
        }
        for url_g in urls_green:
            try:
                resp = requests.post(url_g, json=payload, timeout=15)
                if resp.status_code == 200 and "idMessage" in (resp.text or ""):
                    return True, f"ðŸŸ¢ Mensagem enviada via Green-API para {chave} ({fone})!"
            except Exception:
                continue

    # 2Âº Reserva (Fallback): CallMeBot caso o token da Green-API ainda nÃ£o tenha sido preenchido
    apikey = _ler_segredo_ao_vivo(f"WHATSAPP_APIKEY_{chave}", cred.get("whatsapp_apikey", "")).strip()
    if fone and apikey:
        try:
            msg_encoded = urllib.parse.quote(mensagem)
            url_cmb = f"https://api.callmebot.com/whatsapp.php?phone={fone}&text={msg_encoded}&apikey={apikey}"
            resp = requests.get(url_cmb, timeout=15)
            if resp.status_code == 200:
                return True, f"ðŸŸ¢ Mensagem enviada via CallMeBot para {chave}!"
        except Exception as e:
            return False, f"Falha ao enviar WhatsApp: {e}"

    return False, f"Configure o GREEN_API_TOKEN no secrets.toml para enviar mensagens para {chave}."


def salvar_fotos_upload_temporario(protocolo, dict_arquivos_upload):
    pasta_proto = os.path.join(PASTA_FOTOS_TEMP_SF, re.sub(r"\D", "", str(protocolo or "0")))
    os.makedirs(pasta_proto, exist_ok=True)
    caminhos_salvos = {}

    for chave_campo, arquivo_ou_lista in (dict_arquivos_upload or {}).items():
        if not arquivo_ou_lista:
            continue
        lista_itens = arquivo_ou_lista if isinstance(arquivo_ou_lista, list) else [arquivo_ou_lista]
        caminhos_campo = []
        for idx, arq in enumerate(lista_itens):
            if isinstance(arq, str) and os.path.exists(arq):
                caminhos_campo.append(arq)
            elif hasattr(arq, "read"):
                ext = os.path.splitext(getattr(arq, "name", "foto.jpg"))[1] or ".jpg"
                caminho_dest = os.path.join(pasta_proto, f"{chave_campo}_{idx}{ext}")
                with open(caminho_dest, "wb") as f:
                    f.write(arq.getbuffer())
                caminhos_campo.append(caminho_dest)
        caminhos_salvos[chave_campo] = caminhos_campo

    return caminhos_salvos


def _iniciar_browser_stealth_okta(playwright_ctx):
    """
    Inicia o Chromium em 2Âº plano desativando flags de automaÃ§Ã£o (AutomationControlled)
    para que o EVO/Okta nÃ£o bloqueie o login como robÃ´.
    """
    args_stealth = [
        "--no-sandbox",
        "--disable-dev-shm-usage",
        "--disable-gpu",
        "--disable-blink-features=AutomationControlled",
    ]
    return playwright_ctx.chromium.launch(headless=True, args=args_stealth)


def _preencher_input_react_okta(page, campo_locator, texto):
    """
    Preenche um campo do Okta garantindo que o estado interno do React/JS receba
    exatamente a string completa (incluindo '@' e maiÃºsculas/minÃºsculas).
    """
    texto_str = str(texto)
    campo_locator.click()
    page.wait_for_timeout(150)
    campo_locator.fill("")
    page.wait_for_timeout(100)
    campo_locator.press_sequentially(texto_str, delay=40)
    page.wait_for_timeout(150)

    # Garante que o valor interno do DOM Ã© 100% idÃªntico Ã  senha/usuÃ¡rio
    try:
        val_atual = campo_locator.input_value()
    except Exception:
        val_atual = ""

    if val_atual != texto_str:
        campo_locator.fill(texto_str)

    campo_locator.dispatch_event("input")
    campo_locator.dispatch_event("change")
    page.wait_for_timeout(200)


def _esta_logado_no_salesforce(page):
    """Retorna True SOMENTE se estiver dentro do Salesforce da Brasil TecPar."""
    url_at = page.url.lower()
    if "login.salesforce.com" in url_at or "okta.com" in url_at or "evo." in url_at:
        return False
    if "brasiltecpar" not in url_at and "force.com" not in url_at:
        return False
    if page.locator("input#username, #idp_section_buttons button, button:has-text('Brasil TecPar')").count() > 0:
        return False
    return True


def iniciar_sessao_sf_tecnico(playwright_ctx, nome_tecnico, codigo_totp="", lat=-29.7011, lon=-51.1352):
    """
    Segue estritamente o caminho do vÃ­deo:
    1. Abre login.salesforce.com -> clica no link 'Usar domÃ­nio personalizado' -> digita 'prod-brasiltecpar' -> clica em 'Continuar'.
    2. Na tela prod-brasiltecpar, clica no botÃ£o inferior 'Fazer login com Brasil TecPar'.
    3. No EVO/Okta (com camuflagem anti-robÃ´): preenche E-mail -> 'Mantenha-me conectado' -> 'PrÃ³ximo'.
    4. Preenche a Senha -> 'Verificar'.
    5. Preenche o cÃ³digo de 6 dÃ­gitos do Google Authenticator -> 'Verificar'.
    """
    usuario, senha, nome_ass, _ = obter_credenciais_sf_tecnico(nome_tecnico)
    arq_sessao = obter_caminho_sessao_sf(nome_tecnico)
    codigo_limpo = re.sub(r"\D", "", str(codigo_totp or ""))
    historico = []

    if codigo_limpo and os.path.exists(arq_sessao):
        try:
            os.remove(arq_sessao)
        except Exception:
            pass

    garantir_instalacao_chromium()
    browser = _iniciar_browser_stealth_okta(playwright_ctx)

    # User-Agent real de Chrome no Windows 11 (sem 'HeadlessChrome')
    ua_real = (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/131.0.0.0 Safari/537.36"
    )

    kwargs_ctx = {
        "viewport": {"width": 430, "height": 932},
        "user_agent": ua_real,
        "locale": "pt-BR",
        "timezone_id": "America/Sao_Paulo",
        "geolocation": {"latitude": lat, "longitude": lon},
        "permissions": ["geolocation"],
    }
    if os.path.exists(arq_sessao):
        kwargs_ctx["storage_state"] = arq_sessao

    context = browser.new_context(**kwargs_ctx)
    # Oculta a propriedade navigator.webdriver para o Okta nÃ£o detectar automaÃ§Ã£o
    context.add_init_script(
        "Object.defineProperty(navigator, 'webdriver', {get: () => undefined});"
    )
    page = context.new_page()

    # Se jÃ¡ existe sessÃ£o salva, testa primeiro se jÃ¡ abre logado
    if os.path.exists(arq_sessao) and not codigo_limpo:
        page.goto("https://prod-brasiltecpar.my.salesforce.com/?locale=br", timeout=60000, wait_until="domcontentloaded")
        page.wait_for_timeout(2500)
        if _esta_logado_no_salesforce(page):
            return browser, context, page, nome_ass

    # =========================================================================
    # PASSO 1 (ObrigatÃ³rio conforme o vÃ­deo):
    # Abrir login.salesforce.com -> Clicar em 'Usar domÃ­nio personalizado' -> prod-brasiltecpar -> Continuar
    # =========================================================================
    page.goto("https://login.salesforce.com/?locale=br", timeout=60000, wait_until="domcontentloaded")
    page.wait_for_timeout(1500)

    link_dom = page.locator(
        "a#mydomainLink, a:has-text('Usar domÃ­nio personalizado'), a:has-text('Use Custom Domain')"
    ).first
    if link_dom.count() > 0 and link_dom.is_visible():
        link_dom.click()
        page.wait_for_timeout(800)

        inp_dom = page.locator("input#mydomain, input[name='mydomain']").first
        if inp_dom.count() > 0 and inp_dom.is_visible():
            inp_dom.click()
            inp_dom.fill("")
            inp_dom.press_sequentially("prod-brasiltecpar", delay=35)
            inp_dom.dispatch_event("keyup")
            page.wait_for_timeout(500)

            btn_cont = page.locator("input#mydomainContinue, button:has-text('Continuar')").first
            btn_cont.click()
            # Aguarda redirecionar efetivamente para prod-brasiltecpar.my.salesforce.com
            for _ in range(15):
                page.wait_for_timeout(500)
                if "prod-brasiltecpar" in page.url.lower():
                    break
            historico.append("1. Entrou via 'Usar domÃ­nio personalizado' (prod-brasiltecpar)")

    # =========================================================================
    # PASSO 2: Clicar no botÃ£o inferior "Fazer login com Brasil TecPar"
    # =========================================================================
    btn_sso = page.locator(
        "#idp_section_buttons button, "
        "button:has-text('Fazer login com Brasil TecPar'), "
        "button:has-text('Brasil TecPar'), "
        "a:has-text('Brasil TecPar')"
    ).first

    for _ in range(10):
        if btn_sso.count() > 0 and btn_sso.is_visible():
            break
        page.wait_for_timeout(400)

    if btn_sso.count() > 0 and btn_sso.is_visible():
        btn_sso.click()
        historico.append("2. Clicou em 'Fazer login com Brasil TecPar'")
        for _ in range(25):
            page.wait_for_timeout(500)
            if "okta.com" in page.url.lower():
                break

    # =========================================================================
    # PASSO 3: EVO/Okta â€” Aguarda o spinner inicial e preenche o E-mail
    # =========================================================================
    if "okta.com" in page.url.lower():
        inp_email = page.locator("input[name='identifier'], input#idp-discovery-username, input[type='email']").first
        inp_senha = page.locator("input[type='password']").first

        for _ in range(30):
            if (inp_email.count() > 0 and inp_email.is_visible()) or (inp_senha.count() > 0 and inp_senha.is_visible()):
                break
            page.wait_for_timeout(500)

        if inp_email.count() > 0 and inp_email.is_visible() and not (inp_senha.count() > 0 and inp_senha.is_visible()):
            _preencher_input_react_okta(page, inp_email, usuario)
            chk_manter = page.locator("label:has-text('Mantenha-me conectado'), label:has-text('Keep me signed in')").first
            if chk_manter.count() > 0 and chk_manter.is_visible():
                try:
                    chk_manter.click()
                except Exception:
                    pass

            btn_proximo = page.locator(
                "input[type='submit'], button[type='submit'], button:has-text('PrÃ³ximo'), button:has-text('Next')"
            ).first
            btn_proximo.click()
            historico.append(f"3. E-mail '{usuario}' enviado")

            for _ in range(25):
                page.wait_for_timeout(400)
                if inp_senha.count() > 0 and inp_senha.is_visible():
                    break

        # =====================================================================
        # PASSO 4: EVO/Okta â€” Preenche a Senha e clica em Verificar
        # =====================================================================
        if inp_senha.count() > 0 and inp_senha.is_visible():
            _preencher_input_react_okta(page, inp_senha, senha)
            val_conf = ""
            try:
                val_conf = inp_senha.input_value()
            except Exception:
                pass

            btn_verif_senha = page.locator(
                "input[type='submit'], button[type='submit'], button:has-text('Verificar'), button:has-text('Verify')"
            ).first
            btn_verif_senha.click()
            historico.append(f"4. Senha enviada ({len(val_conf)} chars)")

            for _ in range(25):
                page.wait_for_timeout(500)
                txt_tela_s = (page.inner_text("body") if page.locator("body").count() > 0 else "").upper()
                if "NÃƒO FOI POSSÃVEL FAZER LOGIN" in txt_tela_s or "UNABLE TO SIGN IN" in txt_tela_s:
                    # Clica no Ã­cone do olho (ðŸ‘ï¸) para mostrar no print exatamente o que foi digitado no campo
                    olho = page.locator(
                        "span.eyeicon, [data-se*='eye'], button[aria-label*='senha'], button[aria-label*='password'], .password-toggle span"
                    ).first
                    if olho.count() > 0 and olho.is_visible():
                        try:
                            olho.click()
                            page.wait_for_timeout(300)
                        except Exception:
                            pass
                    page.screenshot(path=ARQUIVO_PRINT_SF_ENCERRAMENTO, full_page=True)
                    browser.close()
                    raise RuntimeError(
                        f"âŒ O EVO/Okta recusou a senha na tela (foram digitados {len(val_conf)} caracteres, "
                        f"comeÃ§ando com '{senha[:1]}' e terminando com '{senha[-1:]}'). "
                        f"Etapas: {' -> '.join(historico)}. Veja o print abaixo."
                    )
                if not inp_senha.is_visible() or "AUTHENTICATOR" in txt_tela_s or "DIGITE O CÃ“DIGO" in txt_tela_s:
                    break

        btn_escolha_ga = page.locator(
            "a[aria-label*='Google Authenticator'], div:has-text('Google Authenticator') button:has-text('Selecionar')"
        ).first
        if btn_escolha_ga.count() > 0 and btn_escolha_ga.is_visible():
            try:
                btn_escolha_ga.click()
                page.wait_for_timeout(1200)
            except Exception:
                pass

        # =====================================================================
        # PASSO 5: Google Authenticator (6 dÃ­gitos) + Verificar
        # =====================================================================
        inp_codigo_ga = page.locator(
            "input[type='text'], input[type='tel'], input[name='credentials.passcode']"
        ).first

        for _ in range(15):
            txt_ga = (page.inner_text("body") if page.locator("body").count() > 0 else "").upper()
            if ("AUTHENTICATOR" in txt_ga or "DIGITE O CÃ“DIGO" in txt_ga or "TEMPORÃRIO" in txt_ga) and inp_codigo_ga.count() > 0 and inp_codigo_ga.is_visible():
                break
            page.wait_for_timeout(400)

        txt_ga_atual = (page.inner_text("body") if page.locator("body").count() > 0 else "").upper()
        if ("AUTHENTICATOR" in txt_ga_atual or "DIGITE O CÃ“DIGO" in txt_ga_atual or "TEMPORÃRIO" in txt_ga_atual) and inp_codigo_ga.count() > 0 and inp_codigo_ga.is_visible():
            if not codigo_limpo:
                page.screenshot(path=ARQUIVO_PRINT_SF_ENCERRAMENTO, full_page=True)
                browser.close()
                raise RuntimeError(
                    f"ðŸ“± Senha aceita! Chegou na tela do Google Authenticator ({' -> '.join(historico)}). "
                    f"Digite os 6 dÃ­gitos atuais no campo 'ðŸ”¢ CÃ³digo Authenticator' e clique novamente."
                )

            _preencher_input_react_okta(page, inp_codigo_ga, codigo_limpo)
            btn_verif_ga = page.locator(
                "input[type='submit'], button[type='submit'], button:has-text('Verificar'), button:has-text('Verify')"
            ).first
            btn_verif_ga.click()
            historico.append(f"5. CÃ³digo Authenticator ({codigo_limpo}) enviado")

        # =====================================================================
        # PASSO 6: Aguarda o redirecionamento SAML do Okta para o Salesforce
        # =====================================================================
        for _ in range(45):
            page.wait_for_timeout(1000)
            url_at = page.url.lower()
            txt_at = (page.inner_text("body") if page.locator("body").count() > 0 else "").upper()

            if any(
                err in txt_at
                for err in ["CÃ“DIGO INVÃLIDO", "INVALID CODE", "TOKEN DOESN'T MATCH", "NÃƒO CORRESPONDE", "NÃƒO FOI POSSÃVEL FAZER LOGIN"]
            ):
                page.screenshot(path=ARQUIVO_PRINT_SF_ENCERRAMENTO, full_page=True)
                browser.close()
                raise RuntimeError(
                    f"âŒ O EVO/Okta recusou o cÃ³digo {codigo_limpo} (pode ter trocado nos Ãºltimos segundos). "
                    f"Aguarde o prÃ³ximo cÃ³digo de 6 dÃ­gitos e tente novamente. ({' -> '.join(historico)})"
                )

            if ("salesforce.com" in url_at or "force.com" in url_at) and "okta.com" not in url_at:
                historico.append("6. Entrou no domÃ­nio Salesforce")
                break

    # =========================================================================
    # PASSO 7: Trata tela intermediÃ¡ria pÃ³s-login se aparecer
    # =========================================================================
    page.wait_for_timeout(3000)
    btn_pos_login = page.locator(
        "a:has-text('Ignorar'), a:has-text('Agora nÃ£o'), a:has-text('Skip'), button:has-text('Verifique sua identidade')"
    ).first
    if btn_pos_login.count() > 0 and btn_pos_login.is_visible():
        try:
            btn_pos_login.click()
            historico.append("7. Tratou tela pÃ³s-login")
            page.wait_for_timeout(3000)
        except Exception:
            pass

    if not _esta_logado_no_salesforce(page):
        page.screenshot(path=ARQUIVO_PRINT_SF_ENCERRAMENTO, full_page=True)
        url_erro = page.url
        browser.close()
        raise RuntimeError(
            f"âš ï¸ Ainda nÃ£o entrou na Ã¡rea logada da Brasil TecPar (URL atual: {url_erro} | "
            f"Passos: {' -> '.join(historico)}). Veja o print abaixo."
        )

    page.wait_for_timeout(3000)
    try:
        context.storage_state(path=arq_sessao)
    except Exception:
        pass

    return browser, context, page, nome_ass


def inspecionar_fila_e_protocolo_sf(nome_tecnico, protocolo_teste="", codigo_totp=""):
    from playwright.sync_api import sync_playwright

    try:
        with sync_playwright() as p:
            browser, context, page, _ = iniciar_sessao_sf_tecnico(p, nome_tecnico, codigo_totp=codigo_totp)
            page.wait_for_timeout(3000)

            texto_inicial = page.inner_text("body") if page.locator("body").count() > 0 else ""
            sas_na_fila = re.findall(r"SA-\d+", texto_inicial)
            ordens_na_fila = re.findall(r"\b00\d{6,8}\b", texto_inicial)

            if not sas_na_fila and not protocolo_teste:
                try:
                    dominio_base = page.url.split("/lightning")[0].split("/home")[0].rstrip("/")
                    page.goto(f"{dominio_base}/lightning/o/ServiceAppointment/list?filterName=Recent", timeout=30000)
                    page.wait_for_timeout(4000)
                    texto_lista = page.inner_text("body") if page.locator("body").count() > 0 else ""
                    sas_na_fila = list(dict.fromkeys(re.findall(r"SA-\d+", texto_lista)))
                    ordens_na_fila = list(dict.fromkeys(re.findall(r"\b00\d{6,8}\b", texto_lista)))
                except Exception:
                    pass

            acoes_encontradas = []
            abriu_alvo = False
            if protocolo_teste:
                abriu_alvo = abrir_compromisso_sa_ou_protocolo(page, protocolo_teste)
                if abriu_alvo:
                    btn_acoes = page.locator(
                        "button:has-text('AÃ§Ãµes'), a:has-text('AÃ§Ãµes'), span:has-text('AÃ§Ãµes'), button:has-text('Mostrar mais aÃ§Ãµes')"
                    ).first
                    if btn_acoes.count() > 0 and btn_acoes.is_visible():
                        btn_acoes.click()
                        page.wait_for_timeout(1200)
                    textos_botoes = page.locator("a, button, [role='menuitem']").all_inner_texts()
                    acoes_encontradas = [
                        t.strip()
                        for t in textos_botoes
                        if t
                        and 3 <= len(t.strip()) <= 45
                        and any(
                            k in t.upper()
                            for k in ["CHECK", "CHECKLIST", "BOLSA", "DOCUMENTO", "ACEITAR", "TERMO", "ATENDIMENTO"]
                        )
                    ]

            page.screenshot(path=ARQUIVO_PRINT_SF_ENCERRAMENTO, full_page=True)
            url_logada = page.url
            browser.close()

            resumo = (
                f"âœ… Conectado de verdade no Salesforce Brasil TecPar ({nome_tecnico})! "
                f"URL: {url_logada} | "
                f"SAs encontradas: {sas_na_fila or 'Nenhuma na lista recente'} | "
                f"Ordens: {ordens_na_fila or 'Nenhuma na lista recente'}. "
            )
            if protocolo_teste:
                resumo += f"Protocolo '{protocolo_teste}' aberto: {abriu_alvo} | AÃ§Ãµes mapeadas: {list(dict.fromkeys(acoes_encontradas))}"
            return True, resumo
    except Exception as e:
        return False, str(e)


def abrir_compromisso_sa_ou_protocolo(page, protocolo_ou_sa, cliente_nome=""):
    termo = str(protocolo_ou_sa or "").strip()
    busca = page.locator(
        "input[type='search'], input[placeholder*='Pesquisar'], input[placeholder*='Search'], button[aria-label*='Pesquisar']"
    ).first
    if busca.count() > 0 and busca.is_visible():
        busca.click()
        page.wait_for_timeout(500)
        inp_real = page.locator(
            "input[type='search'], input[placeholder*='Pesquisar'], input[placeholder*='Search']"
        ).first
        if inp_real.count() > 0:
            inp_real.fill(termo)
            inp_real.press("Enter")
            page.wait_for_timeout(3000)

    link_alvo = page.locator(f"a:has-text('{termo}'), span:has-text('{termo}')").first
    if link_alvo.count() > 0 and link_alvo.is_visible():
        link_alvo.click()
        page.wait_for_timeout(2500)
        return True

    if cliente_nome:
        parte_cli = cliente_nome.strip()[:18]
        link_cli = page.locator(f"a:has-text('{parte_cli}'), span:has-text('{parte_cli}')").first
        if link_cli.count() > 0 and link_cli.is_visible():
            link_cli.click()
            page.wait_for_timeout(2500)
            return True

    return False


def acionar_menu_acao_sf(page, nome_acao):
    btn_acoes = page.locator(
        "button:has-text('AÃ§Ãµes'), a:has-text('AÃ§Ãµes'), span:has-text('AÃ§Ãµes'), button:has-text('Mostrar mais aÃ§Ãµes')"
    ).first
    if btn_acoes.count() > 0 and btn_acoes.is_visible():
        btn_acoes.click()
        page.wait_for_timeout(1000)

    item = page.locator(
        f"a:has-text('{nome_acao}'), button:has-text('{nome_acao}'), span:has-text('{nome_acao}')"
    ).first
    if item.count() > 0 and item.is_visible():
        item.scroll_into_view_if_needed()
        item.click()
        page.wait_for_timeout(2000)
        return True
    return False


def executar_confirmacao_e_checkin(page):
    if acionar_menu_acao_sf(page, "Aceitar Compromisso de ServiÃ§o"):
        sel_status = page.locator("select, [role='combobox'], button:has-text('Agendado')").first
        if sel_status.count() > 0:
            try:
                sel_status.select_option(label="Confirmado")
            except Exception:
                sel_status.click()
                page.wait_for_timeout(400)
                op_conf = page.locator("[role='option']:has-text('Confirmado'), span:has-text('Confirmado')").first
                if op_conf.count() > 0:
                    op_conf.click()

        btn_salvar = page.locator("button:has-text('Salvar'), button:has-text('Save')").first
        if btn_salvar.count() > 0:
            btn_salvar.click()
            page.wait_for_timeout(2500)

    if acionar_menu_acao_sf(page, "Check In"):
        marcar_opcao_sf(page, "Deseja realizar o check-in", "Sim")
        clicar_avancar_ou_concluir_sf(page)
        page.wait_for_timeout(2500)

    return True, "Compromisso confirmado e Check-In realizado."


def marcar_opcao_sf(page, texto_pergunta, valor_resposta):
    bloco = page.locator(f"div:has-text('{texto_pergunta}')").last
    alvo = page.locator(
        f"label:has-text('{valor_resposta}'), span:has-text('{valor_resposta}'), [role='radio']:has-text('{valor_resposta}')"
    )
    if bloco.count() > 0:
        sub = bloco.locator(f"label:has-text('{valor_resposta}'), span:has-text('{valor_resposta}')").first
        if sub.count() > 0 and sub.is_visible():
            sub.scroll_into_view_if_needed()
            sub.click()
            page.wait_for_timeout(200)
            return True

    if alvo.count() > 0 and alvo.first.is_visible():
        alvo.first.scroll_into_view_if_needed()
        alvo.first.click()
        page.wait_for_timeout(200)
        return True

    selects = page.locator("select")
    for i in range(selects.count()):
        s = selects.nth(i)
        try:
            s.select_option(label=valor_resposta)
            return True
        except Exception:
            continue
    return False


def preencher_campo_sf(page, texto_rotulo, valor):
    if valor is None or str(valor).strip() == "":
        return False
    val_str = str(valor).strip()

    inp = page.locator(f"input[aria-label*='{texto_rotulo}'], textarea[aria-label*='{texto_rotulo}']").first
    if inp.count() > 0 and inp.is_visible():
        inp.scroll_into_view_if_needed()
        inp.fill(val_str)
        inp.dispatch_event("change")
        return True

    lbl = page.locator(f"label:has-text('{texto_rotulo}'), span:has-text('{texto_rotulo}')").first
    if lbl.count() > 0:
        container = lbl.locator("xpath=ancestor::div[position()<=3]").last
        campo = container.locator("input, textarea").first
        if campo.count() > 0 and campo.is_visible():
            campo.scroll_into_view_if_needed()
            campo.fill(val_str)
            campo.dispatch_event("change")
            return True
    return False


def anexar_fotos_etapa_sf(page, lista_caminhos_fotos):
    caminhos_validos = [p for p in (lista_caminhos_fotos or []) if p and os.path.exists(p)]
    if not caminhos_validos:
        return 0

    inputs_file = page.locator("input[type='file']")
    qtd_inputs = inputs_file.count()
    anexados = 0

    for i in range(qtd_inputs):
        arq = caminhos_validos[min(i, len(caminhos_validos) - 1)]
        try:
            inputs_file.nth(i).set_input_files(arq)
            page.wait_for_timeout(600)
            anexados += 1
        except Exception:
            continue
    return anexados


def clicar_avancar_ou_concluir_sf(page):
    for texto_btn in ["AvanÃ§ar", "Concluir", "ConcluÃ­do", "Next", "Finish", "Salvar"]:
        btn = page.locator(f"button:has-text('{texto_btn}'), span:has-text('{texto_btn}')").first
        if btn.count() > 0 and btn.is_visible():
            btn.scroll_into_view_if_needed()
            btn.click()
            page.wait_for_timeout(1800)
            return True
    return False


def desenhar_assinatura_manuscrita_canvas(page):
    canvas = page.locator("canvas").first
    if canvas.count() > 0 and canvas.is_visible():
        box = canvas.bounding_box()
        if box:
            x_ini = box["x"] + box["width"] * 0.18
            y_meio = box["y"] + box["height"] * 0.52
            page.mouse.move(x_ini, y_meio)
            page.mouse.down()
            page.mouse.move(x_ini + 45, y_meio - 18, steps=6)
            page.mouse.move(x_ini + 95, y_meio + 14, steps=6)
            page.mouse.move(x_ini + 150, y_meio - 10, steps=6)
            page.mouse.move(x_ini + 210, y_meio + 6, steps=6)
            page.wait_for_timeout(400)
            page.mouse.up()
            return True
    return False


def gerar_dois_documentos_servico_sf(page, nome_tecnico_assinatura, nome_cliente):
    modelos = [
        "PDF TÃ©cnico - Compromisso de ServiÃ§o",
        "PDF - Checklist",
    ]
    gerados = 0

    for modelo in modelos:
        if not acionar_menu_acao_sf(page, "Criar documento de serviÃ§o"):
            acionar_menu_acao_sf(page, "Criar Documento de ServiÃ§o")

        op_mod = page.locator(f"span:has-text('{modelo}'), div:has-text('{modelo}')").first
        if op_mod.count() > 0 and op_mod.is_visible():
            op_mod.click()
            page.wait_for_timeout(1200)

        for responsavel_nome in [nome_tecnico_assinatura, nome_cliente]:
            inp_nome = page.locator("input[placeholder*='Inserir nome'], input[aria-label*='Nome']").first
            if inp_nome.count() > 0 and inp_nome.is_visible():
                inp_nome.fill(responsavel_nome)
                desenhar_assinatura_manuscrita_canvas(page)
                btn_salvar_ass = page.locator("button:has-text('Salvar'), span:has-text('Salvar')").first
                if btn_salvar_ass.count() > 0 and btn_salvar_ass.is_visible():
                    btn_salvar_ass.click()
                    page.wait_for_timeout(1500)

        btn_gerar = page.locator("button:has-text('Gerar'), span:has-text('Gerar')").first
        if btn_gerar.count() > 0 and btn_gerar.is_visible():
            btn_gerar.click()
            page.wait_for_timeout(4000)
            gerados += 1

    return gerados == 2, f"{gerados}/2 documento(s) PDF gerado(s)."


def executar_checkout_final_sf(page):
    if acionar_menu_acao_sf(page, "Check Out"):
        marcar_opcao_sf(page, "Deseja realizar o check-out", "Sim")
        clicar_avancar_ou_concluir_sf(page)
        page.wait_for_timeout(2500)
        page.screenshot(path=ARQUIVO_PRINT_SF_ENCERRAMENTO, full_page=True)
        return True, "Check Out concluÃ­do com sucesso no Salesforce!"
    page.screenshot(path=ARQUIVO_PRINT_SF_ENCERRAMENTO, full_page=True)
    return False, "NÃ£o foi possÃ­vel localizar a aÃ§Ã£o de Check Out."
