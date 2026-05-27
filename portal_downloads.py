"""
Download automático dos exports dos portais 718 e Macan via Playwright.

Período automático:
  - Segunda-feira: sexta anterior → hoje  (cobre fim de semana)
  - Demais dias:   ontem → hoje

Credenciais via variáveis de ambiente:
  PORSCHE_718_USER / PORSCHE_718_PASSWORD
  PORSCHE_MACAN_USER / PORSCHE_MACAN_PASSWORD

O arquivo baixado é salvo como <campaign_folder>/input/PlanilhaB.xlsx,
pronto para ser processado por pipeline_leads.py.
"""
import argparse
import logging
import os
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Iterable

from playwright.sync_api import Download, Page, TimeoutError as PlaywrightTimeoutError, sync_playwright

BASE_PATH = Path(r"C:\Users\BRUSCOVI\OneDrive - Dr. Ing. h.c. F. Porsche AG\Documents\Python\Campaigns")
LOG_DIR = BASE_PATH / "logs"


@dataclass(frozen=True)
class PortalConfig:
    name: str
    url: str
    campaign_folder: str
    username_env: str
    password_env: str


PORTALS = [
    PortalConfig(
        name="718",
        url="https://www.718-porsche.com.br/ileads/login/",
        campaign_folder="718",
        username_env="PORSCHE_718_USER",
        password_env="PORSCHE_718_PASSWORD",
    ),
    PortalConfig(
        name="Macan",
        url="https://www.macan-porsche.com.br/ileads/login/",
        campaign_folder="Macan",
        username_env="PORSCHE_MACAN_USER",
        password_env="PORSCHE_MACAN_PASSWORD",
    ),
]


# ---------------------------------------------------------------------------
# Período
# ---------------------------------------------------------------------------

def automatic_period(today: date | None = None) -> tuple[date, date]:
    """Segunda-feira baixa sex→seg; demais dias baixam ontem→hoje."""
    today = today or date.today()
    if today.weekday() == 0:          # 0 = segunda
        return today - timedelta(days=3), today
    return today - timedelta(days=1), today


def resolve_period(args: argparse.Namespace) -> tuple[date, date]:
    if args.start and args.end:
        return args.start, args.end
    if args.start:
        return args.start, args.start
    if args.end:
        return args.end, args.end
    return automatic_period()


# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

def configure_logging() -> Path:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    log_file = LOG_DIR / f"portal_downloads_{datetime.now():%Y_%m_%d_%H_%M_%S}.log"
    logging.basicConfig(
        filename=log_file,
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        force=True,
    )
    return log_file


# ---------------------------------------------------------------------------
# Helpers de page / frame
# ---------------------------------------------------------------------------

def iter_scopes(page: Page):
    """Itera page principal e depois iframes (para portais com frames embutidos)."""
    yield page
    for frame in page.frames:
        if frame != page.main_frame:
            yield frame


def _try_fill(locator, value: str) -> None:
    """Tenta fill nativo; cai no JS se o campo bloquear."""
    try:
        locator.fill(value)
    except Exception:
        locator.evaluate(
            """(el, v) => {
                el.focus();
                el.value = v;
                el.dispatchEvent(new Event('input', {bubbles: true}));
                el.dispatchEvent(new Event('change', {bubbles: true}));
            }""",
            value,
        )


def set_value(page: Page, selectors: Iterable[str], value: str, timeout: int = 1500) -> bool:
    for scope in iter_scopes(page):
        for selector in selectors:
            locator = scope.locator(selector).first
            try:
                locator.wait_for(state="visible", timeout=timeout)
                _try_fill(locator, value)
                return True
            except Exception:
                continue
    return False


def find_visible_locator(page: Page, selectors: Iterable[str], timeout: int = 1500):
    for scope in iter_scopes(page):
        for selector in selectors:
            locator = scope.locator(selector).first
            try:
                locator.wait_for(state="visible", timeout=timeout)
                return scope, locator
            except PlaywrightTimeoutError:
                continue
    return None, None


def try_click(page: Page, selectors: Iterable[str], timeout: int = 1500) -> bool:
    _, locator = find_visible_locator(page, selectors, timeout=timeout)
    if locator is None:
        return False
    locator.click()
    return True


def ensure_page_ready(page: Page) -> None:
    page.wait_for_load_state("domcontentloaded", timeout=60_000)
    try:
        page.wait_for_load_state("networkidle", timeout=10_000)
    except PlaywrightTimeoutError:
        pass


# ---------------------------------------------------------------------------
# Login
# ---------------------------------------------------------------------------

_LOGIN_USER_SELECTORS = [
    "input[name='usuario']",
    "input[name='user']",
    "input[name='email']",
    "input[name='username']",
    "input[name='login']",
    "input[type='email']",
    "#email",
    "#username",
    "#usuario",
]
_LOGIN_PASS_SELECTORS = [
    "input[name='senha']",
    "input[name='password']",
    "input[name='pass']",
    "input[type='password']",
    "#password",
    "#senha",
]
_LOGIN_SUBMIT_SELECTORS = [
    "button[type='submit']",
    "input[type='submit']",
    "text=Entrar",
    "text=Login",
    "text=Acessar",
]


def login_if_needed(page: Page, config: PortalConfig) -> None:
    username = os.getenv(config.username_env)
    password = os.getenv(config.password_env)
    if not username or not password:
        # Sem credenciais configuradas: aguarda login manual por até 2 minutos
        print(f"\n[{config.name}] Credenciais não configuradas.")
        print(f"[{config.name}] Faça o login manualmente no navegador. O script continua automaticamente após o login.")
        print(f"[{config.name}] Aguardando até 120 segundos...")
        try:
            page.wait_for_url("**/*", wait_until="domcontentloaded", timeout=120_000)
            # Espera sair da página de login
            page.wait_for_function(
                "() => !window.location.href.includes('login')",
                timeout=120_000,
            )
        except Exception:
            pass
        return

    typed_user = set_value(page, _LOGIN_USER_SELECTORS, username)
    if not typed_user:
        # Fallback: preenche o primeiro input de texto visível da página
        try:
            first_input = page.locator("input:not([type='hidden']):not([type='password'])").first
            first_input.wait_for(state="visible", timeout=3000)
            _try_fill(first_input, username)
            typed_user = True
        except Exception:
            pass

    typed_pass = set_value(page, _LOGIN_PASS_SELECTORS, password)
    if not typed_pass:
        # Fallback: preenche o primeiro input de senha visível
        try:
            pass_input = page.locator("input[type='password']").first
            pass_input.wait_for(state="visible", timeout=3000)
            _try_fill(pass_input, password)
            typed_pass = True
        except Exception:
            pass

    if typed_user and typed_pass:
        try_click(page, _LOGIN_SUBMIT_SELECTORS, timeout=2500)
        ensure_page_ready(page)


# ---------------------------------------------------------------------------
# Período nos filtros do portal
# ---------------------------------------------------------------------------

_START_SELECTORS = [
    "input[name='start_date']",
    "input[name='date_start']",
    "input[name='from']",
    "input[placeholder*='Inicial']",
    "input[placeholder*='Início']",
    "input[placeholder*='Inicio']",
    "input[aria-label*='Início']",
    "input[aria-label*='Inicio']",
]
_END_SELECTORS = [
    "input[name='end_date']",
    "input[name='date_end']",
    "input[name='to']",
    "input[placeholder*='Final']",
    "input[aria-label*='Final']",
]


def _fill_date_input_by_index(page: Page, index: int, value: str) -> bool:
    """Preenche o Nth input[type='date'] da página (0=start, 1=end)."""
    for scope in iter_scopes(page):
        inputs = scope.locator("input[type='date']")
        if inputs.count() > index:
            try:
                _try_fill(inputs.nth(index), value)
                return True
            except Exception:
                pass
    return False


def set_period(page: Page, start: date, end: date) -> None:
    start_br = start.strftime("%d/%m/%Y")
    end_br = end.strftime("%d/%m/%Y")
    iso_start = start.isoformat()
    iso_end = end.isoformat()

    # Tenta seletores específicos primeiro; fallback para nth(0) do tipo date
    start_set = False
    for _ in range(3):
        start_set = (
            set_value(page, _START_SELECTORS, start_br)
            or _fill_date_input_by_index(page, 0, iso_start)
        )
        if start_set:
            break
        ensure_page_ready(page)

    # FIX: fallback de data-fim usa nth(1), nunca .first, para não sobrescrever início
    end_set = False
    for _ in range(3):
        end_set = (
            set_value(page, _END_SELECTORS, end_br)
            or _fill_date_input_by_index(page, 1, iso_end)
        )
        if end_set:
            break
        ensure_page_ready(page)

    if not end_set:
        # Último recurso: talvez só haja 1 campo de data (até = hoje é default)
        logging.warning("Não foi possível definir a data de fim; o portal pode usar data corrente como padrão.")


# ---------------------------------------------------------------------------
# Captura do download
# ---------------------------------------------------------------------------

def _save_bytes(path: Path, body: bytes) -> None:
    path.write_bytes(body)


def _request_download_via_href(page: Page, href: str, target: Path) -> bool:
    if not href:
        return False
    resolved = (page.url.rstrip("/") + href) if href.startswith("/") else href
    response = page.request.get(resolved, timeout=60_000)
    if not response.ok:
        return False
    body = response.body()
    if not body:
        return False
    _save_bytes(target, body)
    return True


def _extract_href(locator) -> str:
    try:
        return locator.get_attribute("href") or ""
    except Exception:
        return ""


def capture_download_from_click(page: Page, locator):
    """
    Tenta capturar o artefato gerado pelo clique em ordem de preferência:
      download nativo → popup → resposta HTTP → href direto
    Retorna (kind, value) ou (None, None) se nada funcionou.
    """
    try:
        with page.expect_download(timeout=60_000) as dl_info:
            locator.click()
        return "download", dl_info.value
    except PlaywrightTimeoutError:
        pass

    try:
        with page.expect_popup(timeout=60_000) as popup_info:
            locator.click()
        popup = popup_info.value
        popup.wait_for_load_state("domcontentloaded", timeout=60_000)
        return "popup", popup
    except PlaywrightTimeoutError:
        pass

    def _is_export_response(resp):
        ct = (resp.headers.get("content-type", "") or "").lower()
        url = resp.url.lower()
        return (
            resp.request.method == "GET"
            and (
                "excel" in ct
                or "spreadsheet" in ct
                or "download" in url
                or url.endswith((".xlsx", ".xls", ".csv"))
            )
        )

    try:
        with page.expect_response(_is_export_response, timeout=60_000) as resp_info:
            locator.click()
        return "response", resp_info.value
    except PlaywrightTimeoutError:
        pass

    href = _extract_href(locator)
    if href:
        return "href", href

    return None, None


def save_captured_export(
    page: Page, capture_kind: str, capture_value, target: Path, portal_name: str
) -> None:
    if capture_kind == "download":
        dl: Download = capture_value
        suffix = Path(dl.suggested_filename).suffix or ".xlsx"
        tmp = target.with_suffix(suffix)
        dl.save_as(tmp)
        if tmp != target:
            tmp.rename(target)
        return

    if capture_kind == "popup":
        popup = capture_value
        if _request_download_via_href(page, popup.url, target):
            return
        for i in range(popup.locator("a[href]").count()):
            href = popup.locator("a[href]").nth(i).get_attribute("href") or ""
            if _request_download_via_href(page, href, target):
                return
        raise RuntimeError(f"{portal_name}: popup aberto mas nenhum arquivo encontrado para download.")

    if capture_kind == "response":
        body = capture_value.body()
        if body:
            _save_bytes(target, body)
            return
        raise RuntimeError(f"{portal_name}: resposta do export veio vazia.")

    if capture_kind == "href":
        if _request_download_via_href(page, capture_value, target):
            return
        raise RuntimeError(f"{portal_name}: link de exportação encontrado mas download falhou.")

    raise RuntimeError(f"{portal_name}: falha ao capturar exportação (kind={capture_kind!r}).")


# ---------------------------------------------------------------------------
# Debug
# ---------------------------------------------------------------------------

def dump_debug_artifacts(page: Page, portal_name: str) -> None:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    try:
        page.screenshot(path=str(LOG_DIR / f"{portal_name.lower()}_{ts}.png"), full_page=True)
    except Exception:
        pass
    try:
        (LOG_DIR / f"{portal_name.lower()}_{ts}.html").write_text(page.content(), encoding="utf-8")
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Exportação iLeads
# Fluxo: clicar em "Período" → clicar "Últimos 7 dias" → clicar "Buscar" → clicar "Exportar"
# ---------------------------------------------------------------------------

def _click_periodo_field(page: Page) -> bool:
    """Abre o calendário de período. Retorna True se conseguiu abrir."""
    # Estratégia 1: get_by_placeholder (robusto com acentos)
    for placeholder in ["Período", "Periodo", "período", "periodo"]:
        try:
            loc = page.get_by_placeholder(placeholder).first
            loc.wait_for(state="visible", timeout=3_000)
            loc.click(force=True)
            return True
        except Exception:
            continue

    # Estratégia 2: seletor parcial sem acento
    try:
        loc = page.locator("input[placeholder*='odo']").first
        loc.wait_for(state="visible", timeout=3_000)
        loc.click(force=True)
        return True
    except Exception:
        pass

    # Estratégia 3: JavaScript com múltiplos eventos para garantir abertura do popup
    try:
        page.evaluate("""
            const inputs = document.querySelectorAll('input');
            for (const el of inputs) {
                const ph = (el.placeholder || '').toLowerCase();
                if (ph.includes('odo') || ph.includes('erio') || ph.includes('ata')) {
                    el.focus();
                    el.click();
                    el.dispatchEvent(new MouseEvent('mousedown', {bubbles: true, cancelable: true}));
                    el.dispatchEvent(new MouseEvent('mouseup',   {bubbles: true, cancelable: true}));
                    el.dispatchEvent(new MouseEvent('click',     {bubbles: true, cancelable: true}));
                    break;
                }
            }
        """)
        return True
    except Exception:
        return False


def _click_ultimos_7_dias(page: Page) -> bool:
    """Clica na opção 'Últimos 7 dias' do calendário. Retorna True se conseguiu."""
    # Aguarda o popup renderizar — aumentado para headless
    page.wait_for_timeout(1_500)

    selectors = [
        "li:has-text('7 dias')",
        "span:has-text('7 dias')",
        "a:has-text('7 dias')",
        "li:has-text('ltimos 7')",
        ".ranges li:nth-child(3)",   # "Últimos 7 dias" costuma ser o 3º item
        ".daterangepicker li:nth-child(3)",
    ]
    for selector in selectors:
        try:
            loc = page.locator(selector).first
            loc.wait_for(state="visible", timeout=3_000)
            loc.click(force=True)
            return True
        except Exception:
            continue

    # Fallback JavaScript
    try:
        clicked = page.evaluate("""
            const all = document.querySelectorAll('li, a, span, div');
            for (const el of all) {
                if (el.textContent.trim().includes('7 dias')) {
                    el.click();
                    return true;
                }
            }
            return false;
        """)
        return bool(clicked)
    except Exception:
        return False


def trigger_export(page: Page, download_dir: Path, config: PortalConfig, start: date, end: date) -> Path:
    # Garante que está na página de listagem de leads
    if "leads_list" not in page.url:
        base = page.url.split("/ileads/")[0]
        page.goto(f"{base}/ileads/leads_list/", wait_until="domcontentloaded", timeout=30_000)
        ensure_page_ready(page)

    # 1) Abre o seletor de período
    if not _click_periodo_field(page):
        dump_debug_artifacts(page, config.name)
        raise RuntimeError(f"{config.name}: campo de período não encontrado. Debug em {LOG_DIR}.")

    # 2) Clica em "Últimos 7 dias"
    if not _click_ultimos_7_dias(page):
        dump_debug_artifacts(page, config.name)
        raise RuntimeError(f"{config.name}: opção 'Últimos 7 dias' não apareceu. Debug em {LOG_DIR}.")

    page.wait_for_timeout(500)

    # 3) Clica em "Buscar"
    buscar = page.locator("button:has-text('Buscar'), input[value='Buscar']").first
    try:
        buscar.wait_for(state="visible", timeout=5_000)
        buscar.click()
    except Exception:
        dump_debug_artifacts(page, config.name)
        raise RuntimeError(f"{config.name}: botão 'Buscar' não encontrado.")

    ensure_page_ready(page)
    page.wait_for_timeout(1_000)

    # 4) Clica em "Exportar" e captura o download
    exportar = page.locator("button:has-text('Exportar'), a:has-text('Exportar')").first
    try:
        exportar.wait_for(state="visible", timeout=5_000)
    except Exception:
        dump_debug_artifacts(page, config.name)
        raise RuntimeError(f"{config.name}: botão 'Exportar' não encontrado após busca.")

    capture_kind, capture_value = capture_download_from_click(page, exportar)
    if capture_kind is None:
        dump_debug_artifacts(page, config.name)
        raise RuntimeError(f"{config.name}: clique em 'Exportar' não gerou download. Debug em {LOG_DIR}.")

    final_target = download_dir / "PlanilhaB.xlsx"

    if final_target.exists():
        backup = download_dir / f"PlanilhaB_backup_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        final_target.rename(backup)
        logging.info("%s: arquivo anterior salvo como %s", config.name, backup.name)

    temp_target = download_dir / f"PlanilhaB_{config.name}_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    save_captured_export(page, capture_kind, capture_value, temp_target, config.name)

    if temp_target.exists():
        temp_target.rename(final_target)
    else:
        candidates = sorted(
            download_dir.glob("PlanilhaB_*.xlsx"),
            key=lambda p: p.stat().st_mtime,
            reverse=True,
        )
        if not candidates:
            raise RuntimeError(f"{config.name}: arquivo exportado não encontrado após download.")
        candidates[0].rename(final_target)

    logging.info("%s: export salvo em %s", config.name, final_target)
    return final_target


# ---------------------------------------------------------------------------
# Orquestração por portal
# ---------------------------------------------------------------------------

def download_portal(config: PortalConfig, base_path: Path, start: date, end: date, headless: bool) -> Path:
    download_dir = base_path / config.campaign_folder / "input"
    download_dir.mkdir(parents=True, exist_ok=True)
    storage_state = base_path / ".browser_state" / f"{config.name.lower()}_state.json"
    storage_state.parent.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        ctx_kwargs: dict = {
            "accept_downloads": True,
            "viewport": {"width": 1366, "height": 768},  # evita popup fora do viewport em headless
        }
        if storage_state.exists():
            ctx_kwargs["storage_state"] = str(storage_state)
        context = browser.new_context(**ctx_kwargs)
        page = context.new_page()
        try:
            page.goto(config.url, wait_until="domcontentloaded", timeout=60_000)
            ensure_page_ready(page)
            login_if_needed(page, config)
            downloaded = trigger_export(page, download_dir, config, start, end)
        finally:
            # Persiste sessão mesmo em caso de erro (login pode ter funcionado)
            try:
                context.storage_state(path=str(storage_state))
            except Exception:
                pass
            browser.close()

    return downloaded


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Baixa exports dos portais 718 e Macan via Playwright."
    )
    parser.add_argument("--base-path", type=Path, default=BASE_PATH)
    parser.add_argument(
        "--portal",
        choices=[p.name for p in PORTALS] + ["all"],
        default="all",
    )
    parser.add_argument(
        "--start",
        type=lambda v: datetime.strptime(v, "%Y-%m-%d").date(),
        default=None,
        metavar="YYYY-MM-DD",
    )
    parser.add_argument(
        "--end",
        type=lambda v: datetime.strptime(v, "%Y-%m-%d").date(),
        default=None,
        metavar="YYYY-MM-DD",
    )
    parser.add_argument(
        "--headed",
        action="store_true",
        help="Abre o navegador visível (útil para primeiro login ou depuração).",
    )
    return parser.parse_args()


def main() -> None:
    log_file = configure_logging()
    args = parse_args()
    start, end = resolve_period(args)
    selected = PORTALS if args.portal == "all" else [p for p in PORTALS if p.name == args.portal]

    errors: list[tuple[str, Exception]] = []
    for portal in selected:
        logging.info("%s: período %s → %s", portal.name, start.isoformat(), end.isoformat())
        try:
            path = download_portal(portal, args.base_path, start, end, headless=not args.headed)
            print(f"✓ {portal.name}: export salvo em {path}")
            logging.info("%s: concluído", portal.name)
        except Exception as exc:
            logging.error("%s: FALHA — %s", portal.name, exc, exc_info=True)
            print(f"✗ {portal.name}: FALHA — {exc}")
            errors.append((portal.name, exc))

    print(f"\nLog: {log_file}")
    if errors:
        failed = ", ".join(name for name, _ in errors)
        raise SystemExit(f"Portais com falha: {failed}")


if __name__ == "__main__":
    main()
