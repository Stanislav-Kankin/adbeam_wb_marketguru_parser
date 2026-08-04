from __future__ import annotations

import json
import os
import re
import socket
import subprocess
import time
from pathlib import Path
from urllib.error import HTTPError
from urllib.parse import urlencode, urljoin, urlparse
from urllib.request import Request, urlopen

from playwright.sync_api import Browser, BrowserContext, Locator, Page, Response, sync_playwright

from .models import InspectResult, ResearchRow

PRODUCT_GOTO_TIMEOUT_MS = 12_000
PRODUCT_READY_TIMEOUT_MS = 30_000
FAST_PAGE_READY_TIMEOUT_MS = 8_000
REQUISITES_APPEAR_TIMEOUT_MS = 1_500
PUBLIC_API_TIMEOUT_SECONDS = 5
SELLER_GOTO_TIMEOUT_MS = 12_000
WAIT_FOR_LOAD_STATE_TIMEOUT_MS = 8_000
MAX_BATCH_ROW_ATTEMPTS = 3
BASE_TOOLTIP_REVEAL_ROUNDS = 6
DEEP_TOOLTIP_REVEAL_ROUNDS = 12
BASE_TOOLTIP_PAUSE_MS = 60
DEEP_TOOLTIP_PAUSE_MS = 180
SELLER_LINK_DISCOVERY_ROUNDS = 8
SELLER_LINK_DISCOVERY_PAUSE_MS = 180

INN_RE = re.compile(r"\b(?:ИНН)\s*[:№]?\s*(\d{10}|\d{12}|\d{14})\b", re.IGNORECASE)
OGRN_RE = re.compile(r"\b(?:ОГРН)\s*[:№]?\s*(\d{13})\b", re.IGNORECASE)
OGRNIP_RE = re.compile(r"\b(?:ОГРНИП)\s*[:№]?\s*(\d{15})\b", re.IGNORECASE)
BELARUS_RE = re.compile(r"\b(?:Республика\s+Беларусь|Беларусь|УНП)\b", re.IGNORECASE)
KAZAKHSTAN_RE = re.compile(r"\b(?:Республика\s+Казахстан|Казахстан|БИН|ИИН)\b", re.IGNORECASE)
ENTITY_FULL_RE = re.compile(r"(Индивидуальный предприниматель|Общество с ограниченной ответственностью)", re.IGNORECASE)
ENTITY_SHORT_RE = re.compile(r"\b(ИП|ООО)\b")
SELLER_NAME_RE = re.compile(
    r"(?:Продавец|Seller|Поставщик)\s*[:]?\s*([A-Za-zА-Яа-яЁё0-9 .,&\-\"'()]{2,120})",
    re.IGNORECASE,
)
TOOLTIP_BLOCK_RE = re.compile(
    r"<div[^>]*class=\"[^\"]*tooltip-supplier[^\"]*\"[^>]*>(?P<body>.*?)</div>",
    re.IGNORECASE | re.DOTALL,
)
SELLER_HEADER_RE = re.compile(
    r"<h1[^>]*>(?P<name>.*?)</h1>",
    re.IGNORECASE | re.DOTALL,
)
SELLER_DETAILS_TITLE_RE = re.compile(
    r"<(?:h1|h2)[^>]*class=\"[^\"]*seller-details__title[^\"]*\"[^>]*>(?P<name>.*?)</(?:h1|h2)>",
    re.IGNORECASE | re.DOTALL,
)
TOOLTIP_TEXT_RE = re.compile(
    r"<div[^>]*class=\"[^\"]*tooltip__content[^\"]*\"[^>]*>(?P<body>.*?)</div>",
    re.IGNORECASE | re.DOTALL,
)
TAG_RE = re.compile(r"<[^>]+>")
SPACE_RE = re.compile(r"\s+")
WB_NUMERIC_SELLER_PATH_RE = re.compile(r"^/seller/\d+/?$", re.IGNORECASE)
WB_GENERIC_SELLER_PATH_RE = re.compile(r"^/seller/[^/?#]+/?$", re.IGNORECASE)

IGNORED_SELLER_DISPLAY_NAMES = {
    "все товары",
    "главная",
    "адреса",
    "корзина",
    "франкфурт",
    "wildberries",
    "wibes",
}

EMPTY_RESULTS_PATTERNS = [
    "по вашему запросу ничего не найдено",
    "ничего не найдено",
    "ничего не нашлось",
    "nothing found",
    "no results found",
]

ANTI_BOT_PATTERNS = [
    "Подозрительная активность",
    "Что-то не так",
    "Проверяем браузер",
    "Почти готово",
    "captcha-support@rwb.ru",
    "Новая попытка через",
]

PRODUCT_SELLER_LINK_SELECTORS = [
    'a[aria-label="Подробнее о продавце"]',
    'a[href*="/seller/"]',
]

SELLER_TOOLTIP_TRIGGER_SELECTORS = [
    '.seller-details__title-wrap .seller-details__tip-info',
    '.seller-details__tip-info',
    '.seller-details__info-wrap .seller-details__tip-info',
    '.seller-details__info-icon',
    '.seller-info__info-icon',
    '.seller-info i',
]

SAFE_CLICK_SELLER_TOOLTIP_TRIGGER_SELECTORS = {
    '.seller-details__title-wrap .seller-details__tip-info',
    '.seller-details__tip-info',
    '.seller-details__info-wrap .seller-details__tip-info',
    '.seller-details__info-icon',
    '.seller-info__info-icon',
    '.seller-info i',
}

TOOLTIP_VISIBLE_SELECTORS = [
    ".tooltip.tooltip-supplier",
    '[class*="tooltip-supplier"]',
]

DEEP_SELLER_TOOLTIP_TRIGGER_SELECTORS = [
    '[class*="seller"] [class*="tip"]',
    '[class*="seller"] [class*="info"]',
    '[class*="seller"] [class*="tooltip"]',
    '[class*="seller"] [class*="icon"]',
    '[class*="seller" i] button',
    'button[class*="seller" i]',
    '[class*="seller" i] [class*="info" i]',
    '[class*="seller" i] [class*="icon" i]',
]


class TransientWBError(RuntimeError):
    pass


class _BrowserSession:
    def __init__(
        self,
        context: BrowserContext,
        browser: Browser | None = None,
        process: subprocess.Popen | None = None,
    ) -> None:
        self.context = context
        self.browser = browser
        self.process = process

    def close(self) -> None:
        try:
            self.context.close()
        except Exception:
            pass
        if self.browser is not None:
            try:
                self.browser.close()
            except Exception:
                pass
        if self.process is not None and self.process.poll() is None:
            _terminate_process_tree(self.process.pid)


def _extract_requisites_text_via_dom(page: Page, include_body_scan: bool = False) -> str | None:
    try:
        text = page.evaluate(
            r"""
            (includeBodyScan) => {
              const markers = ['ИНН', 'ОГРН', 'ОГРНИП', 'Номер регистрации', 'КПП', 'УНП', 'БИН', 'ИИН', 'Республика Беларусь', 'Казахстан', 'Индивидуальный предприниматель', 'Общество с ограниченной ответственностью'];
              const selectors = [
                '.tooltip.tooltip-supplier',
                '[class*="tooltip-supplier"]',
                '.tooltip__content',
                '.seller-details',
              ];

              if (includeBodyScan) {
                selectors.push('body');
              }

              const normalize = (value) => (value || '').replace(/\s+/g, ' ').trim();
              const candidates = [];

              for (const selector of selectors) {
                const nodes = Array.from(document.querySelectorAll(selector));
                for (const node of nodes.reverse()) {
                  const raw = normalize(node.innerText || node.textContent || '');
                  if (raw && markers.some((marker) => raw.includes(marker))) {
                    candidates.push(raw);
                  }
                }
              }

              if (includeBodyScan) {
                const allNodes = Array.from(document.querySelectorAll('body *'));
                for (const node of allNodes.reverse()) {
                  const raw = normalize(node.innerText || node.textContent || '');
                  if (!raw || raw.length < 15) {
                    continue;
                  }
                  if (markers.some((marker) => raw.includes(marker))) {
                    candidates.push(raw);
                    break;
                  }
                }
              }

              return candidates[0] || null;
            }
            """,
            include_body_scan,
        )
    except Exception:
        return None

    normalized = _normalize_text(text or "")
    return normalized or None




class BatchInspector:
    def __init__(self, artifacts_dir: Path, headful: bool = True, profile_dir: Path | None = None):
        self.artifacts_dir = artifacts_dir
        self.headful = headful
        self.profile_dir = profile_dir
        self._playwright = None
        self._browser_session: _BrowserSession | None = None
        self._context: BrowserContext | None = None
        self._page: Page | None = None

    def __enter__(self) -> "BatchInspector":
        self.artifacts_dir.mkdir(parents=True, exist_ok=True)
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self._close_context()
        if self._playwright is not None:
            self._playwright.stop()
            self._playwright = None

    def _close_context(self) -> None:
        if self._browser_session is not None:
            self._browser_session.close()
            self._browser_session = None
        self._context = None
        self._page = None

    def _restart_context(self) -> None:
        if self._playwright is None:
            raise RuntimeError('Playwright is not started')
        self._close_context()
        self._browser_session = _open_browser_session(
            playwright=self._playwright,
            headful=self.headful,
            profile_dir=self.profile_dir,
        )
        self._context = self._browser_session.context
        self._page = _fresh_batch_page(self._context)

    def _ensure_browser_started(self) -> None:
        if self._playwright is not None:
            return
        self._playwright = sync_playwright().start()
        self._restart_context()

    def _ensure_page(self) -> Page:
        if self._context is None:
            raise RuntimeError('BatchInspector is not started')
        if self._page is None or self._page.is_closed() or _is_unexpected_page_url(self._page.url):
            self._page = _fresh_batch_page(self._context)
        return self._page

    def inspect_row(self, row_number: int, research_row: ResearchRow) -> InspectResult:
        api_result = _inspect_product_via_public_api(
            row_number=row_number,
            research_row=research_row,
            artifacts_dir=self.artifacts_dir,
        )
        if api_result is not None:
            return api_result

        self._ensure_browser_started()
        if self._context is None:
            raise RuntimeError('BatchInspector browser is not started')

        last_error: Exception | None = None
        for attempt in range(MAX_BATCH_ROW_ATTEMPTS):
            try:
                page = self._ensure_page()
                result = _inspect_product_row_on_page(
                    page=page,
                    row_number=row_number,
                    research_row=research_row,
                    artifacts_dir=self.artifacts_dir,
                    profile_dir=self.profile_dir,
                    manual_wait_seconds=0,
                    fast_success_exit=True,
                )

                if _is_unexpected_page_url(result.final_url):
                    raise TransientWBError(f'Unexpected page URL after inspect: {result.final_url}')

                if self._page is not None and _is_unexpected_page_url(self._page.url):
                    self._page = _fresh_batch_page(self._context)
                return result
            except Exception as exc:
                last_error = exc
                if attempt >= MAX_BATCH_ROW_ATTEMPTS - 1:
                    raise

                if _should_restart_context(exc):
                    self._restart_context()
                    continue

                if _is_retryable_row_error(exc):
                    self._page = _fresh_batch_page(self._context)
                    continue

                raise

        if last_error is not None:
            raise last_error
        raise RuntimeError('inspect_row failed without explicit error')


def _inspect_product_via_public_api(
    row_number: int,
    research_row: ResearchRow,
    artifacts_dir: Path,
) -> InspectResult | None:
    if research_row.wb_nm_id is None:
        return None

    card_query = urlencode(
        {
            "appType": 1,
            "curr": "rub",
            "dest": 1259570991,
            "spp": 30,
            "nm": research_row.wb_nm_id,
        }
    )
    try:
        card_payload = _load_public_wb_json(f"https://card.wb.ru/cards/v4/detail?{card_query}")
    except Exception:
        return None

    products = card_payload.get("products") if isinstance(card_payload, dict) else None
    if not isinstance(products, list):
        return None
    product = next(
        (
            item
            for item in products
            if isinstance(item, dict) and str(item.get("id")) == str(research_row.wb_nm_id)
        ),
        None,
    )
    if product is None:
        return None

    supplier_id = product.get("supplierId")
    try:
        supplier_id = int(supplier_id)
    except (TypeError, ValueError):
        return None

    legal_url = f"https://static-basket-01.wbbasket.ru/vol0/data/supplier-by-id/{supplier_id}.json"
    try:
        legal_payload = _load_public_wb_json(legal_url)
    except HTTPError as exc:
        if exc.code != 404:
            return None
        legal_payload = {}
    except Exception:
        return None

    if not isinstance(legal_payload, dict):
        legal_payload = {}

    inn = _normalize_identifier(legal_payload.get("inn"), {10, 12, 14})
    registration_number = _normalize_identifier(legal_payload.get("ogrn"), {13, 15})
    ogrn = registration_number if registration_number and len(registration_number) == 13 else None
    ogrnip = registration_number if registration_number and len(registration_number) == 15 else None
    legal_name = _first_non_empty(
        _string_value(legal_payload.get("supplierName")),
        _string_value(legal_payload.get("name")),
        _string_value(legal_payload.get("tradeName")),
    )
    seller_name = _first_non_empty(
        _string_value(product.get("supplier")),
        research_row.seller_name_raw,
    )
    entity_type = _extract_entity_type(legal_name or "")
    legal_text = json.dumps(legal_payload, ensure_ascii=False)
    seller_country = _detect_seller_country(legal_text)
    seller_url = f"https://www.wildberries.ru/seller/{supplier_id}"

    if inn:
        parse_status = "SUCCESS"
        note = "ИНН получен из открытых данных Wildberries"
    elif seller_country:
        parse_status = seller_country
        note = f"Обнаружен продавец из страны: {seller_country}"
    elif entity_type or legal_payload:
        parse_status = "NEEDS_REVIEW"
        note = "WB вернул данные продавца, но ИНН отсутствует"
    else:
        parse_status = "Нет реквизитов на странице"
        note = "WB не вернул реквизиты продавца"

    result = InspectResult(
        row_number=row_number,
        url=research_row.wb_candidate_url or "",
        final_url=seller_url,
        http_status=200,
        parse_status=parse_status,
        anti_bot_detected=False,
        used_persistent_profile=False,
        seller_url=seller_url,
        navigated_to_seller_page=False,
        inn=inn,
        ogrn=ogrn,
        ogrnip=ogrnip,
        entity_type=entity_type,
        seller_display_name=seller_name,
        note=note,
    )
    _write_result_json(artifacts_dir=artifacts_dir, row_number=row_number, result=result)
    return result


def _load_public_wb_json(url: str) -> dict:
    request = Request(
        url,
        headers={
            "Accept": "application/json",
            "Referer": "https://www.wildberries.ru/",
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 Chrome/150.0.0.0 Safari/537.36"
            ),
        },
    )
    with urlopen(request, timeout=PUBLIC_API_TIMEOUT_SECONDS) as response:
        payload = json.load(response)
    if not isinstance(payload, dict):
        raise ValueError("WB API returned a non-object payload")
    return payload


def _string_value(value: object) -> str | None:
    if value is None:
        return None
    normalized = _normalize_text(str(value))
    return normalized or None


def _normalize_identifier(value: object, valid_lengths: set[int]) -> str | None:
    if value is None:
        return None
    digits = re.sub(r"\D", "", str(value))
    return digits if len(digits) in valid_lengths else None


def _is_unexpected_page_url(url: str | None) -> bool:
    if not url or url == 'about:blank':
        return False
    try:
        parsed = urlparse(url)
        host = (parsed.netloc or '').lower()
    except Exception:
        return False
    if bool(host) and 'wildberries.ru' not in host:
        return True
    if '/seller/' in (parsed.path or '') and not _is_supported_wb_seller_url(url):
        return True
    return False


def _fresh_batch_page(context: BrowserContext) -> Page:
    reusable_page: Page | None = None
    extra_pages: list[Page] = []

    for page in list(context.pages):
        try:
            if page.is_closed():
                continue
        except Exception:
            continue

        if reusable_page is None:
            reusable_page = page
        else:
            extra_pages.append(page)

    for page in extra_pages:
        try:
            page.close()
        except Exception:
            pass

    if reusable_page is None:
        reusable_page = context.new_page()

    try:
        reusable_page.goto('about:blank', wait_until='domcontentloaded', timeout=1_500)
    except Exception:
        pass
    return reusable_page


def _wait_for_product_page_ready(page: Page, nm_id: int | None, timeout_ms: int) -> bool:
    expected_nm_id = str(nm_id) if nm_id is not None else ""
    expected_url_marker = f"/catalog/{expected_nm_id}/" if expected_nm_id else "/catalog/"
    deadline = time.monotonic() + max(timeout_ms, 0) / 1_000

    while time.monotonic() < deadline:
        title = _safe_page_title(page) or ""
        if expected_nm_id and expected_nm_id in title:
            return True
        try:
            product_content_loaded = bool(
                page.evaluate(
                    """
                    ({ expectedId, expectedUrlMarker }) => {
                      const text = document.body?.innerText || '';
                      if (!location.pathname.includes(expectedUrlMarker) || text.length < 300) {
                        return false;
                      }
                      return Boolean(
                        (expectedId && text.includes(expectedId)) ||
                        text.includes('Артикул') ||
                        text.includes('Добавить в корзину') ||
                        text.includes('Купить сейчас')
                      );
                    }
                    """,
                    {"expectedId": expected_nm_id, "expectedUrlMarker": expected_url_marker},
                )
            )
            if product_content_loaded:
                return True
        except Exception:
            pass
        try:
            if page.locator('a[aria-label="Подробнее о продавце"]').count() > 0:
                return True
        except Exception:
            pass
        try:
            page.wait_for_timeout(500)
        except Exception:
            return False
    return False


def _wait_for_seller_page_ready(page: Page, seller_name: str | None, timeout_ms: int = 30_000) -> bool:
    expected_name = _normalize_text(seller_name or "").casefold()
    deadline = time.monotonic() + max(timeout_ms, 0) / 1_000

    while time.monotonic() < deadline:
        title = _normalize_text(_safe_page_title(page) or "")
        try:
            body_text = page.evaluate("() => (document.body?.innerText || '').slice(0, 12000)") or ""
        except Exception:
            body_text = ""
        combined = _normalize_text(f"{title}\n{body_text}")
        if expected_name and expected_name in combined.casefold():
            return True
        if _contains_seller_signal(combined):
            return True
        try:
            if page.locator('.seller-details__title, [class*="seller-details"]').count() > 0:
                return True
        except Exception:
            pass
        try:
            page.wait_for_timeout(500)
        except Exception:
            return False
    return False


def _save_blocked_product_result(
    page: Page,
    response: Response | None,
    row_number: int,
    research_row: ResearchRow,
    screenshot_path: Path,
    html_path: Path,
    text_path: Path,
    artifacts_dir: Path,
    profile_dir: Path | None,
    manual_wait_seconds: int,
    seller_url: str | None = None,
    note: str | None = None,
    save_diagnostics: bool = True,
) -> InspectResult:
    blocked_note = note or (
        "WB не загрузил карточку товара и оставил пустую оболочку сайта."
    )
    if not save_diagnostics:
        page_title = _safe_page_title(page)
        try:
            page_text = page.evaluate("() => (document.body?.innerText || '').slice(0, 12000)") or ""
        except Exception:
            page_text = ""
        http_status = response.status if response else None
        anti_bot_detected = _contains_anti_bot_text(f"{page_title or ''}\n{page_text}", http_status)
        result = InspectResult(
            row_number=row_number,
            url=research_row.wb_candidate_url or "",
            page_title=page_title,
            final_url=page.url,
            http_status=http_status,
            parse_status="ANTI_BOT_PAGE" if anti_bot_detected else "PAGE_LOAD_TIMEOUT",
            anti_bot_detected=anti_bot_detected,
            used_persistent_profile=profile_dir is not None,
            profile_dir=str(profile_dir) if profile_dir else None,
            manual_wait_seconds=manual_wait_seconds,
            seller_url=seller_url,
            navigated_to_seller_page=bool(seller_url),
            seller_display_name=research_row.seller_name_raw,
            note=blocked_note,
        )
        _write_result_json(artifacts_dir=artifacts_dir, row_number=row_number, result=result)
        return result

    try:
        page.screenshot(path=str(screenshot_path), full_page=True)
    except Exception:
        screenshot_path = None
    html = _safe_page_content(page)
    text = _safe_page_text(page)
    html_path.write_text(html, encoding="utf-8")
    text_path.write_text(text, encoding="utf-8")

    result = _build_result(
        row_number=row_number,
        url=research_row.wb_candidate_url or "",
        page=page,
        http_status=response.status if response else None,
        html=html,
        text=text,
        captured_tooltip_text=None,
        screenshot_path=screenshot_path,
        html_path=html_path,
        text_path=text_path,
        used_persistent_profile=profile_dir is not None,
        profile_dir=profile_dir,
        manual_wait_seconds=manual_wait_seconds,
        seller_url=seller_url,
        navigated_to_seller_page=bool(seller_url),
    )
    result.anti_bot_detected = True
    result.parse_status = "ANTI_BOT_PAGE"
    result.note = note or (
        "WB не загрузил карточку товара и оставил пустую оболочку сайта. "
        "Пакетный прогон нужно остановить и повторить после восстановления доступа."
    )
    _write_result_json(artifacts_dir=artifacts_dir, row_number=row_number, result=result)
    return result


def _should_restart_context(exc: Exception) -> bool:
    message = str(exc).lower()
    restart_markers = [
        'failed to open a new tab',
        'target.createtarget',
        'target page, context or browser has been closed',
        'browser has been closed',
        'context has been closed',
    ]
    return isinstance(exc, TransientWBError) or any(marker in message for marker in restart_markers)


def _is_retryable_row_error(exc: Exception) -> bool:
    if isinstance(exc, TransientWBError):
        return True

    message = str(exc).lower()
    retry_markers = [
        'timeout',
        'navigation',
        'net::err',
        'page.goto',
        'page.content',
        'page has been closed',
        'page is navigating',
    ]
    return any(marker in message for marker in retry_markers)


def _inspect_product_row_on_page(
    page: Page,
    row_number: int,
    research_row: ResearchRow,
    artifacts_dir: Path,
    profile_dir: Path | None = None,
    manual_wait_seconds: int = 0,
    fast_success_exit: bool = False,
) -> InspectResult:
    if not research_row.wb_candidate_url:
        raise ValueError('У строки нет wb_candidate_url')

    artifacts_dir.mkdir(parents=True, exist_ok=True)
    screenshot_path = artifacts_dir / f"row_{row_number}.png"
    html_path = artifacts_dir / f"row_{row_number}.html"
    text_path = artifacts_dir / f"row_{row_number}_text.txt"

    previous_url = page.url
    response = page.goto(
        research_row.wb_candidate_url,
        wait_until='commit',
        timeout=PRODUCT_GOTO_TIMEOUT_MS,
    )
    product_ready = _wait_for_product_page_ready(
        page=page,
        nm_id=research_row.wb_nm_id,
        timeout_ms=(
            8_000
            if response is not None and response.status == 498
            else FAST_PAGE_READY_TIMEOUT_MS if fast_success_exit else PRODUCT_READY_TIMEOUT_MS
        ),
    )
    if not product_ready:
        return _save_blocked_product_result(
            page=page,
            response=response,
            row_number=row_number,
            research_row=research_row,
            screenshot_path=screenshot_path,
            html_path=html_path,
            text_path=text_path,
            artifacts_dir=artifacts_dir,
            profile_dir=profile_dir,
            manual_wait_seconds=manual_wait_seconds,
            save_diagnostics=not fast_success_exit,
        )

    _best_effort_wait(page, settle_rounds=1)
    _ensure_not_stuck_on_previous_seller_page(
        page=page,
        research_row=research_row,
        previous_url=previous_url,
    )

    seller_url, seller_response, seller_page_ready = _go_to_seller_page(
        page,
        expected_seller_name=research_row.seller_name_raw,
        ready_timeout_ms=FAST_PAGE_READY_TIMEOUT_MS if fast_success_exit else PRODUCT_READY_TIMEOUT_MS,
    )
    if seller_response is not None:
        response = seller_response
    if seller_url is None and _is_supported_wb_seller_url(page.url):
        seller_url = page.url
    if seller_url and not seller_page_ready:
        return _save_blocked_product_result(
            page=page,
            response=response,
            row_number=row_number,
            research_row=research_row,
            screenshot_path=screenshot_path,
            html_path=html_path,
            text_path=text_path,
            artifacts_dir=artifacts_dir,
            profile_dir=profile_dir,
            manual_wait_seconds=manual_wait_seconds,
            seller_url=seller_url,
            save_diagnostics=not fast_success_exit,
            note=(
                "WB открыл адрес продавца, но не загрузил страницу с реквизитами. "
                "Строка сохранена без ИНН."
            ),
        )

    captured_tooltip_text = _reveal_supplier_requisites(page)
    if manual_wait_seconds > 0:
        time.sleep(manual_wait_seconds)
        captured_tooltip_text = captured_tooltip_text or _reveal_supplier_requisites(page)

    navigated_to_seller_page = bool(seller_url) or '/seller/' in page.url

    if fast_success_exit:
        fast_result = _build_result(
            row_number=row_number,
            url=research_row.wb_candidate_url,
            page=page,
            http_status=response.status if response else None,
            html="",
            text="",
            captured_tooltip_text=captured_tooltip_text,
            screenshot_path=None,
            html_path=None,
            text_path=None,
            used_persistent_profile=profile_dir is not None,
            profile_dir=profile_dir,
            manual_wait_seconds=manual_wait_seconds,
            seller_url=seller_url,
            navigated_to_seller_page=navigated_to_seller_page,
        )
        if not fast_result.seller_display_name and research_row.seller_name_raw:
            fast_result.seller_display_name = research_row.seller_name_raw
        _write_result_json(artifacts_dir=artifacts_dir, row_number=row_number, result=fast_result)
        return fast_result

    html_before_screenshot = _safe_page_content(page)
    captured_tooltip_text = captured_tooltip_text or _extract_tooltip_text_from_page(page) or _extract_tooltip_text_from_html(html_before_screenshot)

    page.screenshot(path=str(screenshot_path), full_page=True)
    html = _safe_page_content(page)
    html_path.write_text(html, encoding='utf-8')
    text = _safe_page_text(page)
    text_path.write_text(text, encoding='utf-8')

    result = _build_result(
        row_number=row_number,
        url=research_row.wb_candidate_url,
        page=page,
        http_status=response.status if response else None,
        html=html,
        text=text,
        captured_tooltip_text=captured_tooltip_text,
        screenshot_path=screenshot_path,
        html_path=html_path,
        text_path=text_path,
        used_persistent_profile=profile_dir is not None,
        profile_dir=profile_dir,
        manual_wait_seconds=manual_wait_seconds,
        seller_url=seller_url,
        navigated_to_seller_page=navigated_to_seller_page,
    )

    if _should_run_deep_recovery(result):
        deep_result = _run_deep_requisites_recovery(
            page=page,
            row_number=row_number,
            research_row=research_row,
            screenshot_path=screenshot_path,
            html_path=html_path,
            text_path=text_path,
            used_persistent_profile=profile_dir is not None,
            profile_dir=profile_dir,
            manual_wait_seconds=manual_wait_seconds,
            seller_url=seller_url,
            initial_http_status=response.status if response else None,
        )
        if _result_information_score(deep_result) > _result_information_score(result):
            result = deep_result

    _write_result_json(artifacts_dir=artifacts_dir, row_number=row_number, result=result)
    return result

def inspect_product_row(
    row_number: int,
    research_row: ResearchRow,
    artifacts_dir: Path,
    headful: bool = False,
    profile_dir: Path | None = None,
    manual_wait_seconds: int = 0,
) -> InspectResult:
    with sync_playwright() as playwright:
        session = _open_browser_session(playwright=playwright, headful=headful, profile_dir=profile_dir)
        try:
            context = session.context
            page = _fresh_batch_page(context)
            return _inspect_product_row_on_page(
                page=page,
                row_number=row_number,
                research_row=research_row,
                artifacts_dir=artifacts_dir,
                profile_dir=profile_dir,
                manual_wait_seconds=manual_wait_seconds,
                fast_success_exit=False,
            )
        finally:
            session.close()


def build_row_error_result(
    row_number: int,
    research_row: ResearchRow,
    error: Exception,
    profile_dir: Path | None = None,
) -> InspectResult:
    return InspectResult(
        row_number=row_number,
        url=research_row.wb_candidate_url or "",
        final_url=None,
        parse_status="ROW_ERROR",
        anti_bot_detected=False,
        used_persistent_profile=profile_dir is not None,
        profile_dir=str(profile_dir) if profile_dir else None,
        seller_url=None,
        navigated_to_seller_page=False,
        note=f"{type(error).__name__}: {error}",
    )


def _open_browser_session(playwright, headful: bool, profile_dir: Path | None) -> _BrowserSession:
    launch_args = [
        "--disable-blink-features=AutomationControlled",
        "--start-maximized",
    ]
    if profile_dir is not None:
        return _open_system_chrome_session(playwright=playwright, profile_dir=profile_dir)

    browser = playwright.chromium.launch(
        headless=not headful,
        args=launch_args,
    )
    context = browser.new_context(viewport={"width": 1600, "height": 1400})
    return _BrowserSession(context=context, browser=browser)


def _open_system_chrome_session(playwright, profile_dir: Path) -> _BrowserSession:
    chrome_path = _find_system_chrome()
    if chrome_path is None:
        raise RuntimeError("Не найден установленный Google Chrome. Он нужен для прохождения защиты WB.")

    profile_dir.mkdir(parents=True, exist_ok=True)
    debug_port = _reserve_local_port()
    creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    process = subprocess.Popen(
        [
            str(chrome_path),
            f"--remote-debugging-port={debug_port}",
            f"--user-data-dir={profile_dir.resolve()}",
            "--no-first-run",
            "--no-default-browser-check",
            "--start-maximized",
            "about:blank",
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        creationflags=creation_flags,
    )

    try:
        _wait_for_chrome_debug_port(debug_port=debug_port, process=process)
        browser = playwright.chromium.connect_over_cdp(f"http://127.0.0.1:{debug_port}")
        if not browser.contexts:
            raise RuntimeError("Chrome запущен, но профиль браузера недоступен")
        return _BrowserSession(context=browser.contexts[0], browser=browser, process=process)
    except Exception:
        if process.poll() is None:
            _terminate_process_tree(process.pid)
        raise


def _find_system_chrome() -> Path | None:
    candidates = [
        Path(os.environ.get("PROGRAMFILES", "")) / "Google/Chrome/Application/chrome.exe",
        Path(os.environ.get("PROGRAMFILES(X86)", "")) / "Google/Chrome/Application/chrome.exe",
        Path(os.environ.get("LOCALAPPDATA", "")) / "Google/Chrome/Application/chrome.exe",
    ]
    return next((path for path in candidates if path.is_file()), None)


def _reserve_local_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])


def _wait_for_chrome_debug_port(
    debug_port: int,
    process: subprocess.Popen,
    timeout_seconds: float = 15.0,
) -> None:
    deadline = time.monotonic() + timeout_seconds
    endpoint = f"http://127.0.0.1:{debug_port}/json/version"
    last_error: Exception | None = None
    while time.monotonic() < deadline:
        if process.poll() is not None:
            raise RuntimeError(f"Google Chrome завершился при запуске, код {process.returncode}")
        try:
            with urlopen(endpoint, timeout=1) as response:
                payload = json.loads(response.read().decode("utf-8"))
            if payload.get("webSocketDebuggerUrl"):
                return
        except Exception as exc:
            last_error = exc
            time.sleep(0.25)
    raise RuntimeError(f"Не удалось подключиться к Google Chrome: {last_error}")


def _terminate_process_tree(pid: int) -> None:
    try:
        subprocess.run(
            ["taskkill", "/PID", str(pid), "/T", "/F"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=10,
            check=False,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except Exception:
        pass


def _go_to_seller_page(
    page: Page,
    expected_seller_name: str | None = None,
    ready_timeout_ms: int = PRODUCT_READY_TIMEOUT_MS,
) -> tuple[str | None, Response | None, bool]:
    if "/seller/" in page.url:
        raise TransientWBError(f'Page is still on seller URL before seller navigation: {page.url}')

    discovered_seller_url = _wait_for_seller_link_or_page(page, allow_slug=True)
    if discovered_seller_url is not None:
        if _click_seller_link(page, discovered_seller_url):
            ready = _wait_for_seller_page_ready(page, expected_seller_name, timeout_ms=ready_timeout_ms)
            return discovered_seller_url, None, ready
        try:
            response = page.goto(
                discovered_seller_url,
                wait_until="commit",
                timeout=SELLER_GOTO_TIMEOUT_MS,
            )
            ready = _wait_for_seller_page_ready(page, expected_seller_name, timeout_ms=ready_timeout_ms)
            return discovered_seller_url, response, ready
        except Exception:
            pass

    for selector in PRODUCT_SELLER_LINK_SELECTORS:
        try:
            locator = page.locator(selector)
            count = locator.count()
            if count == 0:
                continue
            allow_slug = _selector_allows_slug_seller_url(selector)
            for index in range(min(count, 12)):
                candidate = locator.nth(index)
                href = candidate.get_attribute("href")
                target_url = _normalize_candidate_seller_url(page.url, href, allow_slug=allow_slug)
                if target_url is not None:
                    response = page.goto(
                        target_url,
                        wait_until="commit",
                        timeout=SELLER_GOTO_TIMEOUT_MS,
                    )
                    ready = _wait_for_seller_page_ready(page, expected_seller_name, timeout_ms=ready_timeout_ms)
                    return target_url, response, ready

                try:
                    before_url = page.url
                    response = None
                    try:
                        with page.expect_navigation(wait_until="domcontentloaded", timeout=WAIT_FOR_LOAD_STATE_TIMEOUT_MS) as nav:
                            candidate.click(timeout=600)
                        response = nav.value
                    except Exception:
                        try:
                            candidate.click(timeout=600)
                        except Exception:
                            continue

                    navigated_url = _wait_for_seller_link_or_page(page, previous_url=before_url, allow_slug=allow_slug)
                    if navigated_url is None:
                        try:
                            page.go_back(wait_until="domcontentloaded", timeout=WAIT_FOR_LOAD_STATE_TIMEOUT_MS)
                            _best_effort_wait(page, settle_rounds=2)
                        except Exception:
                            pass
                        continue

                    if _is_valid_wb_seller_url(navigated_url) or (allow_slug and _is_supported_wb_seller_url(navigated_url)):
                        ready = _wait_for_seller_page_ready(page, expected_seller_name, timeout_ms=ready_timeout_ms)
                        return navigated_url, response, ready

                    try:
                        page.go_back(wait_until="domcontentloaded", timeout=WAIT_FOR_LOAD_STATE_TIMEOUT_MS)
                        _best_effort_wait(page, settle_rounds=2)
                    except Exception:
                        pass
                except Exception:
                    continue
        except Exception:
            continue
    return None, None, False


def _click_seller_link(page: Page, target_url: str) -> bool:
    for selector in PRODUCT_SELLER_LINK_SELECTORS:
        try:
            locator = page.locator(selector)
            count = locator.count()
        except Exception:
            continue
        for index in range(min(count, 12)):
            candidate = locator.nth(index)
            try:
                candidate_url = _normalize_candidate_seller_url(
                    page.url,
                    candidate.get_attribute("href"),
                    allow_slug=_selector_allows_slug_seller_url(selector),
                )
                if candidate_url != target_url:
                    continue
                candidate.scroll_into_view_if_needed(timeout=2_000)
                candidate.click(timeout=3_000, no_wait_after=True)
                deadline = time.monotonic() + 8
                while time.monotonic() < deadline:
                    if _is_supported_wb_seller_url(page.url):
                        return True
                    page.wait_for_timeout(250)
            except Exception:
                continue
    return False




def _reveal_supplier_requisites(page: Page) -> str | None:
    return _reveal_supplier_requisites_with_strategy(
        page=page,
        max_rounds=BASE_TOOLTIP_REVEAL_ROUNDS,
        pause_ms=BASE_TOOLTIP_PAUSE_MS,
        deep_mode=False,
    )


def _reveal_supplier_requisites_with_strategy(
    page: Page,
    max_rounds: int,
    pause_ms: int,
    deep_mode: bool,
) -> str | None:
    dom_text = _extract_requisites_text_via_dom(page, include_body_scan=deep_mode)
    if _contains_seller_signal(dom_text or ""):
        return dom_text

    captured_tooltip_text = _extract_tooltip_text_from_page(page)
    if _contains_seller_signal(captured_tooltip_text or ""):
        return captured_tooltip_text

    if _page_has_empty_results_state(page):
        return captured_tooltip_text or dom_text

    if not _page_has_requisites_entrypoints(page, deep_mode=deep_mode):
        return captured_tooltip_text

    for _ in range(max_rounds):
        triggered_text = _trigger_supplier_tooltip(page, deep_mode=deep_mode)
        if _contains_seller_signal(triggered_text or ""):
            return triggered_text
        _best_effort_wait(page, include_networkidle=deep_mode, settle_rounds=1 if not deep_mode else 2)

        dom_text = _extract_requisites_text_via_dom(page, include_body_scan=deep_mode)
        if _contains_seller_signal(dom_text or ""):
            return dom_text

        captured_tooltip_text = captured_tooltip_text or _extract_tooltip_text_from_page(page)
        if _contains_seller_signal(captured_tooltip_text or ""):
            return captured_tooltip_text

        html = _safe_page_content(page)
        html_tooltip = _extract_tooltip_text_from_html(html)
        if _contains_seller_signal(html_tooltip or ""):
            return html_tooltip

        text = _safe_page_text(page)
        if _contains_seller_signal(text):
            return text

        try:
            page.wait_for_timeout(pause_ms)
        except Exception:
            break
    return captured_tooltip_text


def _trigger_supplier_tooltip(page: Page, deep_mode: bool = False) -> str | None:
    if _page_has_empty_results_state(page):
        return None

    selectors = list(SELLER_TOOLTIP_TRIGGER_SELECTORS)
    if deep_mode:
        selectors.extend(DEEP_SELLER_TOOLTIP_TRIGGER_SELECTORS)

    for selector in selectors:
        candidates = _selector_candidates(page, selector)
        if not candidates:
            continue
        for locator in candidates:
            try:
                locator.scroll_into_view_if_needed(timeout=375)
                if selector in SAFE_CLICK_SELLER_TOOLTIP_TRIGGER_SELECTORS:
                    if _safe_click_seller_tooltip_trigger(page, locator):
                        requisites_text = _wait_for_requisites_text(page)
                        if requisites_text:
                            return requisites_text

                _hover_like_human(page, locator)
                try:
                    locator.focus(timeout=250)
                except Exception:
                    pass
                _dispatch_tooltip_events(page, locator)
                _force_tooltip_open_via_js(locator)
                requisites_text = _wait_for_requisites_text(page)
                if requisites_text:
                    return requisites_text
            except Exception:
                continue
    return None


def _wait_for_requisites_text(page: Page, timeout_ms: int = REQUISITES_APPEAR_TIMEOUT_MS) -> str | None:
    deadline = time.monotonic() + max(timeout_ms, 0) / 1_000
    while time.monotonic() < deadline:
        requisites_text = _extract_requisites_text_via_dom(page, include_body_scan=False)
        if _contains_seller_signal(requisites_text or ""):
            return requisites_text
        try:
            page.wait_for_timeout(100)
        except Exception:
            break
    return None


def _safe_click_seller_tooltip_trigger(page: Page, locator: Locator) -> bool:
    if not _is_safe_seller_tooltip_click_target(locator):
        return False

    try:
        locator.click(timeout=450, force=True, no_wait_after=True)
    except Exception:
        try:
            locator.dispatch_event('click')
        except Exception:
            return False

    try:
        page.wait_for_timeout(90)
    except Exception:
        pass
    return True


def _is_safe_seller_tooltip_click_target(locator: Locator) -> bool:
    try:
        return bool(
            locator.evaluate(
                """
                (node) => {
                  if (!node) return false;
                  const safeFragments = ['seller-details', 'seller-info', 'sellerinfo', 'sellerheader', 'catalog-page__seller-details', 'tip-info', 'tooltip', 'supplier'];
                  const dangerFragments = ['basket', 'cart', 'product-card', 'add-to-basket', 'favorites', 'postpone', 'orderwrap', 'buy'];

                  let current = node;
                  for (let depth = 0; current && depth < 6; depth += 1, current = current.parentElement) {
                    const parts = [
                      current.className || '',
                      current.id || '',
                      current.getAttribute?.('data-testid') || '',
                      current.getAttribute?.('aria-label') || '',
                      current.getAttribute?.('href') || '',
                    ].join(' ').toLowerCase();

                    if (dangerFragments.some((fragment) => parts.includes(fragment))) {
                      return false;
                    }
                    if (safeFragments.some((fragment) => parts.includes(fragment))) {
                      return true;
                    }
                  }
                  return false;
                }
                """
            )
        )
    except Exception:
        return False


def _force_tooltip_open_via_js(locator: Locator) -> None:
    try:
        locator.evaluate(
            """
            (node) => {
              if (!node) return;
              const mouseEvents = ['mouseenter', 'mouseover', 'mousemove'];
              const pointerEvents = ['pointerenter', 'pointerover'];

              for (const eventName of mouseEvents) {
                node.dispatchEvent(new MouseEvent(eventName, { bubbles: true, cancelable: true, composed: true, view: window }));
              }
              for (const eventName of pointerEvents) {
                if (typeof PointerEvent === 'function') {
                  node.dispatchEvent(new PointerEvent(eventName, { bubbles: true, cancelable: true, composed: true, view: window }));
                }
              }
              if (node instanceof HTMLElement) {
                node.focus?.();
              }
            }
            """
        )
    except Exception:
        return


def _hover_like_human(page: Page, locator: Locator) -> None:
    try:
        box = locator.bounding_box()
        if box:
            x = box["x"] + box["width"] / 2
            y = box["y"] + box["height"] / 2
            page.mouse.move(x - 8, y - 6)
            page.wait_for_timeout(30)
            page.mouse.move(x, y, steps=6)
            page.wait_for_timeout(60)
    except Exception:
        pass
    try:
        locator.hover(timeout=375, force=True)
        page.wait_for_timeout(90)
    except Exception:
        pass


def _dispatch_tooltip_events(page: Page, locator: Locator) -> None:
    try:
        locator.dispatch_event('mouseenter')
    except Exception:
        pass
    try:
        locator.dispatch_event('mouseover')
    except Exception:
        pass
    try:
        locator.dispatch_event('pointerenter')
    except Exception:
        pass
    try:
        page.wait_for_timeout(60)
    except Exception:
        return


def _tooltip_visible(page: Page) -> bool:
    for selector in TOOLTIP_VISIBLE_SELECTORS:
        try:
            locator = page.locator(selector).first
            if locator.count() and locator.is_visible(timeout=500):
                return True
        except Exception:
            continue
    return False


def _page_has_requisites_entrypoints(page: Page, deep_mode: bool) -> bool:
    if _page_has_empty_results_state(page):
        return False

    for selector in SAFE_CLICK_SELLER_TOOLTIP_TRIGGER_SELECTORS:
        try:
            if page.locator(selector).count() > 0:
                return True
        except Exception:
            continue

    selectors = list(SELLER_TOOLTIP_TRIGGER_SELECTORS)
    if deep_mode:
        selectors.extend(DEEP_SELLER_TOOLTIP_TRIGGER_SELECTORS)
    selectors.extend(TOOLTIP_VISIBLE_SELECTORS)

    for selector in selectors:
        for locator in _selector_candidates(page, selector):
            try:
                if locator.is_visible(timeout=250):
                    return True
            except Exception:
                continue
    return False


def _build_result(
    row_number: int,
    url: str,
    page: Page,
    http_status: int | None,
    html: str,
    text: str,
    captured_tooltip_text: str | None,
    screenshot_path: Path | None,
    html_path: Path | None,
    text_path: Path | None,
    used_persistent_profile: bool,
    profile_dir: Path | None,
    manual_wait_seconds: int,
    seller_url: str | None,
    navigated_to_seller_page: bool,
) -> InspectResult:
    page_title = _safe_page_title(page)
    tooltip_text = captured_tooltip_text or _extract_tooltip_text_from_page(page) or _extract_tooltip_text_from_html(html)
    page_has_requisites_entrypoints = _page_has_requisites_entrypoints(page, deep_mode=False)
    seller_country = _detect_seller_country(tooltip_text or "", html, text)
    empty_results_detected = _contains_empty_results_text(page_title or "", text, tooltip_text or "")

    inn = _first_non_empty(
        _first_match(INN_RE, tooltip_text or ""),
        _first_match(INN_RE, html),
        _first_match(INN_RE, text),
        _first_match(INN_RE, page_title or ""),
    )
    ogrn = _first_non_empty(
        _first_match(OGRN_RE, tooltip_text or ""),
        _first_match(OGRN_RE, html),
        _first_match(OGRN_RE, text),
    )
    ogrnip = _first_non_empty(
        _first_match(OGRNIP_RE, tooltip_text or ""),
        _first_match(OGRNIP_RE, html),
        _first_match(OGRNIP_RE, text),
    )

    entity_type = _first_non_empty(
        _extract_entity_type(tooltip_text or ""),
        _extract_entity_type(html),
        _extract_entity_type(text),
    )
    seller_display_name = None
    if not empty_results_detected:
        seller_display_name = _first_non_empty(
            _extract_seller_display_name_from_html(html),
            _extract_seller_display_name(tooltip_text or ""),
            _extract_seller_display_name(text) if navigated_to_seller_page else None,
        )

    combined_text = "\n".join(filter(None, [page_title or "", text, tooltip_text or "", inn or "", ogrn or "", ogrnip or "", entity_type or "", seller_display_name or ""]))
    anti_bot_detected = _contains_anti_bot_text(combined_text + "\n" + html[:5000], http_status)

    parse_status, note = _detect_parse_status(
        http_status=http_status,
        anti_bot_detected=anti_bot_detected,
        inn=inn,
        entity_type=entity_type,
        seller_country=seller_country,
        manual_wait_seconds=manual_wait_seconds,
        navigated_to_seller_page=navigated_to_seller_page,
        page_has_requisites_entrypoints=page_has_requisites_entrypoints,
        empty_results_detected=empty_results_detected,
    )

    return InspectResult(
        row_number=row_number,
        url=url,
        page_title=page_title,
        final_url=page.url,
        http_status=http_status,
        parse_status=parse_status,
        content_text_length=len(text),
        anti_bot_detected=anti_bot_detected,
        used_persistent_profile=used_persistent_profile,
        profile_dir=str(profile_dir) if profile_dir else None,
        manual_wait_seconds=manual_wait_seconds,
        seller_url=seller_url,
        navigated_to_seller_page=navigated_to_seller_page,
        inn=inn,
        ogrn=ogrn,
        ogrnip=ogrnip,
        entity_type=entity_type,
        seller_display_name=seller_display_name,
        note=note,
        screenshot_path=str(screenshot_path) if screenshot_path else None,
        html_path=str(html_path) if html_path else None,
        text_path=str(text_path) if text_path else None,
    )


def _first_non_empty(*values: str | None) -> str | None:
    for value in values:
        if value:
            return value
    return None


def _extract_seller_display_name_from_html(html: str) -> str | None:
    tooltip_text = _extract_tooltip_text_from_html(html)
    if tooltip_text:
        name = _extract_seller_display_name(tooltip_text)
        if name:
            return name

    for match in SELLER_DETAILS_TITLE_RE.finditer(html):
        raw = TAG_RE.sub(' ', match.group('name'))
        normalized = _normalize_text(raw)
        if normalized and _is_plausible_seller_display_name(normalized):
            return normalized[:120]

    for match in SELLER_HEADER_RE.finditer(html):
        raw = TAG_RE.sub(' ', match.group('name'))
        normalized = _normalize_text(raw)
        if normalized and 'Wildberries' not in normalized and _is_plausible_seller_display_name(normalized):
            return normalized[:120]
    return None


def _detect_parse_status(
    http_status: int | None,
    anti_bot_detected: bool,
    inn: str | None,
    entity_type: str | None,
    seller_country: str | None,
    manual_wait_seconds: int,
    navigated_to_seller_page: bool,
    page_has_requisites_entrypoints: bool,
    empty_results_detected: bool,
) -> tuple[str, str]:
    requisites_found = bool(inn or entity_type)

    if anti_bot_detected and requisites_found:
        return "PARTIAL_SUCCESS", "WB показал защиту, но часть реквизитов всё же извлечена"

    if anti_bot_detected:
        if manual_wait_seconds > 0:
            return "MANUAL_CHECK_REQUIRED", "WB показал антибот-страницу даже после ручной паузы"
        return "ANTI_BOT_PAGE", "WB показал антибот-страницу вместо карточки товара"

    if http_status and http_status >= 400 and requisites_found:
        return "PARTIAL_SUCCESS", f"HTTP status {http_status}, но часть реквизитов найдена"

    if http_status and http_status >= 400:
        return "PAGE_NOT_AVAILABLE", f"HTTP status {http_status}"

    if inn:
        return "SUCCESS", "ИНН найден на странице продавца" if navigated_to_seller_page else "ИНН найден в тексте страницы"

    if seller_country:
        return seller_country, f"Обнаружен продавец из страны: {seller_country}"

    if entity_type:
        return "NEEDS_REVIEW", "Есть признаки реквизитов продавца, но ИНН не найден"

    if empty_results_detected:
        return "NOTHING_FOUND_PAGE", "WB открыл пустую страницу с сообщением 'ничего не найдено'"

    if navigated_to_seller_page and page_has_requisites_entrypoints:
        return "SELLER_PAGE_OPENED", "Перешли на страницу продавца, но реквизиты не извлеклись"

    if navigated_to_seller_page and not page_has_requisites_entrypoints:
        return "Нет реквизитов на странице", "Нет реквизитов на странице"

    if manual_wait_seconds > 0:
        return "PRODUCT_PAGE_OPENED", "Карточка открылась после ручной сессии, но реквизиты пока не извлечены"

    return "PAGE_OPENED_NO_REQUISITES", "Страница открылась, но реквизиты продавца не найдены"




def _contains_requisites_text(text: str) -> bool:
    return any(marker in text for marker in ('ИНН', 'ОГРН', 'ОГРНИП', 'Номер регистрации', 'КПП'))


def _contains_empty_results_text(*texts: str | None) -> bool:
    combined_text = _normalize_text("\n".join(filter(None, texts))).lower()
    if not combined_text:
        return False
    return any(pattern in combined_text for pattern in EMPTY_RESULTS_PATTERNS)


def _contains_seller_signal(text: str) -> bool:
    return bool(_contains_requisites_text(text) or _detect_seller_country(text) or _extract_entity_type(text))


def _detect_seller_country(*texts: str) -> str | None:
    combined_text = _normalize_text("\n".join(filter(None, texts)))
    if not combined_text:
        return None
    if BELARUS_RE.search(combined_text):
        return "БЕЛОРУСЬ"
    if KAZAKHSTAN_RE.search(combined_text):
        return "КАЗАХСТАН"
    return None


def _contains_anti_bot_text(text: str, http_status: int | None) -> bool:
    text_lower = text.lower()
    if http_status == 498:
        return True
    return any(pattern.lower() in text_lower for pattern in ANTI_BOT_PATTERNS)


def _extract_tooltip_text_from_page(page: Page) -> str | None:
    dom_text = _extract_requisites_text_via_dom(page, include_body_scan=False)
    if _contains_seller_signal(dom_text or ""):
        return dom_text

    tooltip_selectors = list(TOOLTIP_VISIBLE_SELECTORS)
    tooltip_selectors.extend(['.tooltip__content', '[role="tooltip"]', '[class*="tooltip"]'])

    for selector in tooltip_selectors:
        try:
            for locator in _selector_candidates(page, selector):
                text = locator.inner_text(timeout=500)
                normalized = _normalize_text(text)
                if normalized and _contains_seller_signal(normalized):
                    return normalized
        except Exception:
            continue
    try:
        text = page.locator('.tooltip__content').last.inner_text(timeout=375)
        normalized = _normalize_text(text)
        if normalized and _contains_seller_signal(normalized):
            return normalized
    except Exception:
        pass
    return None


def _extract_tooltip_text_from_html(html: str) -> str | None:
    block_candidates: list[str] = []

    for match in TOOLTIP_BLOCK_RE.finditer(html):
        block_candidates.append(match.group('body'))
    for match in TOOLTIP_TEXT_RE.finditer(html):
        block_candidates.append(match.group('body'))

    for raw_body in block_candidates:
        if not _contains_seller_signal(raw_body):
            continue
        text = TAG_RE.sub(' ', raw_body)
        text = text.replace('&nbsp;', ' ')
        normalized = _normalize_text(text)
        if normalized:
            return normalized

    if _contains_seller_signal(html):
        index_candidates = [idx for marker in ('ИНН', 'ОГРН', 'ОГРНИП', 'КПП', 'Номер регистрации', 'УНП', 'БИН', 'ИИН', 'Республика Беларусь', 'Казахстан') if (idx := html.find(marker)) != -1]
        if index_candidates:
            start_idx = max(0, min(index_candidates) - 800)
            end_idx = min(len(html), max(index_candidates) + 1200)
            snippet = html[start_idx:end_idx]
            text = TAG_RE.sub(' ', snippet).replace('&nbsp;', ' ')
            normalized = _normalize_text(text)
            if normalized:
                return normalized
    return None


def _extract_entity_type(text: str) -> str | None:
    full = _first_match(ENTITY_FULL_RE, text)
    if full:
        if full.lower().startswith("индивидуальный"):
            return "ИП"
        if full.lower().startswith("общество"):
            return "ООО"
    return _first_match(ENTITY_SHORT_RE, text)


def _extract_seller_display_name(text: str) -> str | None:
    if not text:
        return None

    normalized_text = _normalize_text(text)
    for line in normalized_text.splitlines():
        if not line:
            continue
        if any(marker in line for marker in ["ИНН", "ОГРН", "ОГРНИП", "КПП", "Номер регистрации", "Интернет-магазин Wildberries"]):
            continue
        cleaned = re.sub(r"\b(ИП|ООО)\b\s*$", "", line).strip(" ,")
        if _is_plausible_seller_display_name(cleaned):
            return cleaned[:120]

    match = SELLER_NAME_RE.search(text)
    if not match:
        return None
    value = " ".join(match.group(1).split())
    return value[:120] if _is_plausible_seller_display_name(value) else None


def _has_extracted_requisites(result: InspectResult) -> bool:
    return bool(result.inn or result.entity_type or result.ogrn or result.ogrnip)


def _selector_candidates(page: Page, selector: str, max_items: int = 6) -> list[Locator]:
    try:
        locator = page.locator(selector)
        count = min(locator.count(), max_items)
    except Exception:
        return []

    visible_candidates: list[Locator] = []
    fallback_candidates: list[Locator] = []
    for index in range(count):
        candidate = locator.nth(index)
        try:
            if candidate.is_visible(timeout=250):
                visible_candidates.append(candidate)
            else:
                fallback_candidates.append(candidate)
        except Exception:
            fallback_candidates.append(candidate)
    return visible_candidates or fallback_candidates


def _result_information_score(result: InspectResult) -> int:
    score = 0
    if result.inn:
        score += 100
    if result.ogrn or result.ogrnip:
        score += 40
    if result.entity_type:
        score += 20
    if result.parse_status in {"БЕЛОРУСЬ", "КАЗАХСТАН"}:
        score += 25
    if result.seller_display_name:
        score += 5
    if result.note:
        score += 2
    return score


def _write_result_json(artifacts_dir: Path, row_number: int, result: InspectResult) -> None:
    (artifacts_dir / f"row_{row_number}.json").write_text(
        json.dumps(result.model_dump(mode='json'), ensure_ascii=False, indent=2),
        encoding='utf-8',
    )


def _best_effort_wait(page: Page, include_networkidle: bool = False, settle_rounds: int = 2) -> None:
    states = ["domcontentloaded", "load"]
    if include_networkidle:
        states.append("networkidle")

    for state in states:
        try:
            page.wait_for_load_state(state, timeout=WAIT_FOR_LOAD_STATE_TIMEOUT_MS)
        except Exception:
            continue
    for _ in range(settle_rounds):
        try:
            page.wait_for_timeout(60)
        except Exception:
            break


def _safe_page_content(page: Page) -> str:
    last_error: Exception | None = None
    for _ in range(8):
        try:
            _best_effort_wait(page)
            return page.content()
        except Exception as exc:
            last_error = exc
            try:
                page.wait_for_timeout(500)
            except Exception:
                break
    if last_error is not None:
        return f"<!-- page.content failed: {last_error} -->"
    return ""


def _safe_page_text(page: Page) -> str:
    for _ in range(6):
        try:
            return page.locator("body").inner_text(timeout=4_500)
        except Exception:
            try:
                page.wait_for_timeout(90)
            except Exception:
                break
    return ""


def _safe_page_title(page: Page) -> str | None:
    try:
        return page.title()
    except Exception:
        return None


def _first_match(pattern: re.Pattern[str], text: str) -> str | None:
    match = pattern.search(text)
    return match.group(1) if match else None


def _normalize_text(text: str) -> str:
    lines = []
    for line in text.splitlines():
        cleaned = SPACE_RE.sub(" ", line).strip()
        if cleaned:
            lines.append(cleaned)
    return "\n".join(lines)


def _ensure_not_stuck_on_previous_seller_page(
    page: Page,
    research_row: ResearchRow,
    previous_url: str | None,
) -> None:
    current_url = page.url or ""
    expected_product_marker = f"/catalog/{research_row.wb_nm_id}/" if research_row.wb_nm_id is not None else None

    if expected_product_marker and expected_product_marker in current_url:
        return

    if "/seller/" not in current_url:
        return

    if previous_url and previous_url == current_url:
        raise TransientWBError(
            f"Navigation stayed on previous seller page instead of product page: {current_url}"
        )

    raise TransientWBError(
        f"Expected product page for nmID {research_row.wb_nm_id}, got seller page: {current_url}"
    )


def _normalize_candidate_seller_url(base_url: str, href: str | None, allow_slug: bool = False) -> str | None:
    if not href:
        return None
    target_url = urljoin(base_url, href)
    if _is_valid_wb_seller_url(target_url):
        return target_url
    if allow_slug and _is_supported_wb_seller_url(target_url):
        return target_url
    return None


def _wait_for_seller_link_or_page(page: Page, previous_url: str | None = None, allow_slug: bool = False) -> str | None:
    for _ in range(SELLER_LINK_DISCOVERY_ROUNDS):
        current_url = page.url or ""
        if _is_valid_wb_seller_url(current_url) or (allow_slug and _is_supported_wb_seller_url(current_url)):
            return current_url

        discovered_url = _discover_seller_url_on_page(page, allow_slug=allow_slug)
        if discovered_url is not None:
            return discovered_url

        if previous_url and current_url and current_url != previous_url and (
            _is_valid_wb_seller_url(current_url) or (allow_slug and _is_supported_wb_seller_url(current_url))
        ):
            return current_url

        try:
            page.wait_for_timeout(SELLER_LINK_DISCOVERY_PAUSE_MS)
        except Exception:
            break
    return None


def _discover_seller_url_on_page(page: Page, allow_slug: bool = False) -> str | None:
    for selector in PRODUCT_SELLER_LINK_SELECTORS:
        allow_slug_for_selector = allow_slug and _selector_allows_slug_seller_url(selector)
        try:
            locator = page.locator(selector)
            count = locator.count()
        except Exception:
            continue

        for index in range(min(count, 12)):
            try:
                href = locator.nth(index).get_attribute("href")
            except Exception:
                continue
            target_url = _normalize_candidate_seller_url(page.url, href, allow_slug=allow_slug_for_selector)
            if target_url is not None:
                return target_url
    return None


def _selector_allows_slug_seller_url(selector: str) -> bool:
    return selector == PRODUCT_SELLER_LINK_SELECTORS[0]


def _is_supported_wb_seller_url(url: str | None) -> bool:
    if not url:
        return False
    try:
        parsed = urlparse(url)
    except Exception:
        return False

    host = (parsed.netloc or '').lower()
    if host and 'wildberries.ru' not in host:
        return False

    path = parsed.path or ''
    return bool(WB_GENERIC_SELLER_PATH_RE.fullmatch(path))


def _is_valid_wb_seller_url(url: str | None) -> bool:
    if not url:
        return False
    try:
        parsed = urlparse(url)
    except Exception:
        return False

    host = (parsed.netloc or '').lower()
    if host and 'wildberries.ru' not in host:
        return False

    path = parsed.path or ''
    return bool(WB_NUMERIC_SELLER_PATH_RE.fullmatch(path))


def _is_plausible_seller_display_name(value: str | None) -> bool:
    if not value:
        return False

    normalized = _normalize_text(value).strip(" ,")
    if len(normalized) < 2:
        return False

    if normalized.lower() in IGNORED_SELLER_DISPLAY_NAMES:
        return False

    if _contains_empty_results_text(normalized):
        return False

    if re.fullmatch(r"\d+[.,]\d+", normalized):
        return False

    if re.fullmatch(r"[\d\s.,%?]+", normalized):
        return False

    if not re.search(r"[A-Za-zА-Яа-яЁё]", normalized):
        return False

    return True


def _should_run_deep_recovery(result: InspectResult) -> bool:
    return result.parse_status in {
        "SELLER_PAGE_OPENED",
        "PAGE_OPENED_NO_REQUISITES",
        "NEEDS_REVIEW",
        "Нет реквизитов на странице",
    }


def _page_has_empty_results_state(page: Page) -> bool:
    try:
        return bool(
            page.evaluate(
                """
                (patterns) => {
                  const text = (document.body?.innerText || '').toLowerCase();
                  return patterns.some((pattern) => text.includes(pattern));
                }
                """,
                EMPTY_RESULTS_PATTERNS,
            )
        )
    except Exception:
        try:
            body_text = page.locator("body").inner_text(timeout=1_500)
        except Exception:
            return False
        return _contains_empty_results_text(body_text)


def _run_deep_requisites_recovery(
    page: Page,
    row_number: int,
    research_row: ResearchRow,
    screenshot_path: Path,
    html_path: Path,
    text_path: Path,
    used_persistent_profile: bool,
    profile_dir: Path | None,
    manual_wait_seconds: int,
    seller_url: str | None,
    initial_http_status: int | None,
) -> InspectResult:
    http_status = initial_http_status

    if seller_url is None and not _is_supported_wb_seller_url(page.url):
        discovered_seller_url, response, _ = _go_to_seller_page(
            page,
            expected_seller_name=research_row.seller_name_raw,
        )
        if discovered_seller_url is not None:
            seller_url = discovered_seller_url
            http_status = response.status if response else http_status
    elif seller_url is None and _is_supported_wb_seller_url(page.url):
        seller_url = page.url

    if seller_url and page.url.rstrip("/") != seller_url.rstrip("/"):
        try:
            response = page.goto(
                seller_url,
                wait_until="commit",
                timeout=SELLER_GOTO_TIMEOUT_MS,
            )
            http_status = response.status if response else http_status
            _wait_for_seller_page_ready(page, research_row.seller_name_raw)
        except Exception:
            pass

    captured_tooltip_text = _reveal_supplier_requisites_with_strategy(
        page=page,
        max_rounds=DEEP_TOOLTIP_REVEAL_ROUNDS,
        pause_ms=DEEP_TOOLTIP_PAUSE_MS,
        deep_mode=True,
    )

    html = _safe_page_content(page)
    text = _safe_page_text(page)

    page.screenshot(path=str(screenshot_path), full_page=True)
    html_path.write_text(html, encoding='utf-8')
    text_path.write_text(text, encoding='utf-8')

    return _build_result(
        row_number=row_number,
        url=research_row.wb_candidate_url,
        page=page,
        http_status=http_status,
        html=html,
        text=text,
        captured_tooltip_text=captured_tooltip_text,
        screenshot_path=screenshot_path,
        html_path=html_path,
        text_path=text_path,
        used_persistent_profile=used_persistent_profile,
        profile_dir=profile_dir,
        manual_wait_seconds=manual_wait_seconds,
        seller_url=seller_url,
        navigated_to_seller_page=bool(seller_url) or '/seller/' in page.url,
    )


