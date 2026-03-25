from __future__ import annotations

import json
import re
import time
from pathlib import Path
from urllib.parse import urljoin

from playwright.sync_api import BrowserContext, Locator, Page, Response, sync_playwright

from .models import InspectResult, ResearchRow

INN_RE = re.compile(r"\b(?:ИНН)\s*[:№]?\s*(\d{10}|\d{12})\b", re.IGNORECASE)
OGRN_RE = re.compile(r"\b(?:ОГРН)\s*[:№]?\s*(\d{13})\b", re.IGNORECASE)
OGRNIP_RE = re.compile(r"\b(?:ОГРНИП)\s*[:№]?\s*(\d{15})\b", re.IGNORECASE)
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
TOOLTIP_TEXT_RE = re.compile(
    r"<div[^>]*class=\"[^\"]*tooltip__content[^\"]*\"[^>]*>(?P<body>.*?)</div>",
    re.IGNORECASE | re.DOTALL,
)
TAG_RE = re.compile(r"<[^>]+>")
SPACE_RE = re.compile(r"\s+")

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
    '.seller-details__parameter-value .seller-details__tip',
    '.seller-details__parameter-value .tip-question',
    '.seller-details__info-icon',
    '.seller-info__info-icon',
    '.seller-info i',
    '.seller-rating__ico',
    '[class*="sellerInfoRatingIcon"]',
    '[class*="sellerInfoNameDefault"] + div',
    '[class*="sellerInfo"] [class*="icon"]',
]

TOOLTIP_VISIBLE_SELECTORS = [
    ".tooltip.tooltip-supplier",
    '[class*="tooltip-supplier"]',
]


def _extract_requisites_text_via_dom(page: Page) -> str | None:
    try:
        text = page.evaluate(
            """
            () => {
              const markers = ['ИНН', 'ОГРН', 'ОГРНИП', 'Номер регистрации', 'КПП'];
              const selectors = [
                '.tooltip.tooltip-supplier',
                '[class*="tooltip-supplier"]',
                '.tooltip__content',
                '.seller-details',
                'body',
              ];

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

              return candidates[0] || null;
            }
            """
        )
    except Exception:
        return None

    normalized = _normalize_text(text or "")
    return normalized or None


def inspect_product_row(
    row_number: int,
    research_row: ResearchRow,
    artifacts_dir: Path,
    headful: bool = False,
    profile_dir: Path | None = None,
    manual_wait_seconds: int = 0,
) -> InspectResult:
    if not research_row.wb_candidate_url:
        raise ValueError("У строки нет wb_candidate_url")

    artifacts_dir.mkdir(parents=True, exist_ok=True)
    screenshot_path = artifacts_dir / f"row_{row_number}.png"
    html_path = artifacts_dir / f"row_{row_number}.html"
    text_path = artifacts_dir / f"row_{row_number}_text.txt"

    with sync_playwright() as playwright:
        context = _open_context(playwright=playwright, headful=headful, profile_dir=profile_dir)
        try:
            page = context.pages[0] if context.pages else context.new_page()
            response = page.goto(research_row.wb_candidate_url, wait_until="domcontentloaded", timeout=60_000)
            _best_effort_wait(page)

            seller_url, seller_response = _go_to_seller_page(page)
            if seller_response is not None:
                response = seller_response

            captured_tooltip_text = _reveal_supplier_requisites(page)

            if manual_wait_seconds > 0:
                time.sleep(manual_wait_seconds)
                captured_tooltip_text = captured_tooltip_text or _reveal_supplier_requisites(page)

            html_before_screenshot = _safe_page_content(page)
            captured_tooltip_text = captured_tooltip_text or _extract_tooltip_text_from_page(page) or _extract_tooltip_text_from_html(html_before_screenshot)

            page.screenshot(path=str(screenshot_path), full_page=True)
            html = _safe_page_content(page)
            html_path.write_text(html, encoding="utf-8")
            text = _safe_page_text(page)
            text_path.write_text(text, encoding="utf-8")

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
                navigated_to_seller_page=bool(seller_url) or "/seller/" in page.url,
            )
            (artifacts_dir / f"row_{row_number}.json").write_text(
                json.dumps(result.model_dump(mode="json"), ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            return result
        finally:
            context.close()


def _open_context(playwright, headful: bool, profile_dir: Path | None) -> BrowserContext:
    launch_args = [
        "--disable-blink-features=AutomationControlled",
        "--start-maximized",
    ]
    if profile_dir is not None:
        profile_dir.mkdir(parents=True, exist_ok=True)
        return playwright.chromium.launch_persistent_context(
            user_data_dir=str(profile_dir),
            headless=False,
            viewport={"width": 1600, "height": 1400},
            args=launch_args,
        )

    browser = playwright.chromium.launch(
        headless=not headful,
        args=launch_args,
    )
    return browser.new_context(viewport={"width": 1600, "height": 1400})


def _go_to_seller_page(page: Page) -> tuple[str | None, Response | None]:
    if "/seller/" in page.url:
        return page.url, None

    for selector in PRODUCT_SELLER_LINK_SELECTORS:
        try:
            locator = page.locator(selector).first
            if locator.count() == 0:
                continue
            href = locator.get_attribute("href")
            target_url = urljoin(page.url, href) if href else None
            if target_url and "/seller/" in target_url:
                response = page.goto(target_url, wait_until="domcontentloaded", timeout=30_000)
                _best_effort_wait(page)
                return target_url, response

            try:
                with page.expect_navigation(wait_until="domcontentloaded", timeout=20_000) as nav:
                    locator.click(timeout=5_000)
                response = nav.value
                _best_effort_wait(page)
                return page.url, response
            except Exception:
                continue
        except Exception:
            continue
    return None, None




def _reveal_supplier_requisites(page: Page) -> str | None:
    captured_tooltip_text: str | None = None
    for _ in range(8):
        _trigger_supplier_tooltip(page)
        _best_effort_wait(page)

        dom_text = _extract_requisites_text_via_dom(page)
        if _contains_requisites_text(dom_text or ""):
            return dom_text

        captured_tooltip_text = captured_tooltip_text or _extract_tooltip_text_from_page(page)
        if _contains_requisites_text(captured_tooltip_text or ""):
            return captured_tooltip_text

        html = _safe_page_content(page)
        html_tooltip = _extract_tooltip_text_from_html(html)
        if _contains_requisites_text(html_tooltip or ""):
            return html_tooltip

        text = _safe_page_text(page)
        if _contains_requisites_text(text):
            return text

        try:
            page.wait_for_timeout(700)
        except Exception:
            break
    return captured_tooltip_text


def _trigger_supplier_tooltip(page: Page) -> None:
    for selector in SELLER_TOOLTIP_TRIGGER_SELECTORS:
        try:
            locator = page.locator(selector).first
            if locator.count() == 0:
                continue
            locator.scroll_into_view_if_needed(timeout=3_000)
            _hover_like_human(page, locator)
            try:
                locator.click(timeout=3_000, force=True)
            except Exception:
                pass
            _dispatch_tooltip_events(page, locator)
            _force_tooltip_open_via_js(page, selector)
            if _tooltip_visible(page) or _contains_requisites_text(_extract_requisites_text_via_dom(page) or ""):
                return
        except Exception:
            continue


def _force_tooltip_open_via_js(page: Page, selector: str) -> None:
    try:
        page.evaluate(
            """
            (selector) => {
              const node = document.querySelector(selector);
              if (!node) return;
              const events = ['mouseenter', 'mouseover', 'mousemove', 'mousedown', 'mouseup', 'click'];
              for (const eventName of events) {
                node.dispatchEvent(new MouseEvent(eventName, { bubbles: true, cancelable: true, composed: true, view: window }));
              }
            }
            """,
            selector,
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
            page.wait_for_timeout(120)
            page.mouse.move(x, y, steps=8)
            page.wait_for_timeout(250)
    except Exception:
        pass
    try:
        locator.hover(timeout=3_000, force=True)
        page.wait_for_timeout(350)
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
        locator.dispatch_event('click')
    except Exception:
        pass
    try:
        page.wait_for_timeout(250)
    except Exception:
        return


def _tooltip_visible(page: Page) -> bool:
    for selector in TOOLTIP_VISIBLE_SELECTORS:
        try:
            locator = page.locator(selector).first
            if locator.count() and locator.is_visible(timeout=1_000):
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
    screenshot_path: Path,
    html_path: Path,
    text_path: Path,
    used_persistent_profile: bool,
    profile_dir: Path | None,
    manual_wait_seconds: int,
    seller_url: str | None,
    navigated_to_seller_page: bool,
) -> InspectResult:
    page_title = _safe_page_title(page)
    tooltip_text = captured_tooltip_text or _extract_tooltip_text_from_page(page) or _extract_tooltip_text_from_html(html)

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
    seller_display_name = _first_non_empty(
        _extract_seller_display_name_from_html(html),
        _extract_seller_display_name(tooltip_text or ""),
        _extract_seller_display_name(text),
    )

    combined_text = "\n".join(filter(None, [page_title or "", text, tooltip_text or "", inn or "", ogrn or "", ogrnip or "", entity_type or "", seller_display_name or ""]))
    anti_bot_detected = _contains_anti_bot_text(combined_text + "\n" + html[:5000], http_status)

    parse_status, note = _detect_parse_status(
        http_status=http_status,
        anti_bot_detected=anti_bot_detected,
        inn=inn,
        entity_type=entity_type,
        manual_wait_seconds=manual_wait_seconds,
        navigated_to_seller_page=navigated_to_seller_page,
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
        screenshot_path=str(screenshot_path),
        html_path=str(html_path),
        text_path=str(text_path),
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

    for match in SELLER_HEADER_RE.finditer(html):
        raw = TAG_RE.sub(' ', match.group('name'))
        normalized = _normalize_text(raw)
        if normalized and 'Wildberries' not in normalized:
            return normalized[:120]
    return None


def _detect_parse_status(
    http_status: int | None,
    anti_bot_detected: bool,
    inn: str | None,
    entity_type: str | None,
    manual_wait_seconds: int,
    navigated_to_seller_page: bool,
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

    if entity_type:
        return "NEEDS_REVIEW", "Есть признаки реквизитов продавца, но ИНН не найден"

    if navigated_to_seller_page:
        return "SELLER_PAGE_OPENED", "Перешли на страницу продавца, но реквизиты не извлеклись"

    if manual_wait_seconds > 0:
        return "PRODUCT_PAGE_OPENED", "Карточка открылась после ручной сессии, но реквизиты пока не извлечены"

    return "PAGE_OPENED_NO_REQUISITES", "Страница открылась, но реквизиты продавца не найдены"




def _contains_requisites_text(text: str) -> bool:
    return any(marker in text for marker in ('ИНН', 'ОГРН', 'ОГРНИП', 'Номер регистрации', 'КПП'))


def _contains_anti_bot_text(text: str, http_status: int | None) -> bool:
    text_lower = text.lower()
    if http_status == 498:
        return True
    return any(pattern.lower() in text_lower for pattern in ANTI_BOT_PATTERNS)


def _extract_tooltip_text_from_page(page: Page) -> str | None:
    dom_text = _extract_requisites_text_via_dom(page)
    if _contains_requisites_text(dom_text or ""):
        return dom_text

    for selector in TOOLTIP_VISIBLE_SELECTORS:
        try:
            locator = page.locator(selector).first
            if locator.count() == 0:
                continue
            text = locator.inner_text(timeout=2_000)
            normalized = _normalize_text(text)
            if normalized and _contains_requisites_text(normalized):
                return normalized
        except Exception:
            continue
    try:
        text = page.locator('.tooltip__content').last.inner_text(timeout=1_500)
        normalized = _normalize_text(text)
        if normalized and _contains_requisites_text(normalized):
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
        if not _contains_requisites_text(raw_body):
            continue
        text = TAG_RE.sub(' ', raw_body)
        text = text.replace('&nbsp;', ' ')
        normalized = _normalize_text(text)
        if normalized:
            return normalized

    if _contains_requisites_text(html):
        index_candidates = [idx for marker in ('ИНН', 'ОГРН', 'ОГРНИП', 'КПП', 'Номер регистрации') if (idx := html.find(marker)) != -1]
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
        if len(cleaned) >= 2:
            return cleaned[:120]

    match = SELLER_NAME_RE.search(text)
    if not match:
        return None
    value = " ".join(match.group(1).split())
    return value[:120] if value else None


def _best_effort_wait(page: Page) -> None:
    for state in ("domcontentloaded", "load", "networkidle"):
        try:
            page.wait_for_load_state(state, timeout=10_000)
        except Exception:
            continue
    for _ in range(3):
        try:
            page.wait_for_timeout(700)
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
                page.wait_for_timeout(800)
            except Exception:
                break
    if last_error is not None:
        return f"<!-- page.content failed: {last_error} -->"
    return ""


def _safe_page_text(page: Page) -> str:
    for _ in range(6):
        try:
            return page.locator("body").inner_text(timeout=10_000)
        except Exception:
            try:
                page.wait_for_timeout(700)
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
