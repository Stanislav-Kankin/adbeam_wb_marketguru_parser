from __future__ import annotations

import json
import re
import time
from pathlib import Path

from playwright.sync_api import BrowserContext, Page, sync_playwright

from .models import InspectResult, ResearchRow

INN_RE = re.compile(r"\b(?:ИНН)\s*[:№]?\s*(\d{10}|\d{12})\b", re.IGNORECASE)
OGRN_RE = re.compile(r"\b(?:ОГРН)\s*[:№]?\s*(\d{13})\b", re.IGNORECASE)
OGRNIP_RE = re.compile(r"\b(?:ОГРНИП)\s*[:№]?\s*(\d{15})\b", re.IGNORECASE)
ENTITY_RE = re.compile(r"\b(ИП|ООО)\b")
SELLER_NAME_RE = re.compile(
    r"(?:Продавец|Seller|Поставщик)\s*[:]?\s*([A-Za-zА-Яа-яЁё0-9 .,&\-\"'()]{2,120})",
    re.IGNORECASE,
)

ANTI_BOT_PATTERNS = [
    "Подозрительная активность",
    "Что-то не так",
    "Проверяем браузер",
    "captcha-support@rwb.ru",
    "Новая попытка через",
]


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

            if manual_wait_seconds > 0:
                time.sleep(manual_wait_seconds)
                _best_effort_wait(page)

            page.screenshot(path=str(screenshot_path), full_page=True)
            html = page.content()
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
                screenshot_path=screenshot_path,
                html_path=html_path,
                text_path=text_path,
                used_persistent_profile=profile_dir is not None,
                profile_dir=profile_dir,
                manual_wait_seconds=manual_wait_seconds,
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


def _build_result(
    row_number: int,
    url: str,
    page: Page,
    http_status: int | None,
    html: str,
    text: str,
    screenshot_path: Path,
    html_path: Path,
    text_path: Path,
    used_persistent_profile: bool,
    profile_dir: Path | None,
    manual_wait_seconds: int,
) -> InspectResult:
    page_title = _safe_page_title(page)
    combined_text = "\n".join(filter(None, [page_title or "", text, html[:50_000]]))
    anti_bot_detected = _contains_anti_bot_text(combined_text, http_status)
    inn = _first_match(INN_RE, combined_text)
    ogrn = _first_match(OGRN_RE, combined_text)
    ogrnip = _first_match(OGRNIP_RE, combined_text)
    entity_type = _first_match(ENTITY_RE, combined_text)
    seller_display_name = _extract_seller_display_name(combined_text)

    parse_status, note = _detect_parse_status(
        http_status=http_status,
        anti_bot_detected=anti_bot_detected,
        inn=inn,
        entity_type=entity_type,
        manual_wait_seconds=manual_wait_seconds,
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


def _detect_parse_status(
    http_status: int | None,
    anti_bot_detected: bool,
    inn: str | None,
    entity_type: str | None,
    manual_wait_seconds: int,
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
        return "SUCCESS", "ИНН найден в тексте страницы"

    if entity_type:
        return "NEEDS_REVIEW", "Есть признаки реквизитов продавца, но ИНН не найден"

    if manual_wait_seconds > 0:
        return "PRODUCT_PAGE_OPENED", "Карточка открылась после ручной сессии, но реквизиты пока не извлечены"

    return "PAGE_OPENED_NO_REQUISITES", "Страница открылась, но реквизиты продавца не найдены"


def _contains_anti_bot_text(text: str, http_status: int | None) -> bool:
    text_lower = text.lower()
    if http_status == 498:
        return True
    return any(pattern.lower() in text_lower for pattern in ANTI_BOT_PATTERNS)


def _extract_seller_display_name(text: str) -> str | None:
    match = SELLER_NAME_RE.search(text)
    if not match:
        return None
    value = " ".join(match.group(1).split())
    return value[:120] if value else None


def _best_effort_wait(page: Page) -> None:
    try:
        page.wait_for_load_state("networkidle", timeout=15_000)
    except Exception:
        pass


def _safe_page_text(page: Page) -> str:
    try:
        return page.locator("body").inner_text(timeout=10_000)
    except Exception:
        return ""


def _safe_page_title(page: Page) -> str | None:
    try:
        return page.title()
    except Exception:
        return None


def _first_match(pattern: re.Pattern[str], text: str) -> str | None:
    match = pattern.search(text)
    return match.group(1) if match else None
