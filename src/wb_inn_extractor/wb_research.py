from __future__ import annotations

import json
import re
from pathlib import Path

from playwright.sync_api import Page, sync_playwright

from .models import InspectResult, ResearchRow

INN_RE = re.compile(r"\b(?:ИНН)\s*[:№]?\s*(\d{10}|\d{12})\b", re.IGNORECASE)
OGRN_RE = re.compile(r"\b(?:ОГРН)\s*[:№]?\s*(\d{13})\b", re.IGNORECASE)
OGRNIP_RE = re.compile(r"\b(?:ОГРНИП)\s*[:№]?\s*(\d{15})\b", re.IGNORECASE)
ENTITY_RE = re.compile(r"\b(ИП|ООО)\b")


def inspect_product_row(row_number: int, research_row: ResearchRow, artifacts_dir: Path, headful: bool = False) -> InspectResult:
    if not research_row.wb_candidate_url:
        raise ValueError("У строки нет wb_candidate_url")

    artifacts_dir.mkdir(parents=True, exist_ok=True)
    screenshot_path = artifacts_dir / f"row_{row_number}.png"
    html_path = artifacts_dir / f"row_{row_number}.html"

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=not headful)
        context = browser.new_context(viewport={"width": 1600, "height": 1400})
        page = context.new_page()
        page.goto(research_row.wb_candidate_url, wait_until="domcontentloaded", timeout=60_000)
        _best_effort_wait(page)
        page.screenshot(path=str(screenshot_path), full_page=True)
        html = page.content()
        html_path.write_text(html, encoding="utf-8")
        text = page.locator("body").inner_text(timeout=10_000)
        result = InspectResult(
            row_number=row_number,
            url=research_row.wb_candidate_url,
            page_title=page.title(),
            final_url=page.url,
            content_text_length=len(text),
            inn=_first_match(INN_RE, text),
            ogrn=_first_match(OGRN_RE, text),
            ogrnip=_first_match(OGRNIP_RE, text),
            entity_type=_first_match(ENTITY_RE, text),
            note="Regex-поиск выполнен по тексту страницы без гарантии открытия seller tooltip",
            screenshot_path=str(screenshot_path),
            html_path=str(html_path),
        )
        (artifacts_dir / f"row_{row_number}.json").write_text(
            json.dumps(result.model_dump(mode="json"), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        browser.close()
        return result


def _best_effort_wait(page: Page) -> None:
    try:
        page.wait_for_load_state("networkidle", timeout=15_000)
    except Exception:
        pass


def _first_match(pattern: re.Pattern[str], text: str) -> str | None:
    match = pattern.search(text)
    return match.group(1) if match else None
