from __future__ import annotations

from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook

from .models import AnalyzeSummary, ResearchRow


REQUIRED_HEADERS = {
    "product_name": "Товар",
    "article": "Артикул",
    "brand": "Бренд",
    "seller": "Продавец",
}


def load_active_sheet(input_path: Path):
    workbook = load_workbook(input_path, read_only=True, data_only=True)
    if not workbook.sheetnames:
        raise ValueError("В книге нет листов")
    return workbook[workbook.sheetnames[0]]


def analyze_workbook(input_path: Path) -> AnalyzeSummary:
    sheet = load_active_sheet(input_path)
    headers = [str(value).strip() if value is not None else "" for value in next(sheet.iter_rows(min_row=1, max_row=1, values_only=True))]
    preview_rows: list[dict[str, Any]] = []
    for row in sheet.iter_rows(min_row=2, max_row=min(sheet.max_row, 6), values_only=True):
        preview_rows.append(dict(zip(headers, row, strict=False)))

    has_seller_name = REQUIRED_HEADERS["seller"] in headers
    has_brand = REQUIRED_HEADERS["brand"] in headers
    has_article = REQUIRED_HEADERS["article"] in headers
    has_wb_url = any("wildberries" in header.lower() or "url" in header.lower() or "ссылка" in header.lower() for header in headers)

    candidate_key_field = None
    if has_article:
        candidate_key_field = REQUIRED_HEADERS["article"]
    elif has_seller_name:
        candidate_key_field = REQUIRED_HEADERS["seller"]

    return AnalyzeSummary(
        input_path=input_path,
        sheet_name=sheet.title,
        total_rows=sheet.max_row,
        total_columns=sheet.max_column,
        headers=headers,
        has_seller_name=has_seller_name,
        has_brand=has_brand,
        has_wb_url=has_wb_url,
        has_article=has_article,
        candidate_key_field=candidate_key_field,
        preview_rows=preview_rows,
    )


def build_wb_candidate_url(nm_id: int | None) -> str | None:
    if nm_id is None:
        return None
    return f"https://www.wildberries.ru/catalog/{nm_id}/detail.aspx"


def extract_research_rows(input_path: Path, limit: int) -> list[ResearchRow]:
    sheet = load_active_sheet(input_path)
    headers = [str(value).strip() if value is not None else "" for value in next(sheet.iter_rows(min_row=1, max_row=1, values_only=True))]
    header_index = {header: idx for idx, header in enumerate(headers)}

    result: list[ResearchRow] = []
    for excel_row_index, row in enumerate(sheet.iter_rows(min_row=2, max_row=sheet.max_row, values_only=True), start=2):
        nm_id_raw = row[header_index[REQUIRED_HEADERS["article"]]] if REQUIRED_HEADERS["article"] in header_index else None
        try:
            nm_id = int(nm_id_raw) if nm_id_raw not in (None, "", "—") else None
        except (TypeError, ValueError):
            nm_id = None

        research_row = ResearchRow(
            source_sheet=sheet.title,
            source_row_index=excel_row_index,
            product_name=_get_cell(row, header_index, REQUIRED_HEADERS["product_name"]),
            wb_nm_id=nm_id,
            brand=_get_cell(row, header_index, REQUIRED_HEADERS["brand"]),
            seller_name_raw=_get_cell(row, header_index, REQUIRED_HEADERS["seller"]),
            wb_candidate_url=build_wb_candidate_url(nm_id),
            parse_status="READY_FOR_RESEARCH" if nm_id else "NO_NM_ID",
            parse_note=None if nm_id else "В строке нет корректного Артикул/nmID",
        )
        result.append(research_row)
        if len(result) >= limit:
            break
    return result


def save_research_sample(output_path: Path, rows: list[ResearchRow]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "research_sample"
    headers = [
        "source_sheet",
        "source_row_index",
        "product_name",
        "wb_nm_id",
        "brand",
        "seller_name_raw",
        "wb_candidate_url",
        "parse_status",
        "parse_note",
    ]
    sheet.append(headers)
    for row in rows:
        sheet.append([
            row.source_sheet,
            row.source_row_index,
            row.product_name,
            row.wb_nm_id,
            row.brand,
            row.seller_name_raw,
            row.wb_candidate_url,
            row.parse_status,
            row.parse_note,
        ])
    workbook.save(output_path)


def read_research_row(input_path: Path, row_number: int) -> ResearchRow:
    workbook = load_workbook(input_path, read_only=True, data_only=True)
    sheet = workbook[workbook.sheetnames[0]]
    headers = [str(value).strip() if value is not None else "" for value in next(sheet.iter_rows(min_row=1, max_row=1, values_only=True))]
    target = list(sheet.iter_rows(min_row=row_number, max_row=row_number, values_only=True))
    if not target:
        raise ValueError(f"Строка {row_number} не найдена")
    data = dict(zip(headers, target[0], strict=False))
    return ResearchRow(**data)


def _get_cell(row: tuple[Any, ...], header_index: dict[str, int], header_name: str) -> Any:
    index = header_index.get(header_name)
    if index is None:
        return None
    return row[index]
