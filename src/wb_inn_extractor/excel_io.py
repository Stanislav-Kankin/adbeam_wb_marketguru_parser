from __future__ import annotations

from pathlib import Path
from collections import Counter
from typing import Any, Iterable

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from .models import AnalyzeSummary, ResearchRow, SheetAnalyzeSummary


REQUIRED_HEADERS = {
    "product_name": "Товар",
    "article": "Артикул",
    "brand": "Бренд",
    "seller": "Продавец",
}

HEADER_SCAN_LIMIT = 20
PREVIEW_LIMIT = 5
COMPASS_INN_HEADER_CANDIDATES = [
    "ИНН",
    "ИНН ЮЛ",
    "ИНН ЮЛ/ИП",
    "ИНН компании",
    "ИНН контрагента",
]


def load_active_sheet(input_path: Path):
    workbook = load_workbook(input_path, read_only=True, data_only=True)
    if not workbook.sheetnames:
        raise ValueError("В книге нет листов")
    return workbook[workbook.sheetnames[0]]


def analyze_workbook(input_path: Path, selected_sheets: list[str] | None = None) -> AnalyzeSummary:
    workbook = load_workbook(input_path, read_only=True, data_only=True)
    if not workbook.sheetnames:
        raise ValueError("В книге нет листов")

    normalized_selection = _normalize_sheet_selection(workbook.sheetnames, selected_sheets)
    summaries: list[SheetAnalyzeSummary] = []
    valid_sheet_names: list[str] = []
    skipped_sheet_names: list[str] = []

    for sheet_name in normalized_selection:
        sheet = workbook[sheet_name]
        summary = _analyze_sheet(sheet)
        summaries.append(summary)
        if summary.status == "VALID":
            valid_sheet_names.append(sheet_name)
        else:
            skipped_sheet_names.append(sheet_name)

    first_valid = next((item for item in summaries if item.status == "VALID"), None)
    return AnalyzeSummary(
        input_path=input_path,
        workbook_sheet_count=len(workbook.sheetnames),
        selected_sheet_names=normalized_selection,
        skipped_sheet_names=skipped_sheet_names,
        valid_sheet_names=valid_sheet_names,
        sheet_name=first_valid.sheet_name if first_valid else None,
        total_rows=first_valid.data_rows_count if first_valid else 0,
        total_columns=first_valid.total_columns if first_valid else 0,
        headers=first_valid.headers if first_valid else [],
        has_seller_name=first_valid.has_seller_name if first_valid else False,
        has_brand=first_valid.has_brand if first_valid else False,
        has_wb_url=first_valid.has_wb_url if first_valid else False,
        has_article=first_valid.has_article if first_valid else False,
        candidate_key_field=first_valid.candidate_key_field if first_valid else None,
        preview_rows=first_valid.preview_rows if first_valid else [],
        sheets=summaries,
    )


def build_wb_candidate_url(nm_id: int | None) -> str | None:
    if nm_id is None:
        return None
    return f"https://www.wildberries.ru/catalog/{nm_id}/detail.aspx"


def extract_research_rows(
    input_path: Path,
    limit: int | None = None,
    selected_sheets: list[str] | None = None,
) -> list[ResearchRow]:
    workbook = load_workbook(input_path, read_only=True, data_only=True)
    sheet_names = _normalize_sheet_selection(workbook.sheetnames, selected_sheets)

    result: list[ResearchRow] = []
    seen_pairs: set[tuple[str, str]] = set()

    for sheet_name in sheet_names:
        sheet = workbook[sheet_name]
        sheet_meta = _prepare_sheet_meta(sheet)
        if sheet_meta is None:
            continue

        for excel_row_index, row in enumerate(
            sheet.iter_rows(min_row=sheet_meta["header_row_index"] + 1, values_only=True),
            start=sheet_meta["header_row_index"] + 1,
        ):
            row_values = list(row)
            if _is_effectively_empty_row(row_values):
                continue

            nm_id = _coerce_wb_nm_id(_get_cell(row_values, sheet_meta["header_index"], REQUIRED_HEADERS["article"]))
            brand_raw = _coerce_to_string(_get_cell(row_values, sheet_meta["header_index"], REQUIRED_HEADERS["brand"]))
            seller_raw = _coerce_to_string(_get_cell(row_values, sheet_meta["header_index"], REQUIRED_HEADERS["seller"]))
            product_name = _coerce_to_string(_get_cell(row_values, sheet_meta["header_index"], REQUIRED_HEADERS["product_name"]))

            brand_key = _normalize_key_part(brand_raw)
            seller_key = _normalize_key_part(seller_raw)
            unique_key = (brand_key, seller_key)
            if brand_key or seller_key:
                if unique_key in seen_pairs:
                    continue
                seen_pairs.add(unique_key)

            research_row = ResearchRow(
                source_sheet=sheet.title,
                source_row_index=excel_row_index,
                product_name=product_name,
                wb_nm_id=nm_id,
                brand=brand_raw,
                seller_name_raw=seller_raw,
                wb_candidate_url=build_wb_candidate_url(nm_id),
                parse_status="READY_FOR_RESEARCH" if nm_id else "NO_NM_ID",
                parse_note=None if nm_id else "В строке нет корректного Артикул/nmID",
            )
            result.append(research_row)
            if limit is not None and len(result) >= limit:
                return result
    return result


def summarize_research_rows_by_sheet(rows: list[ResearchRow]) -> dict[str, int]:
    counter = Counter(row.source_sheet for row in rows)
    return dict(sorted(counter.items(), key=lambda item: (item[0].casefold(), item[0])))


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


def read_research_rows_range(input_path: Path, start_row: int = 2, limit: int = 10) -> list[ResearchRow]:
    workbook = load_workbook(input_path, read_only=True, data_only=True)
    sheet = workbook[workbook.sheetnames[0]]
    headers = [str(value).strip() if value is not None else "" for value in next(sheet.iter_rows(min_row=1, max_row=1, values_only=True))]

    result: list[ResearchRow] = []
    max_row = start_row + limit - 1
    for row_number in range(start_row, max_row + 1):
        target = list(sheet.iter_rows(min_row=row_number, max_row=row_number, values_only=True))
        if not target:
            continue
        data = dict(zip(headers, target[0], strict=False))
        result.append(ResearchRow(**data))
    return result


def save_batch_results(output_path: Path, rows: list[dict[str, Any]]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "batch_results"
    headers = [
        "row_number",
        "source_sheet",
        "source_row_index",
        "product_name",
        "wb_nm_id",
        "brand",
        "seller_name_raw",
        "wb_candidate_url",
        "final_url",
        "seller_url",
        "navigated_to_seller_page",
        "seller_display_name",
        "entity_type",
        "inn",
        "ogrn",
        "ogrnip",
        "parse_status",
        "parse_note",
        "http_status",
        "used_persistent_profile",
        "screenshot_path",
        "html_path",
        "text_path",
        "marketguru_source_sheet",
        "marketguru_source_row_index",
        "marketguru_product_name",
        "marketguru_brand",
        "marketguru_seller_name",
        "marketguru_wb_nm_id",
        "marketguru_candidate_url",
        "wb_seller_name",
        "wb_seller_url",
    ]
    sheet.append(headers)
    for row in rows:
        normalized_row = dict(row)
        normalized_row.setdefault("marketguru_source_sheet", row.get("source_sheet"))
        normalized_row.setdefault("marketguru_source_row_index", row.get("source_row_index"))
        normalized_row.setdefault("marketguru_product_name", row.get("product_name"))
        normalized_row.setdefault("marketguru_brand", row.get("brand"))
        normalized_row.setdefault("marketguru_seller_name", row.get("seller_name_raw"))
        normalized_row.setdefault("marketguru_wb_nm_id", row.get("wb_nm_id"))
        normalized_row.setdefault("marketguru_candidate_url", row.get("wb_candidate_url"))
        normalized_row.setdefault("wb_seller_name", row.get("seller_display_name"))
        normalized_row.setdefault("wb_seller_url", row.get("seller_url"))
        sheet.append([normalized_row.get(header) for header in headers])
    workbook.save(output_path)


def merge_batch_results_with_compass(
    batch_results_path: Path,
    compass_path: Path,
    output_path: Path,
) -> dict[str, Any]:
    batch_rows = _read_sheet_as_dicts(batch_results_path)
    compass_payload = _read_compass_rows(compass_path)
    compass_index = compass_payload["index"]

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "final_enriched"

    batch_headers = batch_payload_headers = batch_rows[0].keys() if batch_rows else []
    compass_headers = compass_payload["headers"]
    output_headers = list(batch_headers) + [
        "merge_inn_normalized",
        "compass_match_found",
        "compass_match_count",
        "compass_source_sheet",
        "compass_inn_header",
    ] + [f"compass_{header}" for header in compass_headers]
    sheet.append(output_headers)

    matched_rows = 0
    unmatched_rows = 0
    for row in batch_rows:
        normalized_inn = _normalize_inn_value(row.get("inn"))
        matches = compass_index.get(normalized_inn, []) if normalized_inn else []
        selected_match = matches[0] if matches else None
        if matches:
            matched_rows += 1
        else:
            unmatched_rows += 1

        output_row = dict(row)
        output_row["merge_inn_normalized"] = normalized_inn
        output_row["compass_match_found"] = bool(matches)
        output_row["compass_match_count"] = len(matches)
        output_row["compass_source_sheet"] = compass_payload["sheet_name"] if matches else None
        output_row["compass_inn_header"] = compass_payload["inn_header"] if matches else None
        for header in compass_headers:
            output_row[f"compass_{header}"] = selected_match.get(header) if selected_match else None
        sheet.append([output_row.get(header) for header in output_headers])

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    return {
        "batch_rows": len(batch_rows),
        "matched_rows": matched_rows,
        "unmatched_rows": unmatched_rows,
        "compass_sheet_name": compass_payload["sheet_name"],
        "compass_inn_header": compass_payload["inn_header"],
        "compass_rows_indexed": compass_payload["rows_indexed"],
        "output_path": str(output_path),
    }


def _read_compass_rows(compass_path: Path) -> dict[str, Any]:
    workbook = load_workbook(compass_path, read_only=True, data_only=True)
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        rows_iter = sheet.iter_rows(values_only=True)
        try:
            header_row = next(rows_iter)
        except StopIteration:
            continue
        headers = [_coerce_header(value) for value in header_row]
        if not any(headers):
            continue
        inn_header = _resolve_compass_inn_header(headers)
        if inn_header is None:
            continue

        index: dict[str, list[dict[str, Any]]] = {}
        rows_indexed = 0
        for row in rows_iter:
            row_dict = dict(zip(headers, row, strict=False))
            normalized_inn = _normalize_inn_value(row_dict.get(inn_header))
            if not normalized_inn:
                continue
            index.setdefault(normalized_inn, []).append(row_dict)
            rows_indexed += 1

        return {
            "sheet_name": sheet_name,
            "headers": headers,
            "inn_header": inn_header,
            "rows_indexed": rows_indexed,
            "index": index,
        }
    raise ValueError("В выгрузке Compass не найдена колонка ИНН. Проверь Excel и заголовки столбцов.")


def _read_sheet_as_dicts(input_path: Path) -> list[dict[str, Any]]:
    workbook = load_workbook(input_path, read_only=True, data_only=True)
    sheet = workbook[workbook.sheetnames[0]]
    rows_iter = sheet.iter_rows(values_only=True)
    try:
        header_row = next(rows_iter)
    except StopIteration as exc:
        raise ValueError(f"Пустой Excel: {input_path}") from exc
    headers = [_coerce_header(value) for value in header_row]
    result: list[dict[str, Any]] = []
    for row in rows_iter:
        row_dict = dict(zip(headers, row, strict=False))
        if _is_effectively_empty_row(row_dict.values()):
            continue
        result.append(row_dict)
    return result


def _resolve_compass_inn_header(headers: list[str]) -> str | None:
    normalized_map = {header.casefold(): header for header in headers if header}
    for candidate in COMPASS_INN_HEADER_CANDIDATES:
        found = normalized_map.get(candidate.casefold())
        if found:
            return found
    for header in headers:
        normalized = header.casefold()
        if "инн" in normalized:
            return header
    return None


def _normalize_inn_value(value: Any) -> str | None:
    if value is None:
        return None
    if isinstance(value, float):
        if value.is_integer():
            value = int(value)
        else:
            value = str(value)
    text = str(value).strip()
    if not text:
        return None
    if text.endswith('.0'):
        text = text[:-2]
    digits = ''.join(ch for ch in text if ch.isdigit())
    return digits or None


def _normalize_key_part(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip().casefold()


def _get_cell(row: list[Any] | tuple[Any, ...], header_index: dict[str, int], header_name: str) -> Any:
    index = header_index.get(header_name)
    if index is None or index >= len(row):
        return None
    return row[index]


def _normalize_sheet_selection(all_sheet_names: list[str], selected_sheets: list[str] | None) -> list[str]:
    if not selected_sheets:
        return list(all_sheet_names)

    missing = [sheet_name for sheet_name in selected_sheets if sheet_name not in all_sheet_names]
    if missing:
        raise ValueError(f"Не найдены листы: {', '.join(missing)}")
    return selected_sheets


def _analyze_sheet(sheet: Worksheet) -> SheetAnalyzeSummary:
    sheet_meta = _prepare_sheet_meta(sheet)
    if sheet_meta is None:
        return SheetAnalyzeSummary(
            sheet_name=sheet.title,
            status="SKIPPED",
            reason="На листе не найдены обязательные колонки: Товар / Артикул / Бренд / Продавец",
        )

    preview_rows = _collect_preview_rows(
        sheet,
        header_row_index=sheet_meta["header_row_index"],
        headers=sheet_meta["headers"],
    )

    return SheetAnalyzeSummary(
        sheet_name=sheet.title,
        status="VALID",
        detected_header_row=sheet_meta["header_row_index"],
        total_rows_scanned=sheet_meta["total_rows_scanned"],
        total_columns=len(sheet_meta["headers"]),
        data_rows_count=sheet_meta["data_rows_count"],
        headers=sheet_meta["headers"],
        preview_rows=preview_rows,
        has_seller_name=REQUIRED_HEADERS["seller"] in sheet_meta["header_index"],
        has_brand=REQUIRED_HEADERS["brand"] in sheet_meta["header_index"],
        has_wb_url=any(
            "wildberries" in header.lower() or "url" in header.lower() or "ссылка" in header.lower()
            for header in sheet_meta["headers"]
        ),
        has_article=REQUIRED_HEADERS["article"] in sheet_meta["header_index"],
        candidate_key_field=(
            REQUIRED_HEADERS["article"]
            if REQUIRED_HEADERS["article"] in sheet_meta["header_index"]
            else REQUIRED_HEADERS["seller"]
        ),
    )


def _prepare_sheet_meta(sheet: Worksheet) -> dict[str, Any] | None:
    scanned_rows: list[list[Any]] = []
    header_row_index: int | None = None
    headers: list[str] = []

    for row_number, row in enumerate(sheet.iter_rows(values_only=True), start=1):
        row_values = list(row)
        scanned_rows.append(row_values)
        candidate_headers = [_coerce_header(value) for value in row_values]
        if _is_required_header_row(candidate_headers):
            header_row_index = row_number
            headers = candidate_headers
            break
        if row_number >= HEADER_SCAN_LIMIT:
            break

    if header_row_index is None:
        return None

    header_index = {header: idx for idx, header in enumerate(headers) if header}
    data_rows_count = 0
    for row in sheet.iter_rows(min_row=header_row_index + 1, values_only=True):
        if _is_effectively_empty_row(row):
            continue
        data_rows_count += 1

    return {
        "header_row_index": header_row_index,
        "headers": headers,
        "header_index": header_index,
        "total_rows_scanned": len(scanned_rows),
        "data_rows_count": data_rows_count,
    }


def _collect_preview_rows(sheet: Worksheet, header_row_index: int, headers: list[str]) -> list[dict[str, Any]]:
    preview_rows: list[dict[str, Any]] = []
    for row in sheet.iter_rows(min_row=header_row_index + 1, values_only=True):
        row_values = list(row)
        if _is_effectively_empty_row(row_values):
            continue
        preview_rows.append(dict(zip(headers, row_values, strict=False)))
        if len(preview_rows) >= PREVIEW_LIMIT:
            break
    return preview_rows


def _is_required_header_row(headers: Iterable[str]) -> bool:
    header_set = {header.strip() for header in headers if header and header.strip()}
    return all(required_header in header_set for required_header in REQUIRED_HEADERS.values())


def _coerce_header(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _coerce_to_string(value: Any) -> str | None:
    if value is None:
        return None
    if isinstance(value, str):
        normalized = value.strip()
        return normalized or None
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    normalized = str(value).strip()
    return normalized or None


def _coerce_wb_nm_id(value: Any) -> int | None:
    if value in (None, "", "—"):
        return None
    try:
        if isinstance(value, float):
            if not value.is_integer():
                return None
            return int(value)
        return int(str(value).strip())
    except (TypeError, ValueError):
        return None


def _is_effectively_empty_row(row: Iterable[Any]) -> bool:
    for value in row:
        if value is None:
            continue
        if isinstance(value, str) and not value.strip():
            continue
        return False
    return True



def export_unique_inn_list(batch_results_path: Path, output_path: Path) -> dict[str, Any]:
    batch_rows = _read_sheet_as_dicts(batch_results_path)
    seen: set[str] = set()
    output_rows: list[dict[str, Any]] = []
    for row in batch_rows:
        normalized_inn = _normalize_inn_value(row.get("inn"))
        if not normalized_inn or normalized_inn in seen:
            continue
        seen.add(normalized_inn)
        output_rows.append({
            "inn": normalized_inn,
            "marketguru_seller_name": row.get("marketguru_seller_name") or row.get("seller_name_raw"),
            "marketguru_brand": row.get("marketguru_brand") or row.get("brand"),
            "marketguru_product_name": row.get("marketguru_product_name") or row.get("product_name"),
            "marketguru_source_sheet": row.get("marketguru_source_sheet") or row.get("source_sheet"),
            "wb_seller_name": row.get("wb_seller_name") or row.get("seller_display_name"),
            "wb_seller_url": row.get("wb_seller_url") or row.get("seller_url"),
        })

    output_path.parent.mkdir(parents=True, exist_ok=True)
    _write_dict_rows_to_excel(output_path, output_rows, preferred_headers=[
        "inn",
        "marketguru_seller_name",
        "marketguru_brand",
        "marketguru_product_name",
        "marketguru_source_sheet",
        "wb_seller_name",
        "wb_seller_url",
    ], sheet_name="inn_for_compass")
    return {
        "batch_rows": len(batch_rows),
        "inn_rows": len(output_rows),
        "output_path": str(output_path),
    }



def export_batch_no_inn(batch_results_path: Path, output_path: Path) -> dict[str, Any]:
    batch_rows = _read_sheet_as_dicts(batch_results_path)
    output_rows = [row for row in batch_rows if not _normalize_inn_value(row.get("inn"))]
    output_path.parent.mkdir(parents=True, exist_ok=True)
    _write_dict_rows_to_excel(output_path, output_rows, sheet_name="batch_no_inn")
    return {
        "batch_rows": len(batch_rows),
        "no_inn_rows": len(output_rows),
        "output_path": str(output_path),
    }



def export_compass_unmatched(final_enriched_path: Path, output_path: Path) -> dict[str, Any]:
    enriched_rows = _read_sheet_as_dicts(final_enriched_path)
    output_rows = [row for row in enriched_rows if not _coerce_to_bool(row.get("compass_match_found"))]
    output_path.parent.mkdir(parents=True, exist_ok=True)
    _write_dict_rows_to_excel(output_path, output_rows, sheet_name="compass_unmatched")
    return {
        "final_rows": len(enriched_rows),
        "unmatched_rows": len(output_rows),
        "output_path": str(output_path),
    }



def _write_dict_rows_to_excel(
    output_path: Path,
    rows: list[dict[str, Any]],
    preferred_headers: list[str] | None = None,
    sheet_name: str = "Sheet1",
) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = sheet_name

    headers: list[str] = []
    if preferred_headers:
        headers.extend([header for header in preferred_headers if any(header in row for row in rows)])

    seen = set(headers)
    for row in rows:
        for header in row.keys():
            if header not in seen:
                headers.append(header)
                seen.add(header)

    if not headers:
        headers = preferred_headers or ["note"]
        sheet.append(headers)
        if headers == ["note"]:
            sheet.append(["Нет данных для сохранения"])
    else:
        sheet.append(headers)
        for row in rows:
            sheet.append([row.get(header) for header in headers])

    workbook.save(output_path)



def _coerce_to_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return False
    text = str(value).strip().casefold()
    return text in {"1", "true", "yes", "да"}
