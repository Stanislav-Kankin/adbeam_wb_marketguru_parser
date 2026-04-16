from __future__ import annotations

from pathlib import Path
from typing import Callable

from adbeam_excel_parser.excel_reader import extract_website_rows
from adbeam_excel_parser.models import AuditRunSummary, SiteAuditResult
from adbeam_excel_parser.site_audit import audit_website_url

ProgressCallback = Callable[[int, int, str], None]

COMPANY_NAME_HEADERS = (
    "Наименование",
    "Название компании",
    "Компания",
)


def run_excel_audit(file_path: Path, progress_callback: ProgressCallback | None = None) -> AuditRunSummary:
    website_rows = extract_website_rows(file_path)
    results: list[SiteAuditResult] = []
    total_rows_to_audit = len(website_rows.rows)

    for index, row in enumerate(website_rows.rows, start=1):
        website = row.website or ""
        company_name = find_company_name(row.values)

        if progress_callback is not None:
            progress_callback(index, total_rows_to_audit, website)

        results.append(
            audit_website_url(
                url=website,
                company_name=company_name,
                row_index=row.row_index,
            )
        )

    status_counts: dict[str, int] = {}
    for result in results:
        key = result.status.value
        status_counts[key] = status_counts.get(key, 0) + 1

    return AuditRunSummary(
        file_path=str(file_path),
        sheet_name=website_rows.sheet_name,
        total_rows=website_rows.total_rows,
        audited_rows=total_rows_to_audit,
        status_counts=status_counts,
        results=results,
    )


def find_company_name(values: dict[str, object]) -> str | None:
    for header in COMPANY_NAME_HEADERS:
        value = values.get(header)
        if value is None:
            continue

        text = str(value).strip()
        if text:
            return text

    return None



def attach_output_file_path(summary: AuditRunSummary, output_file_path: Path) -> AuditRunSummary:
    summary.output_file_path = str(output_file_path)
    return summary
