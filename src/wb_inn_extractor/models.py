from __future__ import annotations

from pathlib import Path
from typing import Any

from pydantic import BaseModel, Field


class AnalyzeSummary(BaseModel):
    input_path: Path
    sheet_name: str
    total_rows: int
    total_columns: int
    headers: list[str]
    has_seller_name: bool
    has_brand: bool
    has_wb_url: bool
    has_article: bool
    candidate_key_field: str | None
    preview_rows: list[dict[str, Any]] = Field(default_factory=list)


class ResearchRow(BaseModel):
    source_sheet: str
    source_row_index: int
    product_name: str | None = None
    wb_nm_id: int | None = None
    brand: str | None = None
    seller_name_raw: str | None = None
    wb_candidate_url: str | None = None
    parse_status: str = "PENDING"
    parse_note: str | None = None


class InspectResult(BaseModel):
    row_number: int
    url: str
    page_title: str | None = None
    final_url: str | None = None
    content_text_length: int
    inn: str | None = None
    ogrn: str | None = None
    ogrnip: str | None = None
    entity_type: str | None = None
    note: str | None = None
    screenshot_path: str | None = None
    html_path: str | None = None
