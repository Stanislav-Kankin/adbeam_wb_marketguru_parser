from __future__ import annotations

from pathlib import Path
from typing import Any

from pydantic import BaseModel, Field


class SheetAnalyzeSummary(BaseModel):
    sheet_name: str
    status: str
    reason: str | None = None
    detected_header_row: int | None = None
    total_rows_scanned: int = 0
    total_columns: int = 0
    data_rows_count: int = 0
    headers: list[str] = Field(default_factory=list)
    preview_rows: list[dict[str, Any]] = Field(default_factory=list)
    has_seller_name: bool = False
    has_brand: bool = False
    has_wb_url: bool = False
    has_article: bool = False
    candidate_key_field: str | None = None


class AnalyzeSummary(BaseModel):
    input_path: Path
    workbook_sheet_count: int
    selected_sheet_names: list[str] = Field(default_factory=list)
    skipped_sheet_names: list[str] = Field(default_factory=list)
    valid_sheet_names: list[str] = Field(default_factory=list)
    sheet_name: str | None = None
    total_rows: int = 0
    total_columns: int = 0
    headers: list[str] = Field(default_factory=list)
    has_seller_name: bool = False
    has_brand: bool = False
    has_wb_url: bool = False
    has_article: bool = False
    candidate_key_field: str | None = None
    preview_rows: list[dict[str, Any]] = Field(default_factory=list)
    sheets: list[SheetAnalyzeSummary] = Field(default_factory=list)


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
    http_status: int | None = None
    parse_status: str = "UNKNOWN"
    content_text_length: int = 0
    anti_bot_detected: bool = False
    used_persistent_profile: bool = False
    profile_dir: str | None = None
    manual_wait_seconds: int = 0
    seller_url: str | None = None
    navigated_to_seller_page: bool = False
    inn: str | None = None
    ogrn: str | None = None
    ogrnip: str | None = None
    entity_type: str | None = None
    seller_display_name: str | None = None
    note: str | None = None
    screenshot_path: str | None = None
    html_path: str | None = None
    text_path: str | None = None
