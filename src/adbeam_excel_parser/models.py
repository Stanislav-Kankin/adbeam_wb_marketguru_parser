from __future__ import annotations

from enum import Enum
from typing import Any

from pydantic import BaseModel, Field


class SiteFitStatus(str, Enum):
    FIT_NOW = "FIT_NOW"
    FIT_LATER = "FIT_LATER"
    NOT_FIT = "NOT_FIT"
    BROKEN = "BROKEN"
    HACKED = "HACKED"
    NO_SITE = "NO_SITE"


class RowPreview(BaseModel):
    row_index: int = Field(description="1-based Excel row index")
    website: str | None = None
    values: dict[str, Any] = Field(default_factory=dict)


class ExcelReadSummary(BaseModel):
    file_path: str
    sheet_name: str
    total_rows: int
    headers: list[str]
    website_columns: list[str] = Field(default_factory=list)
    rows_with_websites: int = 0
    preview: list[RowPreview] = Field(default_factory=list)


class WebsiteRow(BaseModel):
    row_index: int
    website: str | None = None
    values: dict[str, Any] = Field(default_factory=dict)


class ExcelWebsiteRows(BaseModel):
    file_path: str
    sheet_name: str
    headers: list[str] = Field(default_factory=list)
    total_rows: int
    website_columns: list[str] = Field(default_factory=list)
    rows: list[WebsiteRow] = Field(default_factory=list)


class SiteSignals(BaseModel):
    has_catalog: bool = False
    has_cart: bool = False
    has_buy_button: bool = False
    has_price: bool = False
    has_delivery: bool = False
    has_payment: bool = False
    has_consumer_language: bool = False
    request_only: bool = False
    callback_only: bool = False
    quote_only: bool = False
    has_b2b_language: bool = False
    has_wildberries_link: bool = False
    has_ozon_link: bool = False
    marketplace_links_found: list[str] = Field(default_factory=list)
    hacked_terms: list[str] = Field(default_factory=list)


class SiteAuditResult(BaseModel):
    row_index: int
    company_name: str | None = None
    original_url: str | None = None
    normalized_url: str | None = None
    final_url: str | None = None
    domain: str | None = None
    title: str | None = None
    http_status: int | None = None
    status: SiteFitStatus
    status_reason: str
    signals: SiteSignals = Field(default_factory=SiteSignals)
    error: str | None = None


class AuditRunSummary(BaseModel):
    file_path: str
    sheet_name: str
    total_rows: int
    audited_rows: int
    status_counts: dict[str, int] = Field(default_factory=dict)
    results: list[SiteAuditResult] = Field(default_factory=list)
    output_file_path: str | None = None
