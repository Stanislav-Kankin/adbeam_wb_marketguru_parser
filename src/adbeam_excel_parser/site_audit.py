from __future__ import annotations

import re
from dataclasses import dataclass
from urllib.parse import urlparse

import httpx
import tldextract
from selectolax.parser import HTMLParser
from tenacity import retry, retry_if_exception_type, stop_after_attempt, wait_fixed

from adbeam_excel_parser.models import SiteAuditResult, SiteFitStatus, SiteSignals

REQUEST_TIMEOUT_SECONDS = 12.0
DEFAULT_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/123.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
}

POSITIVE_CATALOG_TERMS = (
    "каталог",
    "товар",
    "продукц",
    "интернет-магазин",
    "магазин",
    "catalog",
    "shop",
    "product",
)
POSITIVE_CART_TERMS = (
    "корзин",
    "в корзину",
    "моя корзина",
    "cart",
    "basket",
    "shopping-cart",
    "checkout",
)
POSITIVE_BUY_TERMS = (
    "купить",
    "заказать онлайн",
    "buy now",
    "add to cart",
)
POSITIVE_DELIVERY_TERMS = ("доставк", "самовывоз", "delivery", "shipping")
POSITIVE_PAYMENT_TERMS = ("оплат", "visa", "mastercard", "mir", "payment")
POSITIVE_CONSUMER_TERMS = (
    "в наличии",
    "акция",
    "скидк",
    "новинки",
    "заказ онлайн",
    "быстрый заказ",
    "доставка по",
)
NEGATIVE_REQUEST_ONLY_TERMS = (
    "оставить заявку",
    "отправить заявку",
    "получить консультацию",
    "оставьте контакты",
    "связаться с менеджером",
)
NEGATIVE_CALLBACK_TERMS = (
    "заказать звонок",
    "обратный звонок",
    "перезвоните мне",
    "мы вам перезвоним",
    "callback",
)
NEGATIVE_QUOTE_TERMS = (
    "запросить кп",
    "коммерческ",
    "запросить предложение",
    "получить прайс",
    "рассчитать стоимость",
)
NEGATIVE_B2B_TERMS = (
    "оптов",
    "для дилеров",
    "производство",
    "промышлен",
    "корпоративн",
    "b2b",
)
HACKED_TERMS = (
    "casino",
    "viagra",
    "porn",
    "sex",
    "slot",
    "betting",
    "казино",
    "порно",
)
PRICE_PATTERN = re.compile(r"\d[\d\s]{0,12}(?:₽|руб\.?|рублей)", flags=re.IGNORECASE)

CATALOG_HTML_PATTERNS = (
    "/catalog",
    "/catalogue",
    "/shop",
    "/products",
    "/product/",
    "/catalog/",
)
CART_HTML_PATTERNS = (
    "/cart",
    "/basket",
    "cart",
    "basket",
    "shopping-cart",
    "checkout",
    "data-cart",
)
WILDBERRIES_PATTERNS = ("wildberries.ru", "www.wildberries.ru", "wb.ru", "www.wb.ru")
OZON_PATTERNS = ("ozon.ru", "www.ozon.ru")


@dataclass(slots=True)
class FetchedPage:
    requested_url: str
    final_url: str
    status_code: int
    html: str


def audit_website_url(url: str | None, company_name: str | None = None, row_index: int = 0) -> SiteAuditResult:
    normalized_url = normalize_url(url)
    if not normalized_url:
        return SiteAuditResult(
            row_index=row_index,
            company_name=company_name,
            original_url=url,
            normalized_url=None,
            final_url=None,
            domain=None,
            title=None,
            http_status=None,
            status=SiteFitStatus.NO_SITE,
            status_reason="В строке нет корректного сайта.",
            error=None,
        )

    try:
        page = fetch_page(normalized_url)
    except Exception as exc:
        return SiteAuditResult(
            row_index=row_index,
            company_name=company_name,
            original_url=url,
            normalized_url=normalized_url,
            final_url=None,
            domain=extract_domain(normalized_url),
            title=None,
            http_status=None,
            status=SiteFitStatus.BROKEN,
            status_reason="Сайт не удалось загрузить.",
            error=str(exc),
        )

    if page.status_code >= 400:
        return SiteAuditResult(
            row_index=row_index,
            company_name=company_name,
            original_url=url,
            normalized_url=normalized_url,
            final_url=page.final_url,
            domain=extract_domain(page.final_url),
            title=extract_title(page.html),
            http_status=page.status_code,
            status=SiteFitStatus.BROKEN,
            status_reason=f"HTTP {page.status_code}",
            error=None,
        )

    title = extract_title(page.html)
    signals = extract_signals(page.html)
    status, reason = classify_signals(signals)

    return SiteAuditResult(
        row_index=row_index,
        company_name=company_name,
        original_url=url,
        normalized_url=normalized_url,
        final_url=page.final_url,
        domain=extract_domain(page.final_url),
        title=title,
        http_status=page.status_code,
        status=status,
        status_reason=reason,
        signals=signals,
        error=None,
    )


@retry(
    stop=stop_after_attempt(2),
    wait=wait_fixed(1),
    retry=retry_if_exception_type(httpx.HTTPError),
    reraise=True,
)
def fetch_page(url: str) -> FetchedPage:
    with httpx.Client(
        follow_redirects=True,
        timeout=REQUEST_TIMEOUT_SECONDS,
        headers=DEFAULT_HEADERS,
    ) as client:
        response = client.get(url)
        response.raise_for_status()
        return FetchedPage(
            requested_url=url,
            final_url=str(response.url),
            status_code=response.status_code,
            html=response.text,
        )


def extract_title(html: str) -> str | None:
    if not html:
        return None

    parser = HTMLParser(html)
    title_node = parser.css_first("title")
    if title_node is None:
        return None

    title = title_node.text(strip=True)
    return title or None


def extract_signals(html: str) -> SiteSignals:
    if not html:
        return SiteSignals()

    parser = HTMLParser(html)
    text = parser.body.text(separator=" ", strip=True) if parser.body else parser.text(separator=" ", strip=True)
    text_lower = compact_text(text).lower()
    html_lower = html.lower()
    marketplace_links = extract_marketplace_links(parser)

    hacked_terms = find_present_terms(text_lower, HACKED_TERMS)

    return SiteSignals(
        has_catalog=contains_any(text_lower, POSITIVE_CATALOG_TERMS) or contains_any(html_lower, CATALOG_HTML_PATTERNS),
        has_cart=contains_any(text_lower, POSITIVE_CART_TERMS) or contains_any(html_lower, CART_HTML_PATTERNS),
        has_buy_button=contains_any(text_lower, POSITIVE_BUY_TERMS),
        has_price=bool(PRICE_PATTERN.search(text_lower)) or "₽" in html or "руб" in text_lower,
        has_delivery=contains_any(text_lower, POSITIVE_DELIVERY_TERMS),
        has_payment=contains_any(text_lower, POSITIVE_PAYMENT_TERMS),
        has_consumer_language=contains_any(text_lower, POSITIVE_CONSUMER_TERMS),
        request_only=contains_any(text_lower, NEGATIVE_REQUEST_ONLY_TERMS),
        callback_only=contains_any(text_lower, NEGATIVE_CALLBACK_TERMS),
        quote_only=contains_any(text_lower, NEGATIVE_QUOTE_TERMS),
        has_b2b_language=contains_any(text_lower, NEGATIVE_B2B_TERMS),
        has_wildberries_link=any(contains_any(link, WILDBERRIES_PATTERNS) for link in marketplace_links),
        has_ozon_link=any(contains_any(link, OZON_PATTERNS) for link in marketplace_links),
        marketplace_links_found=marketplace_links,
        hacked_terms=hacked_terms,
    )


def classify_signals(signals: SiteSignals) -> tuple[SiteFitStatus, str]:
    if signals.hacked_terms:
        return SiteFitStatus.HACKED, f"Найдены подозрительные слова: {', '.join(signals.hacked_terms)}"

    cart_support_score = sum(
        (
            signals.has_catalog,
            signals.has_buy_button,
            signals.has_price,
            signals.has_delivery,
            signals.has_payment,
            signals.has_consumer_language,
        )
    )
    retail_without_cart_score = sum(
        (
            signals.has_catalog,
            signals.has_buy_button,
            signals.has_price,
            signals.has_consumer_language,
        )
    )
    leadgen_without_cart = not signals.has_cart and (signals.request_only or signals.callback_only or signals.quote_only)
    hard_b2b_without_cart = not signals.has_cart and signals.has_b2b_language and not signals.has_consumer_language

    if signals.has_cart:
        if cart_support_score >= 2 and not leadgen_without_cart:
            return SiteFitStatus.FIT_NOW, "Найдена корзина и есть подтверждающие ecom-сигналы."

        return SiteFitStatus.FIT_LATER, "Корзина найдена, но остальных ecom-сигналов пока мало."

    if leadgen_without_cart:
        return SiteFitStatus.NOT_FIT, "Нет корзины: сайт работает через заявку / звонок / запрос КП."

    if hard_b2b_without_cart:
        return SiteFitStatus.NOT_FIT, "Нет корзины: сайт больше похож на B2B/промышленный корпоративный ресурс."

    if retail_without_cart_score >= 2 and signals.has_catalog:
        return SiteFitStatus.FIT_LATER, "Есть товарная структура, но без корзины сайт не fit now."

    return SiteFitStatus.NOT_FIT, "Нет корзины и недостаточно direct/ecom-сигналов."


def normalize_url(raw_url: str | None) -> str | None:
    if raw_url is None:
        return None

    value = raw_url.strip()
    if not value:
        return None

    if not value.startswith(("http://", "https://")):
        value = f"https://{value}"

    parsed = urlparse(value)
    if not parsed.netloc:
        return None

    return value


def extract_domain(url: str | None) -> str | None:
    if not url:
        return None

    extracted = tldextract.extract(url)
    if not extracted.domain:
        return None

    if extracted.suffix:
        return f"{extracted.domain}.{extracted.suffix}"

    return extracted.domain


def compact_text(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def contains_any(text: str, patterns: tuple[str, ...]) -> bool:
    return any(pattern in text for pattern in patterns)


def find_present_terms(text: str, patterns: tuple[str, ...]) -> list[str]:
    return [pattern for pattern in patterns if pattern in text]


def extract_marketplace_links(parser: HTMLParser) -> list[str]:
    links: list[str] = []
    seen: set[str] = set()

    for node in parser.css("a[href]"):
        href = (node.attributes.get("href") or "").strip()
        if not href:
            continue

        href_lower = href.lower()
        if not contains_any(href_lower, WILDBERRIES_PATTERNS + OZON_PATTERNS):
            continue

        if href_lower in seen:
            continue

        seen.add(href_lower)
        links.append(href)

    return links
