from __future__ import annotations

import re
from dataclasses import dataclass
from urllib.parse import urljoin, urlparse, urlunparse

import httpx
import tldextract
from selectolax.parser import HTMLParser
from tenacity import retry, retry_if_exception_type, stop_after_attempt, wait_fixed

from adbeam_excel_parser.models import SiteAuditResult, SiteFitStatus, SiteSignals

REQUEST_TIMEOUT_SECONDS = 12.0
PLAYWRIGHT_TIMEOUT_SECONDS = 45.0
PLAYWRIGHT_SETTLE_SECONDS = 8.0
MAX_PAGES_PER_SITE = 6
MAX_CRAWL_CANDIDATES = 16
TLD_EXTRACTOR = tldextract.TLDExtract(cache_dir=None, suffix_list_urls=())
BLOCKED_HTTP_STATUSES = (403, 429, 498, 503)
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
POSITIVE_CHECKOUT_TERMS = (
    "оформить заказ",
    "оформление заказа",
    "перейти к оплате",
    "checkout",
    "place order",
)
POSITIVE_BUY_TERMS = (
    "купить",
    "заказать онлайн",
    "добавить в корзину",
    "buy now",
    "add to cart",
    "add-to-cart",
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
POSITIVE_MANUFACTURER_TERMS = (
    "производитель",
    "собственное производство",
    "наш бренд",
    "официальный сайт",
    "бренд",
    "factory",
    "manufacturer",
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
CHECKOUT_HTML_PATTERNS = (
    "/checkout",
    "/order",
    "/personal/order",
    "checkout",
    "place-order",
)
PRODUCT_HTML_PATTERNS = (
    "/product/",
    "/products/",
    "/catalog/",
)
STRUCTURED_PRODUCT_PATTERNS = (
    "schema.org/product",
    '"@type":"product"',
    '"@type": "product"',
    "itemtype=\"https://schema.org/product\"",
    "itemprop=\"price\"",
)
ECOMMERCE_PLATFORM_PATTERNS = (
    "woocommerce",
    "shopify",
    "insales",
    "opencart",
    "cs-cart",
    "bitrix",
    "bitrix:catalog",
    "tilda-cart",
    "ecwid",
    "retailrocket",
)
COMMON_ECOM_PATHS = (
    "/catalog/",
    "/catalog",
    "/shop/",
    "/shop",
    "/products/",
    "/products",
    "/cart/",
    "/cart",
    "/basket/",
    "/basket",
    "/checkout/",
    "/checkout",
    "/delivery/",
    "/delivery",
    "/payment/",
    "/payment",
)
EXCLUDED_LINK_EXTENSIONS = (
    ".pdf",
    ".doc",
    ".docx",
    ".xls",
    ".xlsx",
    ".zip",
    ".rar",
    ".jpg",
    ".jpeg",
    ".png",
    ".webp",
    ".gif",
    ".svg",
)
WILDBERRIES_PATTERNS = ("wildberries.ru", "www.wildberries.ru", "wb.ru", "www.wb.ru")
OZON_PATTERNS = ("ozon.ru", "www.ozon.ru")
MARKETPLACE_DOMAINS = (
    "wildberries.ru",
    "wb.ru",
    "ozon.ru",
    "market.yandex.ru",
    "megamarket.ru",
    "aliexpress.ru",
)


@dataclass(slots=True)
class FetchedPage:
    requested_url: str
    final_url: str
    status_code: int
    html: str
    content_type: str = ""
    rendered: bool = False


@dataclass(slots=True)
class SiteCrawl:
    entry_page: FetchedPage
    pages: list[FetchedPage]
    errors: list[str]


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
        crawl = crawl_site(normalized_url)
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

    successful_pages = [page for page in crawl.pages if page_is_analyzable(page)]
    entry_page = crawl.entry_page
    domain = extract_domain(entry_page.final_url) or extract_domain(normalized_url)

    if not successful_pages:
        if is_marketplace_domain(domain):
            return SiteAuditResult(
                row_index=row_index,
                company_name=company_name,
                original_url=url,
                normalized_url=normalized_url,
                final_url=entry_page.final_url,
                domain=domain,
                title=extract_title(entry_page.html),
                http_status=entry_page.status_code,
                status=SiteFitStatus.NOT_FIT,
                status_reason="Это маркетплейс/карточка, а не собственный direct-сайт бренда.",
                checked_pages=[page.final_url for page in crawl.pages],
                error="; ".join(crawl.errors) or None,
            )

        return SiteAuditResult(
            row_index=row_index,
            company_name=company_name,
            original_url=url,
            normalized_url=normalized_url,
            final_url=entry_page.final_url,
            domain=domain,
            title=extract_title(entry_page.html),
            http_status=entry_page.status_code,
            status=SiteFitStatus.BROKEN,
            status_reason=f"HTTP {entry_page.status_code}",
            checked_pages=[page.final_url for page in crawl.pages],
            error="; ".join(crawl.errors) or None,
        )

    title = extract_title(successful_pages[0].html)
    signals = extract_signals_from_pages(successful_pages)
    signals.is_marketplace_domain = is_marketplace_domain(domain)
    status, reason = classify_signals(signals, domain=domain)
    if any(page.rendered and page.status_code in BLOCKED_HTTP_STATUSES for page in successful_pages):
        reason = f"{reason} Проверено через браузерный fallback после HTTP {entry_page.status_code}."

    return SiteAuditResult(
        row_index=row_index,
        company_name=company_name,
        original_url=url,
        normalized_url=normalized_url,
        final_url=entry_page.final_url,
        domain=domain,
        title=title,
        http_status=entry_page.status_code,
        status=status,
        status_reason=reason,
        signals=signals,
        checked_pages=[page.final_url for page in successful_pages],
        error=None,
    )


def crawl_site(url: str) -> SiteCrawl:
    errors: list[str] = []

    with httpx.Client(
        follow_redirects=True,
        timeout=REQUEST_TIMEOUT_SECONDS,
        headers=DEFAULT_HEADERS,
    ) as client:
        entry_page = fetch_first_available_page(client, url, errors)
        if entry_page is None:
            raise RuntimeError("; ".join(errors) or "Сайт не удалось загрузить.")

        if should_try_browser_fallback(entry_page):
            try:
                entry_page = fetch_page_with_browser(entry_page.final_url or url)
            except Exception as exc:
                errors.append(f"browser fallback {entry_page.final_url}: {exc}")

        pages = [entry_page]
        if (entry_page.status_code >= 400 and not entry_page.rendered) or not looks_like_html(entry_page):
            return SiteCrawl(entry_page=entry_page, pages=pages, errors=errors)

        domain = extract_domain(entry_page.final_url)
        if is_marketplace_domain(domain):
            return SiteCrawl(entry_page=entry_page, pages=pages, errors=errors)

        seen_urls = {canonical_url(entry_page.final_url)}
        for candidate_url in build_crawl_candidate_urls(entry_page, domain):
            if len(pages) >= MAX_PAGES_PER_SITE:
                break

            canonical_candidate = canonical_url(candidate_url)
            if canonical_candidate in seen_urls:
                continue
            seen_urls.add(canonical_candidate)

            try:
                page = fetch_page_with_client(client, candidate_url)
            except Exception as exc:
                errors.append(f"{candidate_url}: {exc}")
                continue

            if page.status_code < 400 and looks_like_html(page):
                pages.append(page)

        return SiteCrawl(entry_page=entry_page, pages=pages, errors=errors)


def fetch_first_available_page(
    client: httpx.Client,
    url: str,
    errors: list[str],
) -> FetchedPage | None:
    first_page: FetchedPage | None = None

    for candidate_url in build_entry_url_candidates(url):
        try:
            page = fetch_page_with_client(client, candidate_url)
        except Exception as exc:
            errors.append(f"{candidate_url}: {exc}")
            continue

        if first_page is None:
            first_page = page

        if page.status_code < 400:
            return page

    return first_page


def fetch_page(url: str) -> FetchedPage:
    with httpx.Client(
        follow_redirects=True,
        timeout=REQUEST_TIMEOUT_SECONDS,
        headers=DEFAULT_HEADERS,
    ) as client:
        return fetch_page_with_client(client, url)


@retry(
    stop=stop_after_attempt(2),
    wait=wait_fixed(1),
    retry=retry_if_exception_type(httpx.HTTPError),
    reraise=True,
)
def fetch_page_with_client(client: httpx.Client, url: str) -> FetchedPage:
    response = client.get(url)
    return FetchedPage(
        requested_url=url,
        final_url=str(response.url),
        status_code=response.status_code,
        html=response.text,
        content_type=response.headers.get("content-type", ""),
    )


def fetch_page_with_browser(url: str) -> FetchedPage:
    from playwright.sync_api import sync_playwright

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=True)
        try:
            context = browser.new_context(
                locale="ru-RU",
                user_agent=DEFAULT_HEADERS["User-Agent"],
                viewport={"width": 1440, "height": 900},
            )
            try:
                page = context.new_page()
                response = page.goto(
                    url,
                    wait_until="domcontentloaded",
                    timeout=int(PLAYWRIGHT_TIMEOUT_SECONDS * 1000),
                )
                page.wait_for_timeout(int(PLAYWRIGHT_SETTLE_SECONDS * 1000))
                return FetchedPage(
                    requested_url=url,
                    final_url=page.url,
                    status_code=response.status if response is not None else 0,
                    html=page.content(),
                    content_type=response.headers.get("content-type", "text/html") if response else "text/html",
                    rendered=True,
                )
            finally:
                context.close()
        finally:
            browser.close()


def extract_title(html: str) -> str | None:
    if not html:
        return None

    parser = HTMLParser(html)
    title_node = parser.css_first("title")
    if title_node is None:
        return None

    title = title_node.text(strip=True)
    return title or None


def extract_signals_from_pages(pages: list[FetchedPage]) -> SiteSignals:
    aggregate = SiteSignals()
    evidence_pages: list[str] = []

    for page in pages:
        page_signals = extract_signals(page.html, page_url=page.final_url)
        merge_signals(aggregate, page_signals)
        if page_signals.direct_ecom_score > 0:
            evidence_pages.append(page.final_url)

    aggregate.ecommerce_pages_found = unique_values(evidence_pages)
    aggregate.marketplace_links_found = unique_values(aggregate.marketplace_links_found)
    aggregate.hacked_terms = unique_values(aggregate.hacked_terms)
    aggregate.direct_ecom_score = calculate_direct_ecom_score(aggregate)
    return aggregate


def extract_signals(html: str, page_url: str | None = None) -> SiteSignals:
    if not html:
        return SiteSignals()

    parser = HTMLParser(html)
    text = parser.body.text(separator=" ", strip=True) if parser.body else parser.text(separator=" ", strip=True)
    text_lower = compact_text(text).casefold()
    html_lower = html.casefold()
    page_path = urlparse(page_url or "").path.casefold()
    marketplace_links = extract_marketplace_links(parser)

    hacked_terms = find_present_terms(text_lower, HACKED_TERMS)
    has_catalog = (
        contains_any(text_lower, POSITIVE_CATALOG_TERMS)
        or contains_any(html_lower, CATALOG_HTML_PATTERNS)
        or contains_any(page_path, CATALOG_HTML_PATTERNS)
    )
    has_cart = contains_any(text_lower, POSITIVE_CART_TERMS) or contains_any(html_lower, CART_HTML_PATTERNS)
    has_checkout = contains_any(text_lower, POSITIVE_CHECKOUT_TERMS) or contains_any(html_lower, CHECKOUT_HTML_PATTERNS)
    has_buy_button = contains_any(text_lower, POSITIVE_BUY_TERMS) or contains_any(html_lower, POSITIVE_BUY_TERMS)
    has_product_page = (
        contains_any(html_lower, PRODUCT_HTML_PATTERNS)
        or contains_any(page_path, PRODUCT_HTML_PATTERNS)
        or contains_any(html_lower, STRUCTURED_PRODUCT_PATTERNS)
    )
    has_structured_product_data = contains_any(html_lower, STRUCTURED_PRODUCT_PATTERNS)
    has_ecommerce_platform = contains_any(html_lower, ECOMMERCE_PLATFORM_PATTERNS)
    has_price = (
        bool(PRICE_PATTERN.search(text_lower))
        or "₽" in html
        or "руб" in text_lower
        or "data-price" in html_lower
        or "itemprop=\"price\"" in html_lower
    )

    signals = SiteSignals(
        has_catalog=has_catalog,
        has_cart=has_cart,
        has_checkout=has_checkout,
        has_buy_button=has_buy_button,
        has_price=has_price,
        has_product_page=has_product_page,
        has_structured_product_data=has_structured_product_data,
        has_ecommerce_platform=has_ecommerce_platform,
        has_delivery=contains_any(text_lower, POSITIVE_DELIVERY_TERMS),
        has_payment=contains_any(text_lower, POSITIVE_PAYMENT_TERMS),
        has_consumer_language=contains_any(text_lower, POSITIVE_CONSUMER_TERMS),
        has_manufacturer_language=contains_any(text_lower, POSITIVE_MANUFACTURER_TERMS),
        request_only=contains_any(text_lower, NEGATIVE_REQUEST_ONLY_TERMS),
        callback_only=contains_any(text_lower, NEGATIVE_CALLBACK_TERMS),
        quote_only=contains_any(text_lower, NEGATIVE_QUOTE_TERMS),
        has_b2b_language=contains_any(text_lower, NEGATIVE_B2B_TERMS),
        has_wildberries_link=any(contains_any(link, WILDBERRIES_PATTERNS) for link in marketplace_links),
        has_ozon_link=any(contains_any(link, OZON_PATTERNS) for link in marketplace_links),
        marketplace_links_found=marketplace_links,
        hacked_terms=hacked_terms,
    )
    signals.direct_ecom_score = calculate_direct_ecom_score(signals)
    return signals


def classify_signals(signals: SiteSignals, domain: str | None = None) -> tuple[SiteFitStatus, str]:
    signals.direct_ecom_score = calculate_direct_ecom_score(signals)

    if signals.is_marketplace_domain or is_marketplace_domain(domain):
        return SiteFitStatus.NOT_FIT, "Это маркетплейс/карточка, а не собственный direct-сайт бренда."

    if signals.hacked_terms:
        return SiteFitStatus.HACKED, f"Найдены подозрительные слова: {', '.join(signals.hacked_terms)}"

    has_order_flow = signals.has_cart or signals.has_checkout
    has_product_structure = signals.has_catalog or signals.has_product_page or signals.has_structured_product_data
    has_commercial_surface = signals.has_price or signals.has_buy_button
    leadgen_without_cart = not has_order_flow and (signals.request_only or signals.callback_only or signals.quote_only)
    hard_b2b_without_cart = not has_order_flow and signals.has_b2b_language and not signals.has_consumer_language

    if has_order_flow:
        if has_product_structure and has_commercial_surface:
            return SiteFitStatus.FIT_NOW, "Найдены корзина/checkout, товарная структура и коммерческие ecom-сигналы."

        if signals.direct_ecom_score >= 6:
            return SiteFitStatus.FIT_NOW, "Найдена корзина/checkout и достаточно подтверждающих direct-ecom сигналов."

        return SiteFitStatus.FIT_LATER, "Корзина/checkout найдены, но каталог, цены или товарные страницы подтверждены слабо."

    if leadgen_without_cart:
        return SiteFitStatus.NOT_FIT, "Нет корзины: сайт работает через заявку / звонок / запрос КП."

    if hard_b2b_without_cart:
        return SiteFitStatus.NOT_FIT, "Нет корзины: сайт больше похож на B2B/промышленный корпоративный ресурс."

    if has_product_structure and has_commercial_surface and signals.direct_ecom_score >= 5:
        return SiteFitStatus.FIT_LATER, "Есть каталог/товарные страницы и коммерческие признаки, но корзина не подтверждена."

    if has_product_structure and signals.direct_ecom_score >= 6:
        return SiteFitStatus.FIT_LATER, "Есть выраженная товарная структура, но корзина и checkout не подтверждены."

    if signals.has_manufacturer_language and has_product_structure and signals.marketplace_links_found:
        return SiteFitStatus.FIT_LATER, "Похоже на бренд/производителя с товарной структурой, но direct-корзина не подтверждена."

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


def build_entry_url_candidates(url: str) -> list[str]:
    parsed = urlparse(url)
    if not parsed.scheme or not parsed.netloc:
        return [url]

    candidates = [url]
    schemes = ["https", "http"] if parsed.scheme == "https" else ["http", "https"]
    hosts = [parsed.netloc]

    if parsed.netloc.startswith("www."):
        hosts.append(parsed.netloc[4:])
    else:
        hosts.append(f"www.{parsed.netloc}")

    for scheme in schemes:
        for host in hosts:
            candidate = urlunparse((
                scheme,
                host,
                parsed.path or "",
                "",
                parsed.query or "",
                "",
            ))
            if candidate not in candidates:
                candidates.append(candidate)

    return candidates


def build_crawl_candidate_urls(entry_page: FetchedPage, domain: str | None) -> list[str]:
    parser = HTMLParser(entry_page.html)
    scored_candidates: dict[str, int] = {}

    for node in parser.css("a[href]"):
        href = (node.attributes.get("href") or "").strip()
        text = node.text(separator=" ", strip=True)
        candidate_url = normalize_crawl_url(entry_page.final_url, href)
        if not candidate_url:
            continue
        if not is_same_domain(candidate_url, domain):
            continue

        score = score_crawl_link(candidate_url, text)
        if score <= 0:
            continue
        scored_candidates[candidate_url] = max(score, scored_candidates.get(candidate_url, 0))

    parsed = urlparse(entry_page.final_url)
    origin = urlunparse((parsed.scheme, parsed.netloc, "", "", "", ""))
    for index, path in enumerate(COMMON_ECOM_PATHS):
        candidate_url = urljoin(origin, path)
        if candidate_url not in scored_candidates:
            scored_candidates[candidate_url] = 50 - index

    return [
        url
        for url, _score in sorted(
            scored_candidates.items(),
            key=lambda item: (-item[1], len(urlparse(item[0]).path), item[0]),
        )
    ][:MAX_CRAWL_CANDIDATES]


def normalize_crawl_url(base_url: str, href: str) -> str | None:
    if not href:
        return None

    lowered = href.casefold()
    if lowered.startswith(("mailto:", "tel:", "javascript:", "#")):
        return None

    url = urljoin(base_url, href)
    parsed = urlparse(url)
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        return None
    if any(parsed.path.casefold().endswith(extension) for extension in EXCLUDED_LINK_EXTENSIONS):
        return None

    return urlunparse((parsed.scheme, parsed.netloc, parsed.path or "/", "", "", ""))


def score_crawl_link(url: str, text: str) -> int:
    value = f"{url} {text}".casefold()
    score = 0

    weighted_terms = (
        (CATALOG_HTML_PATTERNS + POSITIVE_CATALOG_TERMS, 8),
        (PRODUCT_HTML_PATTERNS, 6),
        (CART_HTML_PATTERNS + POSITIVE_CART_TERMS, 5),
        (CHECKOUT_HTML_PATTERNS + POSITIVE_CHECKOUT_TERMS, 5),
        (POSITIVE_DELIVERY_TERMS, 3),
        (POSITIVE_PAYMENT_TERMS, 3),
    )
    for terms, weight in weighted_terms:
        if contains_any(value, terms):
            score += weight

    if contains_any(value, ("blog", "news", "contacts", "about", "vacancy", "privacy", "policy")):
        score -= 6

    return score


def is_same_domain(url: str, domain: str | None) -> bool:
    if not domain:
        return True
    return extract_domain(url) == domain


def canonical_url(url: str) -> str:
    parsed = urlparse(url)
    path = parsed.path.rstrip("/") or "/"
    return urlunparse((parsed.scheme.casefold(), parsed.netloc.casefold(), path, "", "", ""))


def looks_like_html(page: FetchedPage) -> bool:
    content_type = page.content_type.casefold()
    if content_type and "html" not in content_type and "text/plain" not in content_type:
        return False
    return bool(page.html)


def page_is_analyzable(page: FetchedPage) -> bool:
    return looks_like_html(page) and (page.status_code < 400 or page.rendered)


def should_try_browser_fallback(page: FetchedPage) -> bool:
    if page.rendered:
        return False
    if page.status_code in BLOCKED_HTTP_STATUSES:
        return True
    title = extract_title(page.html) or ""
    if title.strip().casefold() in {"загрузка...", "loading..."}:
        return True
    return False


def merge_signals(target: SiteSignals, source: SiteSignals) -> None:
    bool_fields = (
        "has_catalog",
        "has_cart",
        "has_checkout",
        "has_buy_button",
        "has_price",
        "has_product_page",
        "has_structured_product_data",
        "has_ecommerce_platform",
        "has_delivery",
        "has_payment",
        "has_consumer_language",
        "has_manufacturer_language",
        "request_only",
        "callback_only",
        "quote_only",
        "has_b2b_language",
        "is_marketplace_domain",
        "has_wildberries_link",
        "has_ozon_link",
    )
    for field_name in bool_fields:
        setattr(target, field_name, bool(getattr(target, field_name)) or bool(getattr(source, field_name)))

    target.marketplace_links_found.extend(source.marketplace_links_found)
    target.hacked_terms.extend(source.hacked_terms)


def calculate_direct_ecom_score(signals: SiteSignals) -> int:
    return sum(
        (
            3 if signals.has_cart else 0,
            3 if signals.has_checkout else 0,
            2 if signals.has_catalog else 0,
            2 if signals.has_product_page else 0,
            2 if signals.has_structured_product_data else 0,
            2 if signals.has_buy_button else 0,
            2 if signals.has_price else 0,
            1 if signals.has_delivery else 0,
            1 if signals.has_payment else 0,
            1 if signals.has_consumer_language else 0,
            1 if signals.has_ecommerce_platform else 0,
            1 if signals.has_manufacturer_language else 0,
        )
    )


def extract_domain(url: str | None) -> str | None:
    if not url:
        return None

    extracted = TLD_EXTRACTOR(url)
    if not extracted.domain:
        return None

    if extracted.suffix:
        return f"{extracted.domain}.{extracted.suffix}"

    return extracted.domain


def is_marketplace_domain(domain: str | None) -> bool:
    if not domain:
        return False
    return domain.casefold() in MARKETPLACE_DOMAINS


def compact_text(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def contains_any(text: str, patterns: tuple[str, ...]) -> bool:
    return any(pattern in text for pattern in patterns)


def find_present_terms(text: str, patterns: tuple[str, ...]) -> list[str]:
    result: list[str] = []

    for pattern in patterns:
        expression = rf"(?<![\w]){re.escape(pattern)}(?![\w])"
        if re.search(expression, text, flags=re.IGNORECASE):
            result.append(pattern)

    return result


def unique_values(values: list[str]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()

    for value in values:
        normalized = value.strip()
        if not normalized:
            continue
        key = normalized.casefold()
        if key in seen:
            continue
        seen.add(key)
        result.append(normalized)

    return result


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
