#!/usr/bin/env python3
"""Scrape Valea.bg products and populate the supplied Temu XLSX template.

The scraper uses Valea's public WooCommerce Store API and edits only the
Template worksheet XML inside the original workbook. This preserves the
template's formulas, validations, conditional formatting, hidden sheets and
category-specific rules.
"""

from __future__ import annotations

import argparse
import csv
import html
import json
import math
import os
import re
import shutil
import sys
import tempfile
import time
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from dataclasses import dataclass
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Iterable
from xml.etree import ElementTree as ET


API_BASE = "https://valea.bg/wp-json/wc/store/v1"
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_XML = "http://www.w3.org/XML/1998/namespace"
ET.register_namespace("", NS_MAIN)
ET.register_namespace("r", NS_REL)

ROOT_CATEGORY_SLUGS = {
    "sutieni",
    "bikini",
    "prashki",
    "body",
    "korseti",
    "kolani",
    "top",
    "erotichno-belyo",
    "za-karmachki",
    "bokserki",
}

TEMU_CATEGORY = {
    "bras": "29094",            # Everyday Bras
    "nursing_bras": "29091",    # Nursing Bras
    "briefs": "29081",          # Briefs / high waist
    "bikinis": "29082",         # Bikini panties
    "boy_shorts": "29083",      # Boy Shorts
    "thongs": "29086",          # G-Strings & Thongs
    "bodysuits": "29110",       # Shapewear Bodysuits
    "corsets": "29103",         # Corsets
    "garter_belts": "29120",    # Garter Belts
    "tops": "53993",            # Women Basic Tops
    "lingerie_sets": "29096",   # Lingerie Sets
    "baby_dolls": "29105",      # Baby Dolls
}

BG_COLOR_TO_TEMU = {
    "бял": "White",
    "бяло": "White",
    "черен": "Black",
    "черно": "Black",
    "червен": "Red",
    "червено": "Red",
    "бежов": "Beige",
    "бежово": "Beige",
    "кремав": "Creamy White",
    "кремаво": "Creamy White",
    "сметана": "Creamy White",
    "екрю": "Ivory White",
    "розов": "Pink",
    "розово": "Pink",
    "сив": "Grey",
    "сиво": "Grey",
    "тъмно син": "Navy Blue",
    "тъмносин": "Navy Blue",
    "син": "Blue",
    "синьо": "Blue",
    "кафяв": "Brown",
    "кафяво": "Brown",
    "шоколадов": "Chocolate",
    "шампанско": "Champagne",
    "лилав": "Purple",
    "лилаво": "Purple",
    "зелен": "Green",
    "зелено": "Green",
    "оранжев": "Orange",
    "оранжево": "Orange",
    "многоцветен": "Multicolor",
    "многоцветно": "Multicolor",
}

COMPOSITION = {
    "Polyamide": 80,
    "Elastane": 10,
    "Cotton": 5,
    "Polyester": 5,
}


def clean_text(value: Any, limit: int | None = None) -> str:
    if value is None:
        return ""
    text = html.unescape(str(value))
    text = re.sub(r"<br\s*/?>", "\n", text, flags=re.I)
    text = re.sub(r"</(?:p|li|div|h\d)>", "\n", text, flags=re.I)
    text = re.sub(r"<[^>]+>", " ", text)
    text = unicodedata.normalize("NFKC", text).replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\s*\n\s*", "\n", text)
    text = text.strip(" \n-–•")
    if limit and len(text) > limit:
        text = text[: limit - 1].rstrip() + "…"
    return text


def slugify(value: str) -> str:
    value = unicodedata.normalize("NFKD", value)
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    value = re.sub(r"[^A-Za-z0-9]+", "-", value).strip("-")
    return value.upper() or "ITEM"


def money(raw: Any, minor_unit: int = 2) -> Decimal:
    try:
        return (Decimal(str(raw)) / (Decimal(10) ** minor_unit)).quantize(
            Decimal("0.01"), rounding=ROUND_HALF_UP
        )
    except Exception:
        return Decimal("0.00")


def excel_col_to_num(ref: str) -> int:
    letters = re.match(r"[A-Z]+", ref)
    if not letters:
        raise ValueError(f"Invalid cell reference: {ref}")
    result = 0
    for ch in letters.group(0):
        result = result * 26 + ord(ch) - 64
    return result


def num_to_excel_col(number: int) -> str:
    letters = []
    while number:
        number, remainder = divmod(number - 1, 26)
        letters.append(chr(65 + remainder))
    return "".join(reversed(letters))


class HttpClient:
    def __init__(self, timeout: int = 90, retries: int = 4, delay: float = 0.35):
        self.timeout = timeout
        self.retries = retries
        self.delay = delay
        self.user_agent = "Mozilla/5.0 (compatible; ValeaTemuScraper/1.3)"

    def get_json(self, url: str) -> tuple[Any, dict[str, str]]:
        last_error: Exception | None = None
        for attempt in range(1, self.retries + 1):
            request = urllib.request.Request(
                url,
                headers={"User-Agent": self.user_agent, "Accept": "application/json"},
            )
            try:
                with urllib.request.urlopen(request, timeout=self.timeout) as response:
                    payload = json.loads(response.read().decode("utf-8"))
                    headers = {key.lower(): value for key, value in response.headers.items()}
                if self.delay:
                    time.sleep(self.delay)
                return payload, headers
            except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError, json.JSONDecodeError) as exc:
                last_error = exc
                if attempt == self.retries:
                    break
                wait = min(12.0, 1.5 * (2 ** (attempt - 1)))
                print(f"  retry {attempt}/{self.retries} after {type(exc).__name__}: {url}")
                time.sleep(wait)
        raise RuntimeError(f"Could not fetch {url}: {last_error}")


def fetch_products_page_size(
    client: HttpClient,
    page_size: int,
    search: str = "",
    limit: int = 0,
) -> list[dict[str, Any]]:
    """Fetch all matching products with one fixed WooCommerce page size."""
    products: list[dict[str, Any]] = []
    product_ids: set[int] = set()
    page = 1
    total_pages: int | None = None
    while total_pages is None or page <= total_pages:
        params = {"per_page": str(page_size), "page": str(page)}
        if search:
            params["search"] = search
        url = f"{API_BASE}/products?{urllib.parse.urlencode(params)}"
        print(
            f"Fetching products page {page}{f'/{total_pages}' if total_pages else ''} "
            f"({page_size} per request) …"
        )
        batch, headers = client.get_json(url)
        if not isinstance(batch, list):
            raise RuntimeError("Valea API returned an unexpected products response")
        if total_pages is None:
            header_total = headers.get("x-wp-totalpages")
            total_pages = int(header_total) if header_total else None
        for product in batch:
            if product.get("parent"):
                continue
            category_slugs = {str(cat.get("slug", "")) for cat in product.get("categories", [])}
            if category_slugs and not (category_slugs & ROOT_CATEGORY_SLUGS):
                continue
            product_id = int(product.get("id") or 0)
            if product_id and product_id in product_ids:
                continue
            products.append(product)
            if product_id:
                product_ids.add(product_id)
            if limit and len(products) >= limit:
                return products[:limit]
        if not batch:
            break
        if total_pages is None and len(batch) < page_size:
            break
        page += 1
    return products[:limit] if limit else products


def fetch_products(client: HttpClient, search: str = "", limit: int = 0) -> list[dict[str, Any]]:
    """Fetch products, automatically reducing request size after server errors.

    Valea's Store API sometimes returns HTTP 500 for large collection requests.
    Restarting pagination is required when page size changes because page numbers
    point to different product offsets.
    """
    page_sizes = (25, 10, 5, 1)
    last_error: RuntimeError | None = None
    for attempt, page_size in enumerate(page_sizes, 1):
        try:
            return fetch_products_page_size(client, page_size, search=search, limit=limit)
        except RuntimeError as exc:
            last_error = exc
            if attempt == len(page_sizes):
                break
            next_size = page_sizes[attempt]
            print(
                f"Valea API failed with {page_size} products per request: {exc}\n"
                f"Restarting safely with {next_size} products per request …"
            )
    raise RuntimeError(f"Valea API failed even with one product per request: {last_error}")


def select_temu_category(product: dict[str, Any]) -> str:
    slugs = {str(cat.get("slug", "")) for cat in product.get("categories", [])}
    name = clean_text(product.get("name")).casefold()
    if "za-karmachki" in slugs or "кърмач" in name:
        return TEMU_CATEGORY["nursing_bras"]
    if "bokserki" in slugs or "боксер" in name:
        return TEMU_CATEGORY["boy_shorts"]
    if "prashki" in slugs or "прашк" in name or " string" in f" {name}":
        return TEMU_CATEGORY["thongs"]
    if "body" in slugs or re.search(r"\bбоди\b", name):
        return TEMU_CATEGORY["bodysuits"]
    if "korseti" in slugs or "корсет" in name:
        return TEMU_CATEGORY["corsets"]
    if "kolani" in slugs or "колан" in name:
        return TEMU_CATEGORY["garter_belts"]
    if "top" in slugs or re.search(r"\bтоп\b", name):
        return TEMU_CATEGORY["tops"]
    if "erotichno-belyo" in slugs:
        if any(word in name for word in ("бебидол", "baby doll", "babydoll", "нощница")):
            return TEMU_CATEGORY["baby_dolls"]
        return TEMU_CATEGORY["lingerie_sets"]
    if "bikini" in slugs:
        if any(word in name for word in ("figi big", "висока талия", "high waist")):
            return TEMU_CATEGORY["briefs"]
        return TEMU_CATEGORY["bikinis"]
    return TEMU_CATEGORY["bras"]


def product_attribute_maps(product: dict[str, Any]) -> tuple[dict[str, dict[str, str]], dict[str, list[str]]]:
    term_maps: dict[str, dict[str, str]] = {}
    display_values: dict[str, list[str]] = {}
    for attribute in product.get("attributes", []):
        name = clean_text(attribute.get("name"))
        values: list[str] = []
        mapping: dict[str, str] = {}
        for term in attribute.get("terms", []):
            term_name = clean_text(term.get("name"))
            term_slug = urllib.parse.unquote(str(term.get("slug", ""))).casefold()
            mapping[term_slug] = term_name
            mapping[term_name.casefold()] = term_name
            values.append(term_name)
        term_maps[name.casefold()] = mapping
        display_values[name.casefold()] = values
    return term_maps, display_values


def resolve_attribute(term_maps: dict[str, dict[str, str]], name: str, value: str) -> str:
    mapping = term_maps.get(clean_text(name).casefold(), {})
    decoded = urllib.parse.unquote(str(value)).casefold()
    return mapping.get(decoded) or mapping.get(clean_text(value).casefold()) or clean_text(value)


def temu_color(raw_colors: Iterable[str]) -> str:
    colors = [clean_text(x) for x in raw_colors if clean_text(x)]
    if not colors:
        return "Multicolor"
    mapped: list[str] = []
    for color in colors:
        key = color.casefold()
        mapped_color = BG_COLOR_TO_TEMU.get(key)
        if not mapped_color:
            for bg, en in BG_COLOR_TO_TEMU.items():
                if bg in key:
                    mapped_color = en
                    break
        if mapped_color and mapped_color not in mapped:
            mapped.append(mapped_color)
    if len(mapped) == 1:
        return mapped[0]
    if len(mapped) > 1 or len(colors) > 1:
        return "Multicolor"
    return "Multicolor"


def variation_size(values: list[tuple[str, str]]) -> str:
    by_name = {name.casefold(): clean_text(value) for name, value in values if clean_text(value)}
    band = next((v for k, v in by_name.items() if "подгръдна" in k or "underbust" in k), "")
    cup = next((v for k, v in by_name.items() if "чашка" in k or "cup" in k), "")
    if band and cup:
        return f"{band}{cup.upper()}"
    size_values = [
        value
        for name, value in values
        if not any(word in name.casefold() for word in ("цвят", "color"))
    ]
    return " / ".join(dict.fromkeys(clean_text(x) for x in size_values if clean_text(x))) or "One Size"


def temu_size_selection(size: str) -> tuple[str, str, str]:
    """Return Temu Size Family, Sub-Size Family and standard Size values.

    Valea's bra sizes extend well beyond Temu's predefined cup-band lists, so
    those values use Temu's Custom size family while keeping the exact Valea
    size in both the standard and custom size fields. Common alpha sizes use
    the regular Temu dropdown values.
    """
    original = clean_text(size) or "One Size"
    normalized = re.sub(r"[\s_-]+", "", original).upper()
    if normalized in {"ONESIZE", "ONE", "UNIVERSAL", "УНИВЕРСАЛЕН"}:
        return "2 - Regular Size", "1 - One Size", "one-size"

    alpha_aliases = {
        "XXS": "XXS",
        "XS": "XS",
        "S": "S",
        "M": "M",
        "L": "L",
        "XL": "XL",
        "XXL": "XXL",
        "2XL": "XXL",
        "XXXL": "3XL",
        "3XL": "3XL",
        "XXXXL": "4XL",
        "4XL": "4XL",
        "5XL": "5XL",
        "6XL": "6XL",
        "7XL": "7XL",
        "8XL": "8XL",
        "9XL": "9XL",
        "10XL": "10XL",
    }
    if normalized in alpha_aliases:
        return "2 - Regular Size", "10 - Alpha", alpha_aliases[normalized]

    return "101 - Custom size", "10 - Alpha", original


def alpha_size_rank(size: str) -> int:
    normalized = re.sub(r"[^A-Z0-9]", "", size.upper())
    ranks = {"XXS": 0, "XS": 1, "S": 2, "M": 3, "L": 4, "XL": 5, "XXL": 6, "2XL": 6,
             "XXXL": 7, "3XL": 7, "4XL": 8, "5XL": 9, "6XL": 10}
    return ranks.get(normalized, 3)


def size_measurements(category: str, size: str) -> dict[str, float]:
    """Fill category-required size-chart fields with conservative size-based defaults.

    Valea does not publish product measurements through its API. These values are
    centralized here so they can be replaced with the merchant's official chart.
    """
    result: dict[str, float] = {}
    bra = re.fullmatch(r"(\d{2,3})\s*([A-K])", size.upper().replace(" ", ""))
    rank = alpha_size_rank(size)
    chest = 82 + rank * 6
    waist = 62 + rank * 6
    hip = 88 + rank * 6
    if category in {"29094", "29091", "29103"}:
        band = int(bra.group(1)) if bra else max(65, waist)
        cup_index = (ord(bra.group(2)) - ord("A")) if bra else 2
        result.update({
            "t_5_Size Chart Element:9:Product:10011": float(band),
            "t_5_Size Chart Element:9:Product:10023": float(12 + cup_index),
        })
    if category in {"29081", "29082", "29083", "29086", "29096"}:
        result.update({
            "t_5_Size Chart Element:32:Product:10005": float(waist),
            "t_5_Size Chart Element:32:Product:10008": float(23 + rank),
        })
    if category in {"29110", "29096"}:
        result.update({
            "t_5_Size Chart Element:31:Product:10002": float(chest),
            "t_5_Size Chart Element:31:Product:10005": float(waist),
            "t_5_Size Chart Element:31:Product:10003": float(66 + rank * 2),
        })
    if category in {"53993", "29096"}:
        result.update({
            "t_5_Size Chart Element:5:Product:10002": float(chest),
            "t_5_Size Chart Element:5:Product:10003": float(54 + rank * 2),
        })
    if category in {"29105", "29096"}:
        result.update({
            "t_5_Size Chart Element:7:Product:10002": float(chest),
            "t_5_Size Chart Element:7:Product:10010": float(72 + rank * 2),
        })
    if category == "29096":
        result.update({
            "t_5_Size Chart Element:33:Product:10002": float(chest),
            "t_5_Size Chart Element:33:Product:10003": float(66 + rank * 2),
            "t_5_Size Chart Element:33:Product:10006": float(hip),
            "t_5_Size Chart Element:33:Product:10005": float(waist),
            "t_5_Size Chart Element:34:Product:10005": float(waist),
            "t_5_Size Chart Element:34:Product:10002": float(chest),
            "t_5_Size Chart Element:34:Product:10003": float(36 + rank),
        })
    return result


def fallback_size_measurement(key: str, size: str) -> float:
    """Return a deterministic measurement for a Temu-required size-chart key."""
    rank = alpha_size_rank(size)
    bra = re.fullmatch(r"(\d{2,3})\s*([A-K])", size.upper().replace(" ", ""))
    band = float(int(bra.group(1))) if bra else float(65 + rank * 5)
    cup_height = float(12 + (ord(bra.group(2)) - ord("A"))) if bra else 14.0
    chest = float(82 + rank * 6)
    waist = float(62 + rank * 6)
    hip = float(88 + rank * 6)
    code = key.rsplit(":", 1)[-1]
    if " - " in code:
        code = code.split(" - ", 1)[0]
    values = {
        "10001": chest * 0.42,   # shoulder
        "10002": chest,
        "10003": float(62 + rank * 2),
        "10004": float(58 + rank),
        "10005": waist,
        "10006": hip,
        "10007": hip,
        "10008": float(94 + rank * 2),
        "10009": float(68 + rank * 2),
        "10010": float(70 + rank * 2),
        "10011": band,
        "10013": float(23 + rank * 0.7),
        "10014": float(8.5 + rank * 0.2),
        "10016": float(22 + rank),
        "10017": float(10 + rank * 0.5),
        "10018": float(20 + rank),
        "10023": cup_height,
        "10024": float(40 + rank),
        "10027": float(26 + rank),
        "10028": float(20 + rank),
        "10032": 1.0,
        "10033": float(15 + rank),
        "10036": 1.0,
        "20002": float(5 + rank),
        "20003": float(10 + rank),
        "20004": float(60 + rank * 2),
        "20005": 150.0,
        "20020": 1.0,
        "20030": float(35 + rank),
        "20033": float(23 + rank * 0.7),
        "30000": chest,
        "30001": waist,
        "30002": hip,
        "30003": float(160 + rank * 3),
        "30010": float(25 + rank),
        "30011": float(35 + rank),
    }
    value = float(values.get(code, 1.0))
    if key.endswith(" - min"):
        return max(0.1, round(value - 2, 1))
    if key.endswith(" - max"):
        return round(value + 2, 1)
    return round(value, 1)


@dataclass
class OutputRow:
    product_id: int
    product_name: str
    category: str
    size: str
    color: str
    values: dict[str, Any]


def build_rows(products: list[dict[str, Any]], config: dict[str, Any]) -> list[OutputRow]:
    rows: list[OutputRow] = []
    factor = Decimal(str(config["base_price_factor"]))
    for product_index, product in enumerate(products, 1):
        product_id = int(product.get("id", 0))
        name = clean_text(product.get("name"), 500)
        category = select_temu_category(product)
        term_maps, display_values = product_attribute_maps(product)
        color_values: list[str] = []
        for attr_name, values in display_values.items():
            if "цвят" in attr_name or "color" in attr_name:
                color_values.extend(values)
        base_color = temu_color(color_values)
        description = clean_text(product.get("description") or product.get("short_description"), 2000)
        bullets = [line.strip(" -–•") for line in description.splitlines() if line.strip(" -–•")]
        bullets = list(dict.fromkeys(bullets))[:6]
        images = [str(img.get("src")) for img in product.get("images", []) if img.get("src")][:10]
        prices = product.get("prices") or {}
        minor = int(prices.get("currency_minor_unit", 2) or 2)
        current_price = money(prices.get("price"), minor)
        regular_price = money(prices.get("regular_price"), minor) or current_price
        list_price = max(current_price, regular_price)
        base_price = (current_price * factor).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        if base_price >= list_price:
            base_price = max(Decimal("0.01"), list_price - Decimal("0.01"))
        variations = product.get("variations") or []
        if not variations:
            variations = [{"id": product_id, "attributes": []}]
        for variation_index, variation in enumerate(variations, 1):
            resolved: list[tuple[str, str]] = []
            variation_colors: list[str] = []
            for attribute in variation.get("attributes", []):
                attr_name = clean_text(attribute.get("name"))
                attr_value = resolve_attribute(term_maps, attr_name, str(attribute.get("value", "")))
                resolved.append((attr_name, attr_value))
                if "цвят" in attr_name.casefold() or "color" in attr_name.casefold():
                    variation_colors.append(attr_value)
            size = variation_size(resolved)
            color = temu_color(variation_colors) if variation_colors else base_color
            size_family, sub_size_family, standard_size = temu_size_selection(size)
            variation_id = int(variation.get("id") or product_id)
            parent_code = f"VALEA-{product_id}"
            sku_code = f"{parent_code}-{variation_id}"
            row_values: dict[str, Any] = {
                "t_1_Category": category,
                "t_1_Product Name": name,
                "t_1_Contribution Goods": parent_code,
                "t_1_Contribution SKU": sku_code,
                "t_2_Product Description": description or name,
                "t_3_Property:12": "Polyamide",
                "t_3_Property:26": "Solid color",
                "t_4_Variation Theme": "Color × Size",
                "t_4_Size Family": size_family,
                "t_4_Sub-Size Family": sub_size_family,
                "t_4_Size:3001": standard_size,
                "t_4_Sale Property:1001": color,
                "t_4_Custom Spec:1001": color,
                "t_4_Custom Spec:3001": size,
                "t_5_Unit": "cm-g-ml",
                "t_5_Size Chart Method": "Upload size chart image",
                "t_5_Size Chart Image": config["size_chart_image_url"],
                "t_6_Quantity": int(config["default_quantity"])
                if variation.get("is_in_stock", product.get("is_in_stock", True))
                else 0,
                "t_6_Base Price - EUR": float(base_price),
                "t_6_Reference Link": str(product.get("permalink") or ""),
                "t_6_List Price - EUR": float(list_price),
                "t_6_Weight - g": float(config["package_weight_g"]),
                "t_6_Length - cm": float(config["package_length_cm"]),
                "t_6_Width - cm": float(config["package_width_cm"]),
                "t_6_Height - cm": float(config["package_height_cm"]),
                "t_6_Individually packed": "Yes",
                "t_6_Total packaging quantity": 1,
                "t_6_Packaging unit": "piece",
                "t_7_Shipping Template": config["shipping_template"],
                "t_7_Handling Time": config["handling_time"],
                "t_7_Fulfillment Channel": "I will ship this item myself",
                "t_8_Country/Region of Origin": config["country_of_origin"],
                "t_8_Governance Property:1100100115": parent_code,
                "t_8_Governance Property:3": config["manufacturer"],
            }
            row_values.update(size_measurements(category, size))
            for i, bullet in enumerate(bullets[:6]):
                row_values[f"__bullet_{i}"] = clean_text(bullet, 700)
            for i, image_url in enumerate(images):
                row_values[f"__detail_image_{i}"] = image_url
                row_values[f"__sku_image_{i}"] = image_url
            rows.append(OutputRow(product_id, name, category, size, color, row_values))
        print(f"[{product_index}/{len(products)}] {name}: {len(variations)} size row(s)")
    return rows


class TemuWorkbook:
    def __init__(self, template_path: Path):
        self.template_path = template_path
        self.sheet_member = self._find_sheet_member("Template")
        self.shared_strings = self._read_shared_strings()
        self.header_columns = self._read_header_columns()
        self.special_columns = self._discover_special_columns()
        self.required_columns = self._read_required_columns()

    def _find_sheet_member(self, sheet_name: str) -> str:
        with zipfile.ZipFile(self.template_path) as archive:
            workbook = ET.fromstring(archive.read("xl/workbook.xml"))
            sheet_rid = None
            for sheet in workbook.findall(f"{{{NS_MAIN}}}sheets/{{{NS_MAIN}}}sheet"):
                if sheet.get("name") == sheet_name:
                    sheet_rid = sheet.get(f"{{{NS_REL}}}id")
                    break
            if not sheet_rid:
                raise RuntimeError(f"The workbook has no worksheet named {sheet_name!r}")
            relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
            for rel in relationships.findall(f"{{{NS_PKG_REL}}}Relationship"):
                if rel.get("Id") == sheet_rid:
                    target = str(rel.get("Target", "")).lstrip("/")
                    if target.startswith("xl/"):
                        return target
                    return "xl/" + target
        raise RuntimeError(f"Could not resolve worksheet XML for {sheet_name!r}")

    def _read_shared_strings(self) -> list[str]:
        with zipfile.ZipFile(self.template_path) as archive:
            if "xl/sharedStrings.xml" not in archive.namelist():
                return []
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
        result = []
        for item in root.findall(f"{{{NS_MAIN}}}si"):
            result.append("".join(t.text or "" for t in item.iter(f"{{{NS_MAIN}}}t")))
        return result

    def _cell_text(self, cell: ET.Element) -> str:
        cell_type = cell.get("t")
        value = cell.find(f"{{{NS_MAIN}}}v")
        if cell_type == "s" and value is not None and value.text:
            return self.shared_strings[int(value.text)]
        if cell_type == "inlineStr":
            return "".join(t.text or "" for t in cell.iter(f"{{{NS_MAIN}}}t"))
        return value.text if value is not None and value.text else ""

    def _read_header_columns(self) -> dict[str, list[int]]:
        with zipfile.ZipFile(self.template_path) as archive:
            root = ET.fromstring(archive.read(self.sheet_member))
        result: dict[str, list[int]] = {}
        for row in root.findall(f".//{{{NS_MAIN}}}row"):
            if row.get("r") != "4":
                continue
            for cell in row.findall(f"{{{NS_MAIN}}}c"):
                key = self._cell_text(cell)
                if key:
                    result.setdefault(key, []).append(excel_col_to_num(str(cell.get("r"))))
            break
        return result

    def _read_required_columns(self) -> dict[str, set[str]]:
        member = self._find_sheet_member("GoodsLevelMode")
        with zipfile.ZipFile(self.template_path) as archive:
            root = ET.fromstring(archive.read(member))
        row_values: dict[int, dict[int, str]] = {}
        for row in root.findall(f".//{{{NS_MAIN}}}row"):
            row_number = int(row.get("r", "0"))
            row_values[row_number] = {
                excel_col_to_num(str(cell.get("r"))): self._cell_text(cell)
                for cell in row.findall(f"{{{NS_MAIN}}}c")
            }
        headers = row_values.get(1, {})
        result: dict[str, set[str]] = {}
        for number, values in row_values.items():
            if number == 1:
                continue
            marker = values.get(1, "")
            match = re.fullmatch(r"(\d+)_require", marker)
            if not match:
                continue
            result[match.group(1)] = {
                headers[column]
                for column, value in values.items()
                if column in headers and value == "require"
            }
        return result

    def _discover_special_columns(self) -> dict[str, list[int]]:
        bullets = self.header_columns.get("t_2_Bullet Point", [])[:6]
        detail_images = self.header_columns.get("t_2_Detail Images URL", [])[:10]
        sku_images = self.header_columns.get("t_6_SKU Images URL", [])[:10]
        composition = []
        for key, columns in self.header_columns.items():
            if key.startswith("t_3_Property:15:"):
                composition.extend(columns)
        required = {
            "bullets": bullets,
            "detail_images": detail_images,
            "sku_images": sku_images,
            "composition": composition,
        }
        if not all(required.values()):
            raise RuntimeError(f"Could not identify all required template column groups: {required}")
        return required

    def _column_for_key(self, key: str) -> int | None:
        columns = self.header_columns.get(key, [])
        return columns[0] if columns else None

    def row_to_columns(self, row: OutputRow) -> dict[int, Any]:
        values: dict[int, Any] = {}
        for key, value in row.values.items():
            if key.startswith("__"):
                continue
            column = self._column_for_key(key)
            if column is not None and value not in (None, ""):
                values[column] = value

        # Every enabled Composition field is explicitly supplied. The four
        # user-provided fibres sum to 100%; all other fibres are zero.
        for column in self.special_columns["composition"]:
            values[column] = 0
        composition_key_suffixes = {
            "Polyester": ":35386",
            "Polyamide": ":35387",
            "Elastane": ":35388",
            "Cotton": ":35391",
        }
        for material, percentage in COMPOSITION.items():
            suffix = composition_key_suffixes[material]
            key = next((k for k in self.header_columns if k.startswith("t_3_Property:15:") and k.endswith(suffix)), None)
            if key:
                values[self.header_columns[key][0]] = percentage

        for index, column in enumerate(self.special_columns["bullets"]):
            value = row.values.get(f"__bullet_{index}")
            if value:
                values[column] = value
        for index, column in enumerate(self.special_columns["detail_images"]):
            value = row.values.get(f"__detail_image_{index}")
            if value:
                values[column] = value
        for index, column in enumerate(self.special_columns["sku_images"]):
            value = row.values.get(f"__sku_image_{index}")
            if value:
                values[column] = value

        # Some categories allow several alternative size-chart groups. Temu's
        # matrix marks the active measurements as required. Fill exactly those
        # required measurement columns and leave disabled columns untouched.
        required_keys = self.required_columns.get(row.category, set())
        for key in required_keys:
            if not key.startswith("t_5_Size Chart Element:"):
                continue
            column = self._column_for_key(key)
            if column is not None and column not in values:
                values[column] = fallback_size_measurement(key, row.size)
        return values

    @staticmethod
    def _set_cell(cell: ET.Element, value: Any) -> None:
        for child in list(cell):
            if child.tag in {f"{{{NS_MAIN}}}v", f"{{{NS_MAIN}}}is", f"{{{NS_MAIN}}}f"}:
                cell.remove(child)
        if isinstance(value, bool):
            cell.set("t", "b")
            ET.SubElement(cell, f"{{{NS_MAIN}}}v").text = "1" if value else "0"
        elif isinstance(value, (int, float, Decimal)) and not isinstance(value, bool):
            cell.attrib.pop("t", None)
            if isinstance(value, float) and value.is_integer():
                text = str(int(value))
            else:
                text = format(value, "f") if isinstance(value, Decimal) else str(value)
            ET.SubElement(cell, f"{{{NS_MAIN}}}v").text = text
        else:
            cell.set("t", "inlineStr")
            inline = ET.SubElement(cell, f"{{{NS_MAIN}}}is")
            text_element = ET.SubElement(inline, f"{{{NS_MAIN}}}t")
            string = str(value)
            if string != string.strip():
                text_element.set(f"{{{NS_XML}}}space", "preserve")
            text_element.text = string

    def write(self, output_path: Path, rows: list[OutputRow]) -> None:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(self.template_path, "r") as source:
            sheet_xml = source.read(self.sheet_member)
            root = ET.fromstring(sheet_xml)
            sheet_data = root.find(f"{{{NS_MAIN}}}sheetData")
            if sheet_data is None:
                raise RuntimeError("Template worksheet has no sheetData element")
            row_map = {int(r.get("r", "0")): r for r in sheet_data.findall(f"{{{NS_MAIN}}}row")}
            style_by_column: dict[int, str] = {}
            sample_row = row_map.get(5)
            if sample_row is not None:
                for cell in sample_row.findall(f"{{{NS_MAIN}}}c"):
                    if cell.get("s") is not None:
                        style_by_column[excel_col_to_num(str(cell.get("r")))] = str(cell.get("s"))

            for offset, output_row in enumerate(rows):
                row_number = 5 + offset
                row_element = row_map.get(row_number)
                if row_element is None:
                    row_element = ET.Element(f"{{{NS_MAIN}}}row", {"r": str(row_number)})
                    sheet_data.append(row_element)
                    row_map[row_number] = row_element
                cells = {
                    excel_col_to_num(str(cell.get("r"))): cell
                    for cell in row_element.findall(f"{{{NS_MAIN}}}c")
                }
                for column, value in self.row_to_columns(output_row).items():
                    cell = cells.get(column)
                    if cell is None:
                        ref = f"{num_to_excel_col(column)}{row_number}"
                        cell = ET.Element(f"{{{NS_MAIN}}}c", {"r": ref})
                        if column in style_by_column:
                            cell.set("s", style_by_column[column])
                        row_element.append(cell)
                        cells[column] = cell
                    self._set_cell(cell, value)
                ordered = sorted(row_element.findall(f"{{{NS_MAIN}}}c"), key=lambda c: excel_col_to_num(str(c.get("r"))))
                for cell in list(row_element):
                    if cell.tag == f"{{{NS_MAIN}}}c":
                        row_element.remove(cell)
                row_element.extend(ordered)

            new_sheet_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED, allowZip64=True) as target:
                for item in source.infolist():
                    payload = new_sheet_xml if item.filename == self.sheet_member else source.read(item.filename)
                    target.writestr(item, payload)


def load_config(path: Path) -> dict[str, Any]:
    defaults = {
        "base_price_factor": 0.89,
        "default_quantity": 100,
        "package_weight_g": 150,
        "package_length_cm": 20,
        "package_width_cm": 15,
        "package_height_cm": 5,
        "shipping_template": "Офис",
        "handling_time": "1 Day",
        "country_of_origin": "Bulgaria",
        "manufacturer": "VALEA BG 10 Ltd",
        "size_chart_image_url": "https://valea.bg/wp-content/uploads/2025/04/size-table.jpg",
        "max_rows_per_file": 1900,
    }
    if path.exists():
        with path.open("r", encoding="utf-8") as handle:
            supplied = json.load(handle)
        defaults.update(supplied)
    return defaults


def chunks(rows: list[OutputRow], maximum: int) -> Iterable[list[OutputRow]]:
    for start in range(0, len(rows), maximum):
        yield rows[start : start + maximum]


def write_report(output_dir: Path, products: list[dict[str, Any]], rows: list[OutputRow], config: dict[str, Any]) -> None:
    report = {
        "source": "https://valea.bg/",
        "products": len(products),
        "sku_rows": len(rows),
        "parts": math.ceil(len(rows) / int(config["max_rows_per_file"])),
        "composition": COMPOSITION,
        "category_counts": {},
        "notes": [
            "One workbook row is created for every configured Valea variation.",
            "Valea does not expose official size-chart measurements via its Store API; size-based defaults are centralized in size_measurements().",
            "EU Responsible person is intentionally blank because the configured manufacturer is in Bulgaria (EU).",
        ],
    }
    for row in rows:
        report["category_counts"][row.category] = report["category_counts"].get(row.category, 0) + 1
    with (output_dir / "scrape-report.json").open("w", encoding="utf-8") as handle:
        json.dump(report, handle, ensure_ascii=False, indent=2)
    with (output_dir / "rows-preview.csv").open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["Valea product ID", "Product name", "Temu category", "Size", "Color"])
        for row in rows:
            writer.writerow([row.product_id, row.product_name, row.category, row.size, row.color])


def make_results_zip(output_dir: Path) -> Path:
    zip_path = output_dir / "valea-temu-results.zip"
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED, allowZip64=True) as archive:
        for path in sorted(output_dir.iterdir()):
            if path == zip_path or not path.is_file():
                continue
            archive.write(path, path.name)
    return zip_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Valea.bg → Temu XLSX scraper")
    parser.add_argument("--template", default="template/TEMU_TEMPLATE.xlsx", help="Path to the original Temu XLSX template")
    parser.add_argument("--config", default="config.json", help="JSON configuration path")
    parser.add_argument("--output-dir", default="results", help="Directory for generated files")
    parser.add_argument("--search", default="", help="Optional Valea product search text")
    parser.add_argument("--limit-products", type=int, default=0, help="Stop after N parent products (0 = all)")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    template_path = Path(args.template).resolve()
    config_path = Path(args.config).resolve()
    output_dir = Path(args.output_dir).resolve()
    if not template_path.exists():
        print(f"ERROR: Temu template not found: {template_path}", file=sys.stderr)
        return 2
    config = load_config(config_path)
    client = HttpClient()
    products = fetch_products(client, search=args.search, limit=max(0, args.limit_products))
    if not products:
        print("ERROR: No matching Valea products were found", file=sys.stderr)
        return 3
    rows = build_rows(products, config)
    if not rows:
        print("ERROR: Products were found but no SKU rows were generated", file=sys.stderr)
        return 4
    output_dir.mkdir(parents=True, exist_ok=True)
    for old in output_dir.glob("TEMU_VALEA_UPLOAD_part_*.xlsx"):
        old.unlink()
    workbook = TemuWorkbook(template_path)
    max_rows = int(config["max_rows_per_file"])
    for part_number, part_rows in enumerate(chunks(rows, max_rows), 1):
        output_path = output_dir / f"TEMU_VALEA_UPLOAD_part_{part_number:03d}.xlsx"
        print(f"Writing {output_path.name}: {len(part_rows)} rows …")
        workbook.write(output_path, part_rows)
    write_report(output_dir, products, rows, config)
    zip_path = make_results_zip(output_dir)
    print(f"DONE: {len(products)} products, {len(rows)} SKU rows, {zip_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
