"""Extrae promociones de Interbank desde su API publica de IBM Content CMS.

La pagina https://interbank.pe/promociones-catalogo carga sus tarjetas desde
IBM Content CMS. Consultar esa fuente estructurada es mas rapido y estable que
automatizar Chrome, y evita depender del boton "Ver mas" o del desafio web.

Uso:
    python Interbank_Promociones_CMS.py
    python Interbank_Promociones_CMS.py --categoria restaurantes
    python Interbank_Promociones_CMS.py --solo-vigentes
    python Interbank_Promociones_CMS.py --limit 20

Dependencias:
    pip install pandas openpyxl
"""

from __future__ import annotations

import argparse
import html
import json
import random
import re
import ssl
import time
import unicodedata
from datetime import date, datetime, timezone
from html.parser import HTMLParser
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


CMS_HOST = "https://content-us-2.content-cms.com"
TENANT_ID = "9b3f67ef-5a9f-4acc-8ce8-bcc27fa681c7"
SEARCH_URL = f"{CMS_HOST}/api/{TENANT_ID}/delivery/v1/search"
CATALOG_URL = "https://interbank.pe/promociones-catalogo"
DETAIL_BASE_URL = "https://interbank.pe/promociones"
DEFAULT_OUTPUT = "promociones_interbank.xlsx"
CONTENT_TYPE = "CT_PROMOTION_ITEM"
PAGE_SIZE = 500
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/152.0.0.0 Safari/537.36"
)

COLUMNS = [
    "Banco",
    "Tipo",
    "Nombre",
    "Detalle de la promoción",
    "Restricciones",
    "Términos y condiciones",
    "Ubicación",
    "Link",
    "Categorías",
    "Categorías ID",
    "ID promoción",
    "ID CMS",
    "Descuento (%)",
    "Texto del beneficio",
    "Fecha de inicio",
    "Fecha de fin",
    "Estado de vigencia",
    "Productos / medios de pago",
    "Segmentos",
    "Regiones / zonas",
    "Ubigeo CMS",
    "Comercio",
    "Locales",
    "Enlace de compra",
    "Enlaces relacionados",
    "Imagen",
    "Imagen grande",
    "Nombre interno CMS",
    "Orden CMS",
    "Última edición",
    "Estado CMS",
    "Estado de extracción",
    "Fecha de extracción (UTC)",
]

MONTHS = {
    "enero": 1,
    "febrero": 2,
    "marzo": 3,
    "abril": 4,
    "mayo": 5,
    "junio": 6,
    "julio": 7,
    "agosto": 8,
    "septiembre": 9,
    "setiembre": 9,
    "octubre": 10,
    "noviembre": 11,
    "diciembre": 12,
}


class _HTMLTextExtractor(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.parts: list[str] = []

    def handle_data(self, data: str) -> None:
        self.parts.append(data)

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag.lower() == "li":
            self.parts.append("\n• ")
        elif tag.lower() in {"br", "p", "div", "h1", "h2", "h3", "h4"}:
            self.parts.append("\n")

    def handle_endtag(self, tag: str) -> None:
        if tag.lower() in {"li", "p", "div", "h1", "h2", "h3", "h4"}:
            self.parts.append("\n")


def clean_text(value: Any) -> str:
    """Convierte HTML del CMS en texto legible y conserva listas/saltos."""
    if value in (None, ""):
        return ""
    parser = _HTMLTextExtractor()
    try:
        parser.feed(html.unescape(str(value)))
        text = "".join(parser.parts)
    except Exception:
        text = html.unescape(str(value))
    text = text.replace("\xa0", " ").replace("\r\n", "\n").replace("\r", "\n")
    # Excel/OpenXML rechaza caracteres de control que a veces llegan del CMS.
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)
    lines = [re.sub(r"[ \t]+", " ", line).strip() for line in text.split("\n")]
    return "\n".join(line for line in lines if line).strip()


def value(element: Any, default: Any = "") -> Any:
    if not isinstance(element, dict):
        return default
    if "value" in element:
        return element.get("value", default)
    if "values" in element:
        return element.get("values", default)
    return default


def unique_join(items: Any, separator: str = " | ") -> str:
    if items in (None, ""):
        return ""
    if not isinstance(items, (list, tuple, set)):
        items = [items]
    result: list[str] = []
    for item in items:
        text = clean_text(item)
        if text and text not in result:
            result.append(text)
    return separator.join(result)


def slugify(text: str) -> str:
    normalized = unicodedata.normalize("NFKD", clean_text(text))
    normalized = "".join(ch for ch in normalized if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]+", "-", normalized.lower()).strip("-")


def label_from_slug(slug: str) -> str:
    special = {"plin": "Plin", "lima": "Lima"}
    return special.get(slug, slug.replace("-", " ").strip().capitalize())


def product_label(slug: str) -> str:
    slug = slugify(slug)
    aliases = {
        "tarjeta-de-credito": "Tarjeta de crédito",
        "tarjetas-de-credito": "Tarjeta de crédito",
        "tarjeta-credito": "Tarjeta de crédito",
        "tarjetas-credito": "Tarjeta de crédito",
        "tipo-tarjeta-de-credito": "Tarjeta de crédito",
        "tarjeta-de-debito": "Tarjeta de débito",
        "tarjetas-de-debito": "Tarjeta de débito",
        "tarjeta-debito": "Tarjeta de débito",
        "tarjetas-debito": "Tarjeta de débito",
        "tipo-tarjeta-de-debito": "Tarjeta de débito",
        "cuenta-sueldo": "Cuenta sueldo",
        "tipo-cuenta-sueldo": "Cuenta sueldo",
        "cuotas-sin-intereses": "Cuotas sin intereses",
        "plin": "Plin",
        "tipo-plin": "Plin",
    }
    return aliases.get(slug, label_from_slug(slug))


def request_json(url: str, attempts: int = 4, timeout: int = 180) -> dict[str, Any]:
    headers = {
        "User-Agent": USER_AGENT,
        "Accept": "application/json,text/plain,*/*",
        "Accept-Language": "es-PE,es;q=0.9,en;q=0.7",
    }
    context = ssl.create_default_context()
    last_error: Exception | None = None
    for attempt in range(1, attempts + 1):
        try:
            with urlopen(Request(url, headers=headers), timeout=timeout, context=context) as response:
                return json.loads(response.read().decode("utf-8"))
        except (HTTPError, URLError, TimeoutError, OSError, json.JSONDecodeError) as exc:
            last_error = exc
            if attempt == attempts:
                break
            delay = min(12.0, (2 ** (attempt - 1)) + random.random())
            print(f"  Reintento {attempt}/{attempts} en {delay:.1f}s: {exc}")
            time.sleep(delay)
    raise RuntimeError(f"No se pudo consultar {url}: {last_error}")


def search_query(include_inactive: bool = False) -> str:
    query = f"type:{CONTENT_TYPE}"
    if not include_inactive:
        # 'ready' = publicado en Delivery API; 'rtd 1' = habilitado en la malla web.
        query += ' AND status:ready AND text:"rtd 1"'
    return query


def fetch_documents(include_inactive: bool = False) -> list[dict[str, Any]]:
    """Descarga los documentos completos usando el campo indexado `document`."""
    rows: list[dict[str, Any]] = []
    start = 0
    while True:
        params = {
            "q": search_query(include_inactive),
            "fl": "id,name,tags,lastModified,status,document",
            "rows": PAGE_SIZE,
            "start": start,
            "sort": "lastModified desc",
        }
        payload = request_json(f"{SEARCH_URL}?{urlencode(params)}")
        documents = payload.get("documents") or []
        for item in documents:
            raw = item.get("document")
            try:
                document = json.loads(raw) if isinstance(raw, str) else raw
            except json.JSONDecodeError:
                document = None
            if isinstance(document, dict):
                # Algunos metadatos del indice no siempre se repiten en `document`.
                document.setdefault("status", item.get("status", "ready"))
                rows.append(document)
        start += len(documents)
        total = int(payload.get("numFound") or 0)
        print(f"  Descargados {len(rows)} de {total} documentos")
        if not documents or start >= total:
            break
    return rows


def raw_tag_slugs(document: dict[str, Any], prefix: str) -> list[str]:
    result = []
    for tag in document.get("tags") or []:
        tag = clean_text(tag)
        if tag.startswith(prefix):
            slug = tag[len(prefix):].strip()
            if slug and slug not in result:
                result.append(slug)
    return result


def detail_tag_names(details: dict[str, Any]) -> list[str]:
    result: list[str] = []
    tags = (details.get("tags") or {}).get("values") or []
    for tag in tags:
        name = clean_text(value((tag or {}).get("name")))
        if name and name not in result:
            result.append(name)
    return result


def image_url(element: Any) -> str:
    if not isinstance(element, dict):
        return ""
    path = element.get("url")
    if not path:
        path = ((element.get("renditions") or {}).get("default") or {}).get("url")
    if not path:
        path = ((element.get("asset") or {}).get("resourceUri"))
    if not path:
        return ""
    path = clean_text(path)
    return path if path.startswith(("http://", "https://")) else CMS_HOST + path


def links_from_html(*html_values: Any) -> str:
    links: list[str] = []
    for raw in html_values:
        for link in re.findall(r"href=[\"']([^\"']+)", str(raw or ""), flags=re.I):
            link = html.unescape(link).strip()
            if link and link not in links:
                links.append(link)
    return " | ".join(links)


def discount_percentage(*texts: Any) -> int | float | str:
    joined = " ".join(clean_text(item) for item in texts if item)
    match = re.search(r"(?<!\d)(\d{1,3}(?:[.,]\d+)?)\s*%", joined)
    if not match:
        return ""
    number = float(match.group(1).replace(",", "."))
    return int(number) if number.is_integer() else number


def _valid_date(day: int, month: int, year: int | None) -> date | None:
    if year is None or not 2020 <= year <= 2100:
        return None
    try:
        return date(year, month, day)
    except ValueError:
        return None


def dates_from_visible_text(text: str) -> list[date]:
    """Lee fechas numericas y fechas en espanol publicadas en el detalle."""
    plain = unicodedata.normalize("NFKD", clean_text(text).lower())
    plain = "".join(ch for ch in plain if not unicodedata.combining(ch))
    found: list[tuple[int, int, int, int | None]] = []

    for match in re.finditer(r"(?<!\d)(\d{1,2})[/-](\d{1,2})[/-](20\d{2})(?!\d)", plain):
        found.append((match.start(), int(match.group(1)), int(match.group(2)), int(match.group(3))))

    months_pattern = "|".join(MONTHS)
    textual: list[tuple[int, int, int, int | None]] = []
    for match in re.finditer(
        rf"(?<!\d)(\d{{1,2}})\s+(?:de\s+)?({months_pattern})(?:\s+(?:de|del)\s+(20\d{{2}}))?",
        plain,
    ):
        textual.append(
            (match.start(), int(match.group(1)), MONTHS[match.group(2)], int(match.group(3)) if match.group(3) else None)
        )

    explicit = [(pos, year) for pos, _, _, year in found + textual if year]
    for pos, day, month, year in textual:
        if year is None and explicit:
            year = min(explicit, key=lambda item: abs(item[0] - pos))[1]
        found.append((pos, day, month, year))

    result: list[date] = []
    for _, day, month, year in sorted(found):
        parsed = _valid_date(day, month, year)
        if parsed and parsed not in result:
            result.append(parsed)
    return result


def published_dates(document: dict[str, Any], combined_visible_text: str) -> tuple[date | None, date | None, str]:
    """Prioriza las fechas visibles; el CMS conserva a veces fechas antiguas."""
    visible = dates_from_visible_text(combined_visible_text)
    if len(visible) >= 2:
        return min(visible), max(visible), "Texto publicado"
    if len(visible) == 1:
        lowered = clean_text(combined_visible_text).lower()
        # Una fecha precedida por "hasta" representa normalmente el cierre.
        if "hasta" in lowered or "vigente al" in lowered or "disponible al" in lowered:
            return None, visible[0], "Texto publicado"
        return visible[0], visible[0], "Texto publicado"

    elements = document.get("elements") or {}
    parsed: list[date | None] = []
    for field in ("start_date", "end_date"):
        raw = clean_text(value(elements.get(field)))
        try:
            parsed.append(datetime.strptime(raw, "%d/%m/%Y").date())
        except (ValueError, TypeError):
            parsed.append(None)
    return parsed[0], parsed[1], "Metadatos CMS"


def validity_status(start: date | None, end: date | None, active: bool) -> str:
    if not active:
        return "No publicada"
    today = date.today()
    if start and today < start:
        return "Programada"
    if end and today > end:
        return "Finalizada (aún publicada)"
    if start or end:
        return "Vigente"
    return "Publicada (fecha no identificada)"


def row_from(document: dict[str, Any]) -> dict[str, Any]:
    elements = document.get("elements") or {}
    details = value(elements.get("details"), {})
    if not isinstance(details, dict):
        details = {}

    internal_name = clean_text(document.get("name"))
    name = clean_text(value(details.get("name")) or value(details.get("hero_title")) or internal_name)
    description = clean_text(
        value(details.get("description"))
        or value(details.get("nameDescription"))
        or value(details.get("hero_detail"))
    )
    description_html = value(details.get("description_html"))
    conditions_html = value(details.get("conditions_detail"))
    restrictions_html = value(details.get("legal_detail"))
    locations_html = value(details.get("locales_description"))
    visible_text = "\n".join(
        clean_text(item) for item in (description_html, conditions_html, restrictions_html) if item
    )
    start, end, date_source = published_dates(document, visible_text)

    category_slugs = raw_tag_slugs(document, "categoria-")
    product_slugs = raw_tag_slugs(document, "tipo-")
    location_slugs = raw_tag_slugs(document, "ubicacion-")
    display_tags = detail_tag_names(details)
    category_names = [label_from_slug(slug) for slug in category_slugs]
    # Los nombres de detalle respetan mejor tildes y etiquetas editoriales.
    for tag in display_tags:
        normalized = slugify(tag)
        if normalized in category_slugs and tag not in category_names:
            category_names[category_slugs.index(normalized)] = tag

    active = document.get("status", "ready") == "ready" and clean_text(value(elements.get("rtd"))) == "1"
    discount_text = clean_text(value(details.get("discount_text")))
    benefit = discount_text or description or clean_text(value(details.get("nameDescription")))
    buy_link = clean_text(value(details.get("buy_link")) or value(elements.get("form_link")))
    related = links_from_html(description_html, conditions_html, restrictions_html, locations_html)
    if buy_link and buy_link not in related.split(" | "):
        related = unique_join([buy_link, *related.split(" | ")])

    extraction_time = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
    promotion_id = clean_text(value(elements.get("promotion_id")))
    segments = clean_text(value(elements.get("user_segment")))
    cms_products = value(elements.get("promotions"), [])
    product_labels = [product_label(item.strip()) for item in product_slugs]
    for item in cms_products if isinstance(cms_products, list) else [cms_products]:
        label = product_label(clean_text(item))
        if label and label not in product_labels:
            product_labels.append(label)
    product_priority = {
        "Tarjeta de crédito": 1,
        "Tarjeta de débito": 2,
        "Plin": 3,
        "Cuenta sueldo": 4,
        "Cuotas sin intereses": 5,
    }
    product_labels = sorted(
        dict.fromkeys(product_labels),
        key=lambda item: (product_priority.get(item, 99), item),
    )

    return {
        "Banco": "Interbank",
        "Tipo": unique_join(category_names),
        "Nombre": name,
        "Detalle de la promoción": clean_text(description_html) or description,
        "Restricciones": clean_text(restrictions_html),
        "Términos y condiciones": clean_text(conditions_html),
        "Ubicación": unique_join(label_from_slug(item) for item in location_slugs),
        "Link": f"{DETAIL_BASE_URL}/{internal_name}" if internal_name else CATALOG_URL,
        "Categorías": unique_join(category_names),
        "Categorías ID": unique_join(category_slugs),
        "ID promoción": promotion_id,
        "ID CMS": clean_text(document.get("id")),
        "Descuento (%)": discount_percentage(discount_text, description, description_html),
        "Texto del beneficio": benefit,
        "Fecha de inicio": start.isoformat() if start else "",
        "Fecha de fin": end.isoformat() if end else "",
        "Estado de vigencia": validity_status(start, end, active),
        "Productos / medios de pago": unique_join(product_labels),
        "Segmentos": segments,
        "Regiones / zonas": unique_join(label_from_slug(item) for item in location_slugs),
        "Ubigeo CMS": clean_text(value(elements.get("ubigee"))),
        "Comercio": name,
        "Locales": clean_text(locations_html),
        "Enlace de compra": buy_link,
        "Enlaces relacionados": related,
        "Imagen": image_url(details.get("image_small") or details.get("image_medium") or details.get("image_large")),
        "Imagen grande": image_url(details.get("image_large")),
        "Nombre interno CMS": internal_name,
        "Orden CMS": clean_text(value(elements.get("order"))),
        "Última edición": clean_text(document.get("lastModified")),
        "Estado CMS": f"{'Publicada' if active else 'No publicada'} | fechas: {date_source}",
        "Estado de extracción": "OK",
        "Fecha de extracción (UTC)": extraction_time,
    }


def format_excel(path: Path, row_count: int) -> None:
    from openpyxl import load_workbook

    workbook = load_workbook(path)
    sheet = workbook["Promociones"]
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    sheet.row_dimensions[1].height = 32
    fill = PatternFill("solid", fgColor="00A94F")
    for cell in sheet[1]:
        cell.fill = fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    wide = {
        "Nombre": 32,
        "Detalle de la promoción": 55,
        "Restricciones": 60,
        "Términos y condiciones": 80,
        "Ubicación": 24,
        "Link": 50,
        "Productos / medios de pago": 38,
        "Segmentos": 38,
        "Locales": 70,
        "Enlace de compra": 48,
        "Enlaces relacionados": 55,
        "Imagen": 48,
        "Imagen grande": 48,
    }
    for index, name in enumerate(COLUMNS, 1):
        sheet.column_dimensions[get_column_letter(index)].width = wide.get(name, 19)
    for row in sheet.iter_rows(min_row=2, max_row=row_count + 1):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    sheet.sheet_view.showGridLines = False
    workbook.save(path)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Extrae promociones Interbank desde IBM Content CMS")
    parser.add_argument("--output", default=DEFAULT_OUTPUT, help="Ruta del Excel de salida")
    parser.add_argument("--categoria", default="", help="Slug o nombre, por ejemplo: restaurantes")
    parser.add_argument("--producto", default="", help="Ejemplo: tarjeta-de-credito, tarjeta-de-debito o plin")
    parser.add_argument("--ubicacion", default="", help="Ejemplo: lima o provincias")
    parser.add_argument("--solo-vigentes", action="store_true", help="Conserva solo filas con estado Vigente")
    parser.add_argument(
        "--incluir-inactivos",
        action="store_true",
        help="Incluye el histórico completo del CMS; puede contener miles de registros",
    )
    parser.add_argument("--limit", type=int, default=0, help="Límite de filas; 0 = todas")
    return parser.parse_args()


def row_matches(row: dict[str, Any], category: str, product: str, location: str) -> bool:
    category_slug = slugify(category)
    product_slug = slugify(product)
    location_slug = slugify(location)
    categories = {item.strip() for item in row["Categorías ID"].split(" | ") if item.strip()}
    products = {slugify(item) for item in row["Productos / medios de pago"].split(" | ") if item.strip()}
    locations = {slugify(item) for item in row["Regiones / zonas"].split(" | ") if item.strip()}
    return (
        (not category_slug or category_slug in categories)
        and (not product_slug or product_slug in products)
        and (not location_slug or location_slug in locations)
    )


def main() -> None:
    args = parse_args()
    output = Path(args.output).expanduser().resolve()
    started = time.time()

    print(f"Consultando IBM Content CMS: {SEARCH_URL}")
    documents = fetch_documents(args.incluir_inactivos)
    print(f"Documentos únicos recibidos: {len(documents)}")

    seen: set[str] = set()
    rows: list[dict[str, Any]] = []
    for document in documents:
        key = clean_text(document.get("id"))
        if not key or key in seen:
            continue
        row = row_from(document)
        if row_matches(row, args.categoria, args.producto, args.ubicacion):
            rows.append(row)
        seen.add(key)

    if args.solo_vigentes:
        rows = [row for row in rows if row["Estado de vigencia"] == "Vigente"]
    rows.sort(key=lambda row: (row["Tipo"], row["Nombre"], row["ID promoción"]))
    if args.limit > 0:
        rows = rows[: args.limit]

    frame = pd.DataFrame(rows, columns=COLUMNS)
    output.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        frame.to_excel(writer, sheet_name="Promociones", index=False)
    format_excel(output, len(frame))

    duplicates = int(frame["Link"].duplicated().sum()) if not frame.empty else 0
    print(f"Excel generado: {output}")
    print(f"Filas: {len(frame)} | Enlaces duplicados: {duplicates}")
    print(f"Tiempo total: {time.time() - started:.1f} segundos")


if __name__ == "__main__":
    main()
