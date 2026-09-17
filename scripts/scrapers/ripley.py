"""Scraper directo de promociones de Banco Ripley Peru.

El sitio carga las promociones desde Firebase Realtime Database. Este programa
consulta la misma fuente JSON (sin Selenium), normaliza los campos y genera un
Excel listo para analizar o consolidar con otros bancos.

Uso:
    python Banco_Ripley_Promociones_Firebase.py
    python Banco_Ripley_Promociones_Firebase.py --solo-vigentes
    python Banco_Ripley_Promociones_Firebase.py --limit 20

Dependencias:
    pip install pandas openpyxl
"""

from __future__ import annotations

import argparse
import html
import json
import random
import re
import socket
import ssl
import time
from datetime import datetime, timezone
from html.parser import HTMLParser
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import quote
from urllib.request import Request, urlopen
from zoneinfo import ZoneInfo

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


FIREBASE_URL = "https://cms-wl-prd.firebaseio.com/beneficios.json"
BASE_DETAIL_URL = "https://www.bancoripley.com.pe/promociones/detalle-promocion.html"
DEFAULT_OUTPUT = "promociones_banco_ripley.xlsx"
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
    "Categoría ID",
    "ID promoción",
    "Descuento (%)",
    "Texto del beneficio",
    "Precio o beneficio adicional",
    "Fecha de inicio",
    "Fecha de fin",
    "Estado de vigencia",
    "Tarjetas válidas",
    "Tipo de tarjeta",
    "Regiones / zonas",
    "Ciudades",
    "Comercio",
    "Enlaces relacionados",
    "Imagen",
    "Logo",
    "Es nueva",
    "Destacada",
    "Activa en Firebase",
    "Última edición",
    "Estado de extracción",
    "Fecha de extracción (UTC)",
]


class _TextExtractor(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.parts: list[str] = []

    def handle_data(self, data: str) -> None:
        self.parts.append(data)

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag.lower() in {"br", "p", "li", "div"}:
            self.parts.append("\n")

    def handle_endtag(self, tag: str) -> None:
        if tag.lower() in {"p", "li", "div"}:
            self.parts.append("\n")


def clean_text(value: Any) -> str:
    """Convierte fragmentos HTML del CMS en texto limpio."""
    if value in (None, ""):
        return ""
    parser = _TextExtractor()
    try:
        parser.feed(html.unescape(str(value)))
        text = "".join(parser.parts)
    except Exception:
        text = html.unescape(str(value))
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = [re.sub(r"[ \t]+", " ", line).strip(" °•\t") for line in text.split("\n")]
    return "\n".join(line for line in lines if line).strip()


def field_value(record: dict[str, Any], key: str, default: Any = "") -> Any:
    value = record.get(key, default)
    if isinstance(value, dict) and "value" in value:
        return value.get("value", default)
    return value


def fetch_json(url: str, attempts: int = 4, timeout: int = 120) -> dict[str, Any]:
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
        except (HTTPError, URLError, TimeoutError, socket.timeout, json.JSONDecodeError) as exc:
            last_error = exc
            if attempt < attempts:
                time.sleep((1.5 ** attempt) + random.uniform(0.1, 0.7))
    raise RuntimeError(f"No se pudo descargar la fuente Firebase: {last_error}")


def parse_cms_datetime(value: Any) -> datetime | None:
    text = clean_text(value)
    if not text:
        return None
    for pattern in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(text, pattern).replace(tzinfo=ZoneInfo("America/Lima"))
        except ValueError:
            pass
    return None


def programmed_dates(record: dict[str, Any]) -> tuple[datetime | None, datetime | None]:
    program = (record.get("config") or {}).get("programar") or {}
    start = field_value(program, "fechaInicioProgramacion")
    end = field_value(program, "fechaFinProgramacion")
    return parse_cms_datetime(start), parse_cms_datetime(end)


def validity_status(start: datetime | None, end: datetime | None) -> str:
    now = datetime.now(ZoneInfo("America/Lima"))
    if start and now < start:
        return "Programada"
    if end and now > end:
        return "Finalizada"
    if start or end:
        return "Vigente"
    return "Sin fechas estructuradas"


def category_catalog(database: dict[str, Any]) -> list[tuple[str, str]]:
    """Devuelve solo las categorias que la web tiene activadas y en su orden."""
    config = database.get("configBeneficios") or {}
    active = config.get("active") or []
    categories: list[tuple[str, str]] = []
    for item in active:
        if not isinstance(item, dict) or item.get("active") is False:
            continue
        category_id = clean_text(item.get("idPromo") or item.get("page"))
        label = clean_text(item.get("title") or category_id)
        if category_id and category_id in database:
            categories.append((category_id, label))
    return categories


def list_field(record: dict[str, Any], key: str) -> list[Any]:
    value = field_value(record, key, [])
    return value if isinstance(value, list) else []


def zones_from(record: dict[str, Any]) -> list[str]:
    output: list[str] = []
    for item in list_field(record, "idZonaCard1"):
        if isinstance(item, dict):
            text = clean_text(item.get("idzona") or item.get("nombre"))
        else:
            text = clean_text(item)
        if text and text not in output:
            output.append(text)
    return output


def locations_from(record: dict[str, Any]) -> tuple[str, list[str]]:
    lines: list[str] = []
    cities: list[str] = []
    for item in list_field(record, "locales"):
        if not isinstance(item, dict):
            continue
        raw_address = clean_text(item.get("direccion"))
        parts = [p.strip() for p in raw_address.split("*") if p.strip()]
        address = parts[0] if parts else raw_address
        city = parts[-1] if len(parts) > 1 else ""
        if city and city.lower() not in {x.lower() for x in cities}:
            cities.append(city)
        coordinates = ", ".join(
            x for x in (clean_text(item.get("cordX")), clean_text(item.get("cordY"))) if x
        )
        line = address
        if city and city.lower() not in address.lower():
            line += f" — {city}"
        if coordinates:
            line += f" ({coordinates})"
        if line and line not in lines:
            lines.append(line)
    return " | ".join(lines), cities


def related_links(record: dict[str, Any]) -> str:
    output: list[str] = []
    for item in list_field(record, "urlRestricciones"):
        if not isinstance(item, dict):
            continue
        label = clean_text(item.get("textoUrl") or item.get("descripcionUrl"))
        url = clean_text(item.get("url"))
        text = f"{label}: {url}" if label and url else (url or label)
        if text and text not in output:
            output.append(text)
    return "\n".join(output)


def percentage_from(record: dict[str, Any]) -> Any:
    text = clean_text(field_value(record, "dctoCard1"))
    match = re.search(r"(\d+(?:[.,]\d+)?)\s*%", text)
    if not match:
        return ""
    number = match.group(1).replace(",", ".")
    return float(number) if "." in number else int(number)


def detail_url(category_id: str, promotion_id: str) -> str:
    return f"{BASE_DETAIL_URL}?{quote(category_id)}={quote(promotion_id)}"


def row_from(record: dict[str, Any], category_id: str, category_label: str) -> dict[str, Any]:
    config = record.get("config") or {}
    promotion_id = clean_text(config.get("id"))
    start, end = programmed_dates(record)
    location, cities = locations_from(record)
    zones = zones_from(record)
    name = clean_text(field_value(record, "txtHeader1") or field_value(record, "nombreEmpresa"))
    discount_text = clean_text(field_value(record, "dctoCard1"))
    detail = clean_text(field_value(record, "detalleDctoCard1"))
    restrictions = clean_text(field_value(record, "restriccionCard1"))
    legal = clean_text(field_value(record, "legalBeneficio"))

    return {
        "Banco": "Banco Ripley",
        "Tipo": category_label,
        "Nombre": name,
        "Detalle de la promoción": " — ".join(x for x in (discount_text, detail) if x),
        "Restricciones": restrictions,
        "Términos y condiciones": legal,
        "Ubicación": location or " | ".join(zones),
        "Link": detail_url(category_id, promotion_id),
        "Categoría ID": category_id,
        "ID promoción": promotion_id,
        "Descuento (%)": percentage_from(record),
        "Texto del beneficio": discount_text,
        "Precio o beneficio adicional": clean_text(field_value(record, "otrosDctoCard1")),
        "Fecha de inicio": start.strftime("%Y-%m-%d") if start else "",
        "Fecha de fin": end.strftime("%Y-%m-%d") if end else "",
        "Estado de vigencia": validity_status(start, end),
        "Tarjetas válidas": clean_text(field_value(record, "txtTarjeta")),
        "Tipo de tarjeta": clean_text(field_value(record, "typeCardApply")),
        "Regiones / zonas": " | ".join(zones),
        "Ciudades": " | ".join(cities),
        "Comercio": clean_text(field_value(record, "nombreEmpresa") or name),
        "Enlaces relacionados": related_links(record),
        "Imagen": clean_text(field_value(record, "imgCard1")),
        "Logo": clean_text(field_value(record, "imgLogo1")),
        "Es nueva": "Sí" if record.get("isNew") or config.get("isNew") else "No",
        "Destacada": "Sí" if field_value(record, "esDestacadoCard1", False) else "No",
        "Activa en Firebase": "Sí" if config.get("active", True) else "No",
        "Última edición": clean_text(config.get("lastEdit")),
        "Estado de extracción": "OK",
        "Fecha de extracción (UTC)": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S"),
    }


def extract_rows(database: dict[str, Any]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    seen: set[tuple[str, str]] = set()
    for category_id, category_label in category_catalog(database):
        category = database.get(category_id) or {}
        records = category.get("active") or []
        for record in records:
            if not isinstance(record, dict):
                continue
            promotion_id = clean_text((record.get("config") or {}).get("id"))
            key = (category_id, promotion_id)
            if not promotion_id or key in seen:
                continue
            rows.append(row_from(record, category_id, category_label))
            seen.add(key)
    return rows


def format_excel(path: Path, row_count: int) -> None:
    from openpyxl import load_workbook

    workbook = load_workbook(path)
    sheet = workbook["Promociones"]
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    sheet.row_dimensions[1].height = 32
    fill = PatternFill("solid", fgColor="7A1E78")
    for cell in sheet[1]:
        cell.fill = fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    wide = {
        "Nombre": 34, "Detalle de la promoción": 48, "Restricciones": 60,
        "Términos y condiciones": 80, "Ubicación": 55, "Link": 52,
        "Tarjetas válidas": 45, "Regiones / zonas": 30, "Ciudades": 28,
        "Enlaces relacionados": 60, "Imagen": 48, "Logo": 45,
    }
    for index, name in enumerate(COLUMNS, 1):
        sheet.column_dimensions[get_column_letter(index)].width = wide.get(name, 19)
    for row in sheet.iter_rows(min_row=2, max_row=row_count + 1):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    sheet.sheet_view.showGridLines = False
    workbook.save(path)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Extrae promociones Banco Ripley desde Firebase")
    parser.add_argument("--output", default=DEFAULT_OUTPUT, help="Ruta del Excel de salida")
    parser.add_argument(
        "--solo-vigentes",
        action="store_true",
        help="Excluye registros finalizados o programados; por defecto replica la rama activa de la web",
    )
    parser.add_argument("--limit", type=int, default=0, help="Limite de filas; 0 = todas")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    output = Path(args.output).expanduser().resolve()
    started = time.time()

    print(f"Descargando promociones desde Firebase: {FIREBASE_URL}")
    database = fetch_json(FIREBASE_URL)
    rows = extract_rows(database)
    print(f"Promociones únicas en categorías activas: {len(rows)}")

    if args.solo_vigentes:
        rows = [row for row in rows if row["Estado de vigencia"] == "Vigente"]
        print(f"Promociones vigentes después del filtro: {len(rows)}")
    if args.limit > 0:
        rows = rows[: args.limit]

    rows.sort(key=lambda row: (row["Tipo"], row["Nombre"], row["ID promoción"]))
    frame = pd.DataFrame(rows, columns=COLUMNS)
    output.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        frame.to_excel(writer, sheet_name="Promociones", index=False)
    format_excel(output, len(frame))

    print(f"Excel generado: {output}")
    print(f"Filas: {len(frame)} | Enlaces duplicados: {frame['Link'].duplicated().sum()}")
    print(f"Tiempo total: {time.time() - started:.1f} segundos")


if __name__ == "__main__":
    main()
