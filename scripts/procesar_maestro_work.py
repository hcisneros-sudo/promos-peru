from __future__ import annotations

import argparse
import gzip
import hashlib
import json
import os
import re
import time
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timezone
from pathlib import Path


TEMPORARY_HTTP_CODES = {429, 500, 502, 503, 504}


def text(value) -> str:
    return "" if value is None else str(value).strip()


def normalized(value) -> str:
    value = unicodedata.normalize("NFD", text(value).lower())
    value = "".join(char for char in value if unicodedata.category(char) != "Mn")
    return re.sub(r"[^a-z0-9]+", " ", value).strip()


def write_json(path: Path, payload) -> None:
    """Escribe JSON de forma atómica para no perder un checkpoint parcial."""
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f"{path.name}.tmp")
    temporary.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    temporary.replace(path)


def http_error_body(error: urllib.error.HTTPError) -> str:
    try:
        return error.read().decode("utf-8", errors="replace")[:2000]
    except Exception:
        return ""


def load_catalog(folder: Path) -> list[dict]:
    manifest_path = folder / "catalog-manifest.json"
    if not manifest_path.exists():
        return []
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    payload = b"".join((folder / "catalog.parts" / name).read_bytes() for name in manifest.get("parts", []))
    return json.loads(gzip.decompress(payload)) if payload else []


def coordinate_history(catalog: list[dict]) -> dict[str, dict]:
    grouped: dict[str, list[dict]] = {}
    for promotion in catalog:
        for location in promotion.get("locations", []):
            if location.get("lat") is None or location.get("lng") is None:
                continue
            key = normalized(location.get("address"))
            if key:
                grouped.setdefault(key, []).append(location)
    history = {}
    for key, values in grouped.items():
        coords = {(round(float(v["lat"]), 6), round(float(v["lng"]), 6)) for v in values}
        if len(coords) == 1:
            history[key] = values[0]
    return history


def benefit_label(promotion: dict, benefit: dict) -> str:
    kind = benefit.get("tipo")
    value = benefit.get("valor")
    if kind == "descuento_porcentual" and value is not None:
        return f"{'Hasta ' if benefit.get('es_hasta') else ''}{value:g}% de descuento"
    if kind == "descuento_monto" and value is not None:
        currency = "S/" if benefit.get("unidad") == "PEN" else benefit.get("unidad", "")
        return f"{'Hasta ' if benefit.get('es_hasta') else ''}{currency}{value:g} de descuento"
    labels = {
        "cuotas_sin_interes": "Cuotas sin intereses",
        "nxm": "Promoción especial",
        "precio_especial": "Precio especial",
        "cashback_recompensa": "Cashback o devolución",
        "sorteo_premio": "Sorteo o premio",
        "producto_gratis": "Producto gratis",
        "puntos_millas": "Puntos o millas",
    }
    return labels.get(kind, text(promotion.get("nombre")))


DAY_NAMES = {
    "lunes": "Lunes", "martes": "Martes", "miercoles": "Miércoles", "miércoles": "Miércoles",
    "jueves": "Jueves", "viernes": "Viernes", "sabado": "Sábado", "sábado": "Sábado", "domingo": "Domingo",
}


def days_from_source(value) -> list[str]:
    source = normalized(value)
    if not source:
        return []
    order = list(DAY_NAMES)
    result = []
    for key in order:
        if re.search(rf"\b{re.escape(key)}\b", source) and DAY_NAMES[key] not in result:
            result.append(DAY_NAMES[key])
    ranges = [("lunes", "viernes"), ("lunes", "sabado"), ("lunes", "domingo"), ("martes", "domingo")]
    for start, end in ranges:
        if re.search(rf"\b{start}\s+a\s+{end}\b", source):
            canonical = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
            i, j = canonical.index(DAY_NAMES[start]), canonical.index(DAY_NAMES[end])
            result = canonical[i:j + 1]
    return result


def query_geoapify(query: str, api_key: str) -> dict | None:
    params = urllib.parse.urlencode({
        "text": query,
        "filter": "countrycode:pe",
        "format": "json",
        "limit": 1,
        "apiKey": api_key,
    })
    request = urllib.request.Request(
        f"https://api.geoapify.com/v1/geocode/search?{params}",
        headers={"User-Agent": "PromosPeru/1.0"},
    )
    for attempt in range(3):
        try:
            with urllib.request.urlopen(request, timeout=30) as response:
                results = json.load(response).get("results", [])
            return results[0] if results else None
        except urllib.error.HTTPError as error:
            if error.code not in TEMPORARY_HTTP_CODES or attempt == 2:
                raise
            wait = 3 * (attempt + 1)
            print(f"Geoapify respondió HTTP {error.code}; reintento {attempt + 1}/2 en {wait}s.")
            time.sleep(wait)
        except (urllib.error.URLError, TimeoutError) as error:
            if attempt == 2:
                raise
            wait = 3 * (attempt + 1)
            print(f"Error temporal de red ({type(error).__name__}); reintento {attempt + 1}/2 en {wait}s.")
            time.sleep(wait)
    return None


def apply_history_to_master(promotions: list[dict], history: dict[str, dict]) -> int:
    applied = 0
    for promotion in promotions:
        for location in promotion.get("ubicacion", {}).get("locales", []):
            if location.get("tipo") != "local_fisico" or (location.get("latitud") is not None and location.get("longitud") is not None):
                continue
            previous = history.get(normalized(location.get("direccion") or location.get("nombre")))
            if previous:
                location["latitud"], location["longitud"] = previous.get("lat"), previous.get("lng")
                location["estado_geocodificacion"] = "Historial del catálogo"
                applied += 1
    return applied


def geocode_locations(
    promotions: list[dict],
    cache: dict,
    api_key: str,
    requery_all: bool,
    cache_path: Path,
    errors_path: Path,
) -> tuple[int, int, int]:
    requested = reused = failed = 0
    errors: list[dict] = []
    checkpoint_pending = 0

    targets: list[tuple[dict, dict]] = []
    for promotion in promotions:
        for location in promotion.get("ubicacion", {}).get("locales", []):
            if location.get("tipo") != "local_fisico" or not location.get("apta_geoapify"):
                continue
            if not requery_all and location.get("latitud") is not None and location.get("longitud") is not None:
                continue
            targets.append((promotion, location))

    total = len(targets)
    print(f"Direcciones pendientes de geocodificación: {total}")

    try:
        for position, (promotion, location) in enumerate(targets, 1):
            query = text(location.get("consulta_geoapify"))
            if not query:
                continue

            key = hashlib.sha256(query.encode("utf-8")).hexdigest()
            result = cache.get(key)

            if result:
                reused += 1
                print(f"[{position}/{total}] caché: {query[:140]}")
            elif api_key:
                requested += 1
                print(f"[{position}/{total}] Geoapify: {query[:180]}")
                try:
                    result = query_geoapify(query, api_key)
                    cache[key] = result
                    checkpoint_pending += 1
                except urllib.error.HTTPError as error:
                    failed += 1
                    body = http_error_body(error)
                    location["estado_geocodificacion"] = f"Error Geoapify HTTP {error.code}"
                    errors.append({
                        "id_promocion": promotion.get("id"),
                        "nombre_promocion": promotion.get("nombre"),
                        "id_local": location.get("id_local"),
                        "nombre_local": location.get("nombre"),
                        "direccion": location.get("direccion"),
                        "consulta": query,
                        "longitud_consulta": len(query),
                        "codigo_http": error.code,
                        "motivo": str(error.reason or "HTTPError"),
                        "respuesta": body,
                    })
                    print(f"[{position}/{total}] ERROR HTTP {error.code}; se continúa. Consulta: {query[:180]}")
                    write_json(errors_path, errors)
                    continue
                except (urllib.error.URLError, TimeoutError) as error:
                    failed += 1
                    location["estado_geocodificacion"] = f"Error de red Geoapify: {type(error).__name__}"
                    errors.append({
                        "id_promocion": promotion.get("id"),
                        "nombre_promocion": promotion.get("nombre"),
                        "id_local": location.get("id_local"),
                        "nombre_local": location.get("nombre"),
                        "direccion": location.get("direccion"),
                        "consulta": query,
                        "longitud_consulta": len(query),
                        "codigo_http": None,
                        "motivo": f"{type(error).__name__}: {error}",
                        "respuesta": "",
                    })
                    print(f"[{position}/{total}] ERROR DE RED; se continúa. Consulta: {query[:180]}")
                    write_json(errors_path, errors)
                    continue

                if checkpoint_pending >= 20:
                    write_json(cache_path, cache)
                    write_json(errors_path, errors)
                    checkpoint_pending = 0
                    print(f"Checkpoint guardado: {requested} consultas nuevas.")
                time.sleep(0.25)

            if not result:
                continue

            location["latitud"] = result.get("lat")
            location["longitud"] = result.get("lon")
            location["direccion_geoapify"] = result.get("formatted")
            location["departamento_geoapify"] = result.get("state")
            location["provincia_geoapify"] = result.get("county")
            location["distrito_geoapify"] = result.get("city") or result.get("district") or result.get("suburb")
            location["confianza_geoapify"] = (result.get("rank") or {}).get("confidence")
            location["estado_geocodificacion"] = "Geoapify"
    finally:
        write_json(cache_path, cache)
        write_json(errors_path, errors)

    print(f"Geoapify: {requested} consultas nuevas, {reused} reutilizadas, {failed} errores registrados.")
    return requested, reused, failed


def compact(master: dict, history: dict[str, dict], allow_history: bool = True) -> tuple[list[dict], int]:
    catalog = []
    migrated = 0
    for p in master.get("promociones", []):
        trace = p.get("trazabilidad", {})
        restrictions = p.get("restricciones", {})
        vigencia = p.get("vigencia", {})
        locations = []
        for source in p.get("ubicacion", {}).get("locales", []):
            address = text(source.get("direccion")) or text(source.get("nombre"))
            lat, lng = source.get("latitud"), source.get("longitud")
            previous = history.get(normalized(address)) if allow_history else None
            if (lat is None or lng is None) and previous:
                lat, lng = previous.get("lat"), previous.get("lng")
                migrated += 1
            locations.append({
                "id": source.get("id_local"), "type": source.get("tipo"), "name": source.get("nombre"),
                "address": address, "district": source.get("distrito_geoapify") or source.get("distrito") or "",
                "province": source.get("provincia_geoapify") or source.get("provincia") or "",
                "department": source.get("departamento_geoapify") or source.get("departamento") or "",
                "lat": lat, "lng": lng, "mapStatus": source.get("estado_geocodificacion") or "",
                "geoQuery": source.get("consulta_geoapify"), "geocodable": bool(source.get("apta_geoapify")),
            })
        discount = p.get("beneficios", [{}])[0].get("valor") if p.get("beneficios") else None
        if p.get("beneficios", [{}])[0].get("tipo") != "descuento_porcentual":
            discount = None
        first_benefit = p.get("beneficios", [{}])[0] if p.get("beneficios") else {}
        catalog.append({
            "id": p.get("id"), "bank": p.get("banco"), "name": p.get("nombre"), "merchant": p.get("comercio"),
            "category": p.get("categoria"), "subcategory": p.get("subcategoria"), "sourceCategory": p.get("categoria_original"),
            "benefitType": first_benefit.get("tipo"), "benefit": benefit_label(p, first_benefit),
            "discount": discount, "detail": p.get("detalle"), "startDate": vigencia.get("inicio"), "endDate": vigencia.get("fin"),
            "status": vigencia.get("estado_original"), "validDaysText": vigencia.get("dias_validos"), "days": days_from_source(vigencia.get("dias_validos")),
            "paymentMethod": p.get("medio_pago"), "channels": p.get("canales_validos") or [], "couponCode": p.get("codigo_cupon"),
            "purchaseMinimum": restrictions.get("monto_minimo_compra"), "discountCap": restrictions.get("tope_descuento"),
            "accumulable": restrictions.get("acumulable"), "requiresCoupon": restrictions.get("requiere_cupon"),
            "requiresReservation": restrictions.get("requiere_reserva"), "newCustomersOnly": restrictions.get("solo_nuevos_clientes"),
            "subjectToStock": restrictions.get("sujeto_stock"), "restrictions": p.get("restricciones_originales"), "terms": p.get("terminos"),
            "link": p.get("url"), "purchaseLink": p.get("enlace_compra"), "image": p.get("imagen"), "logo": p.get("logo"),
            "needsReview": trace.get("requiere_revision", False), "reviewReason": trace.get("motivo_revision"),
            "sourceFile": trace.get("archivo_origen"), "locations": locations,
        })
    catalog.sort(key=lambda p: (text(p.get("bank")), text(p.get("merchant")), text(p.get("name"))))
    return catalog, migrated


def write_catalog(catalog: list[dict], output: Path, part_size: int, migrated: int, geocode_stats: tuple[int, int, int]) -> dict:
    raw = json.dumps(catalog, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    compressed = gzip.compress(raw, compresslevel=9, mtime=0)
    parts = output / "catalog.parts"
    parts.mkdir(parents=True, exist_ok=True)
    for old in parts.glob("part-*.bin"):
        old.unlink()
    names = []
    for index, start in enumerate(range(0, len(compressed), part_size), 1):
        name = f"part-{index:03d}.bin"
        (parts / name).write_bytes(compressed[start:start + part_size])
        names.append(name)
    mapped = sum(1 for p in catalog for loc in p["locations"] if loc["lat"] is not None and loc["lng"] is not None and loc["type"] == "local_fisico")
    manifest = {
        "version": 3,
        "generatedAt": datetime.now(timezone.utc).isoformat(),
        "source": "ChatGPT Work master",
        "promotions": len(catalog),
        "mappedLocations": mapped,
        "migratedCoordinates": migrated,
        "geoapifyRequested": geocode_stats[0],
        "geoapifyReused": geocode_stats[1],
        "geoapifyErrors": geocode_stats[2],
        "uncompressedBytes": len(raw),
        "compressedBytes": len(compressed),
        "sha256": hashlib.sha256(compressed).hexdigest(),
        "parts": names,
    }
    (output / "catalog-manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return manifest


def main() -> int:
    parser = argparse.ArgumentParser(description="Publica el maestro preparado por ChatGPT Work sin reinterpretar direcciones.")
    parser.add_argument("--maestro", type=Path, required=True)
    parser.add_argument("--salida", type=Path, required=True)
    parser.add_argument("--catalogo-actual", type=Path, default=Path("site/data"))
    parser.add_argument("--cache", type=Path, default=Path("data/geocoding/work_queries.json"))
    parser.add_argument("--errores-geoapify", type=Path, default=Path("reports/errores_geoapify.json"))
    parser.add_argument("--regeocodificar-todo", action="store_true")
    parser.add_argument("--tamano-parte", type=int, default=96_000)
    args = parser.parse_args()

    if args.maestro.suffix == ".gz":
        with gzip.open(args.maestro, mode="rt", encoding="utf-8") as fh:
            master = json.load(fh)
    else:
        master = json.loads(args.maestro.read_text(encoding="utf-8"))

    old = load_catalog(args.catalogo_actual)
    history = coordinate_history(old)
    history_applied = 0 if args.regeocodificar_todo else apply_history_to_master(master.get("promociones", []), history)
    cache = json.loads(args.cache.read_text(encoding="utf-8")) if args.cache.exists() else {}
    if args.regeocodificar_todo:
        cache = {}
    api_key = os.environ.get("GEOAPIFY_API_KEY", "")

    geocode_stats = geocode_locations(
        master.get("promociones", []),
        cache,
        api_key,
        args.regeocodificar_todo,
        args.cache,
        args.errores_geoapify,
    )

    catalog, migrated = compact(master, history, allow_history=not args.regeocodificar_todo)
    manifest = write_catalog(catalog, args.salida, args.tamano_parte, migrated + history_applied, geocode_stats)
    print(json.dumps(manifest, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
