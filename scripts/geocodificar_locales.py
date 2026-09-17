from __future__ import annotations

import argparse
import hashlib
import json
import math
import os
import re
import sqlite3
import time
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path


NOMINATIM_ENDPOINT = "https://nominatim.openstreetmap.org/search"
GEOAPIFY_ENDPOINT = "https://api.geoapify.com/v1/geocode/search"
PERU_BOUNDS = {"lat_min": -18.6, "lat_max": 0.6, "lon_min": -82.0, "lon_max": -68.0}
STOPWORDS = {"av", "avenida", "calle", "jr", "jiron", "nro", "numero", "local", "tienda", "interior", "int", "piso", "nivel", "peru"}
PROVIDER_LABELS = {
    "geoapify": "Geoapify / OpenStreetMap",
    "nominatim": "Nominatim / OpenStreetMap",
}


def fold(value: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", value.lower()) if not unicodedata.combining(c))


def tokens(value: str) -> set[str]:
    return {part for part in re.findall(r"[a-z0-9]+", fold(value)) if len(part) > 1 and part not in STOPWORDS}


def clean_address_for_query(value: str) -> str:
    text = unicodedata.normalize("NFKC", str(value or ""))
    text = re.sub(r"[\u200b-\u200f\u2060\ufeff]", "", text)
    text = re.sub(r"\b(?:avenida\b|av(?=\s*\.|\s))\s*\.*\s*", "Av. ", text, flags=re.I)
    text = re.sub(r"\b(?:jir[oó]n\b|jr(?=\s*\.|\s))\s*\.*\s*", "Jr. ", text, flags=re.I)
    text = re.sub(r"\b(?:n(?:ro|um(?:ero)?)?(?:\s*[.°º])+|nro\.?|#)\s*(?=\d)", "", text, flags=re.I)
    text = re.sub(r"\.{2,}", ".", text)
    text = re.sub(r"\bSurcol\b", "Surco", text, flags=re.I)
    text = re.sub(r"\bJes[uú]s Mar[ií]a[\"']?l\b", "Jesús María", text, flags=re.I)
    text = re.sub(r"\bProlog(?:aci[oó]n|\.)?\b", "Prolongación", text, flags=re.I)
    text = re.sub(r"\bMall\s+Av\.\s*entura\b", "Mall Aventura", text, flags=re.I)
    text = re.sub(r"\s+", " ", text).strip(" ,;:-")
    if ":" in text:
        left, right = text.split(":", 1)
        if re.search(r"\b(?:Av\.|Jr\.|Calle|Carretera|Prolongaci[oó]n|Pasaje|Malec[oó]n)\b|\d{2,}", right, flags=re.I):
            text = right.strip(" ,;:-")
    return text


def normalize_query(location: dict) -> str:
    parts = [
        clean_address_for_query(location.get("direccion_limpia", "")), location.get("distrito", ""),
        location.get("provincia", ""), location.get("departamento", ""), "Perú",
    ]
    text = ", ".join(str(part).strip(" ,") for part in parts if str(part).strip(" ,"))
    return re.sub(r"\s+", " ", text).strip()


def strip_unit_noise(value: str) -> str:
    """Retira tienda/interior/piso sin borrar una vía que aparezca después."""
    text = clean_address_for_query(value)
    text = re.sub(
        r"(?:\s*[-–—,]\s*)?\b(?:tiendas?|tda\.?|local(?:es)?|loc\.?|int(?:erior)?\.?|oficina|stand|nivel|piso)\s*[A-Z]?[\w./-]*(?:\s*(?:,|y)\s*[A-Z]?[\w./-]+)*",
        " ", text, flags=re.I,
    )
    text = re.sub(r"\b(?:primer|segundo|tercer|1er|2do|3er)\s+(?:nivel|piso)\b", " ", text, flags=re.I)
    return re.sub(r"\s+", " ", text).strip(" ,;:-–—")


def route_segment(value: str) -> str:
    """Prioriza la vía y número cuando el texto inicia con el nombre del centro comercial."""
    text = strip_unit_noise(value)
    matches = list(re.finditer(r"\b(?:Av\.|Avenida|Jr\.|Jir[oó]n|Calle|Carretera|Prolongaci[oó]n|Pasaje|Malec[oó]n)\s+", text, flags=re.I))
    if not matches:
        return text
    for match in reversed(matches):
        segment = text[match.start():].strip(" ,;:-–—")
        if re.search(r"\b\d+[A-Za-z]?\b|\bS\s*/?\s*N\b", segment, flags=re.I):
            return segment
    return text


def place_match(location: dict, known_places: list[dict]) -> dict | None:
    haystack = fold(" ".join(str(location.get(field, "")) for field in ("nombre_local", "direccion_limpia", "direccion_original")))
    for place in known_places:
        aliases = place.get("alias", [])
        if any(re.search(rf"\b{re.escape(fold(alias))}\b", haystack) for alias in aliases if alias):
            return place
        required = [fold(part) for part in place.get("contiene", []) if part]
        if required and all(re.search(rf"\b{re.escape(part)}\b", haystack) for part in required):
            return place
    return None


def enrich_from_known_place(location: dict, place: dict | None) -> dict:
    enriched = dict(location)
    if not place:
        return enriched
    for field in ("distrito", "provincia", "departamento"):
        if place.get(field):
            enriched[field] = place[field]
    return enriched


def query_variants(location: dict, known_place: dict | None = None) -> list[str]:
    """Consultas progresivas: vía, localidad y establecimiento conocido."""
    location = enrich_from_known_place(location, known_place)
    primary = normalize_query(location)
    variants = [primary]
    address = clean_address_for_query(location.get("direccion_limpia", ""))
    simplified = strip_unit_noise(address)
    if simplified and simplified != address:
        fallback_location = dict(location)
        fallback_location["direccion_limpia"] = simplified
        variants.append(normalize_query(fallback_location))

    street = route_segment(address)
    if street and street != address and street != simplified:
        fallback_location = dict(location)
        fallback_location["direccion_limpia"] = street
        variants.append(normalize_query(fallback_location))

    if known_place and known_place.get("consulta"):
        variants.append(clean_address_for_query(known_place["consulta"]))

    place = clean_address_for_query(location.get("nombre_local", ""))
    if place and re.search(r"\b(?:mall|plaza|centro comercial|c\.?\s*c\.?)\b", fold(place)):
        parts = [place, location.get("distrito", ""), location.get("provincia", ""), "Perú"]
        place_query = ", ".join(str(part).strip(" ,") for part in parts if str(part).strip(" ,"))
        variants.append(re.sub(r"\s+", " ", place_query).strip())

    unique = []
    for variant in variants:
        if variant and query_digest(variant) not in {query_digest(item) for item in unique}:
            unique.append(variant)
    return unique


def load_known_places(path: Path) -> list[dict]:
    if not path.exists():
        return []
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise RuntimeError(f"No se pudo leer el diccionario de lugares: {path}") from error
    if not isinstance(payload, list):
        raise RuntimeError("El diccionario de lugares debe contener una lista JSON.")
    return payload


def query_digest(query: str) -> str:
    """Clave estable: ignora mayúsculas, tildes y espacios accidentales."""
    canonical = re.sub(r"\s+", " ", fold(query)).strip(" ,")
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()


def within_peru(lat: float, lon: float) -> bool:
    return PERU_BOUNDS["lat_min"] <= lat <= PERU_BOUNDS["lat_max"] and PERU_BOUNDS["lon_min"] <= lon <= PERU_BOUNDS["lon_max"]


def score_candidate(query: str, location: dict, candidate: dict) -> float:
    query_tokens, result_tokens = tokens(query), tokens(candidate.get("display_name", ""))
    similarity = len(query_tokens & result_tokens) / max(1, len(query_tokens | result_tokens))
    address = candidate.get("address") or {}
    country_bonus = 0.15 if address.get("country_code") == "pe" else -1.0
    expected_places = [fold(location.get(field, "")) for field in ("distrito", "provincia", "departamento")]
    expected_places = [place for place in expected_places if place]
    address_text = fold(" ".join(str(value) for value in address.values()))
    locality_hits = sum(place in address_text for place in expected_places)
    locality_bonus = min(0.24, 0.12 * locality_hits) if expected_places else 0
    if expected_places and not locality_hits:
        locality_bonus = -0.22
    query_numbers = set(re.findall(r"\b\d{1,5}[a-z]?\b", fold(query)))
    result_numbers = set(re.findall(r"\b\d{1,5}[a-z]?\b", fold(candidate.get("display_name", ""))))
    number_bonus = 0.18 if query_numbers & result_numbers else (-0.12 if query_numbers else 0)
    type_bonus = 0.06 if candidate.get("addresstype") in {"house", "building", "shop", "amenity", "office"} else 0
    try:
        provider_bonus = 0.10 * float(candidate.get("provider_confidence") or 0)
    except (TypeError, ValueError):
        provider_bonus = 0
    # Escala intuitiva de 0 a 1; los bonos pueden llevar la suma bruta por encima de 1.
    return round(max(0.0, min(1.0, similarity + country_bonus + locality_bonus + number_bonus + type_bonus + provider_bonus)), 4)


class Cache:
    def __init__(self, path: Path):
        self.connection = sqlite3.connect(path)
        self.connection.execute("""
            CREATE TABLE IF NOT EXISTS geocoding_cache (
                query_hash TEXT PRIMARY KEY,
                query TEXT NOT NULL,
                response_json TEXT NOT NULL,
                provider TEXT NOT NULL,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)
        self.connection.execute("""
            CREATE TABLE IF NOT EXISTS provider_cache (
                provider TEXT NOT NULL,
                query_hash TEXT NOT NULL,
                query TEXT NOT NULL,
                response_json TEXT NOT NULL,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                PRIMARY KEY (provider, query_hash)
            )
        """)
        self.connection.execute("""
            CREATE TABLE IF NOT EXISTS geocoding_resolutions (
                query_hash TEXT PRIMARY KEY,
                query TEXT NOT NULL,
                candidate_json TEXT NOT NULL,
                score REAL NOT NULL,
                provider TEXT NOT NULL,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                last_used_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)
        # Conserva las respuestas de versiones anteriores del programa.
        legacy_rows = self.connection.execute(
            "SELECT query, response_json, provider FROM geocoding_cache"
        ).fetchall()
        for query, response_json, provider_label in legacy_rows:
            provider = "nominatim" if "nominatim" in provider_label.lower() else provider_label.lower()
            self.connection.execute(
                "INSERT OR IGNORE INTO provider_cache (provider, query_hash, query, response_json) VALUES (?, ?, ?, ?)",
                (provider, query_digest(query), query, response_json),
            )
        self.connection.commit()

    def get_response(self, provider: str, query: str):
        row = self.connection.execute(
            "SELECT response_json FROM provider_cache WHERE provider = ? AND query_hash = ?",
            (provider, query_digest(query)),
        ).fetchone()
        return json.loads(row[0]) if row else None

    def put_response(self, provider: str, query: str, response: list):
        self.connection.execute(
            "INSERT OR REPLACE INTO provider_cache (provider, query_hash, query, response_json) VALUES (?, ?, ?, ?)",
            (provider, query_digest(query), query, json.dumps(response, ensure_ascii=False)),
        )
        self.connection.commit()

    def get_resolution(self, query: str):
        row = self.connection.execute(
            "SELECT candidate_json, score, provider FROM geocoding_resolutions WHERE query_hash = ?",
            (query_digest(query),),
        ).fetchone()
        if not row:
            return None
        self.connection.execute(
            "UPDATE geocoding_resolutions SET last_used_at = CURRENT_TIMESTAMP WHERE query_hash = ?",
            (query_digest(query),),
        )
        self.connection.commit()
        return json.loads(row[0]), float(row[1]), row[2]

    def put_resolution(self, query: str, candidate: dict, score: float, provider: str):
        self.connection.execute(
            """
            INSERT INTO geocoding_resolutions
                (query_hash, query, candidate_json, score, provider)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(query_hash) DO UPDATE SET
                query = excluded.query,
                candidate_json = excluded.candidate_json,
                score = excluded.score,
                provider = excluded.provider,
                last_used_at = CURRENT_TIMESTAMP
            """,
            (query_digest(query), query, json.dumps(candidate, ensure_ascii=False), score, provider),
        )
        self.connection.commit()

    def count_resolutions(self) -> int:
        return int(self.connection.execute("SELECT COUNT(*) FROM geocoding_resolutions").fetchone()[0])

    def close(self):
        self.connection.close()


def request_nominatim(query: str, user_agent: str, email: str, delay: float) -> list:
    parameters = {
        "q": query, "format": "jsonv2", "addressdetails": 1, "limit": 3,
        "countrycodes": "pe", "viewbox": "-82.0,0.6,-68.0,-18.6", "bounded": 1,
        "accept-language": "es",
    }
    if email:
        parameters["email"] = email
    url = f"{NOMINATIM_ENDPOINT}?{urllib.parse.urlencode(parameters)}"
    request = urllib.request.Request(url, headers={"User-Agent": user_agent, "Accept": "application/json"})
    for attempt in range(3):
        try:
            time.sleep(delay)
            with urllib.request.urlopen(request, timeout=30) as response:
                return json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as error:
            if error.code not in {429, 500, 502, 503, 504} or attempt == 2:
                raise
            time.sleep(5 * (attempt + 1))
        except urllib.error.URLError:
            if attempt == 2:
                raise
            time.sleep(5 * (attempt + 1))
    return []


def request_geoapify(query: str, api_key: str, delay: float) -> list:
    parameters = {
        "text": query,
        "format": "json",
        "lang": "es",
        "limit": 3,
        "filter": "countrycode:pe",
        "apiKey": api_key,
    }
    url = f"{GEOAPIFY_ENDPOINT}?{urllib.parse.urlencode(parameters)}"
    request = urllib.request.Request(url, headers={"Accept": "application/json"})
    for attempt in range(3):
        try:
            time.sleep(delay)
            with urllib.request.urlopen(request, timeout=30) as response:
                payload = json.loads(response.read().decode("utf-8"))
                return [normalize_geoapify_candidate(row) for row in payload.get("results", [])]
        except urllib.error.HTTPError as error:
            if error.code not in {429, 500, 502, 503, 504} or attempt == 2:
                raise
            time.sleep(3 * (attempt + 1))
        except urllib.error.URLError:
            if attempt == 2:
                raise
            time.sleep(3 * (attempt + 1))
    return []


def normalize_geoapify_candidate(row: dict) -> dict:
    rank = row.get("rank") or {}
    return {
        "lat": row.get("lat"),
        "lon": row.get("lon"),
        "display_name": row.get("formatted", ""),
        "addresstype": row.get("result_type", ""),
        "address": {
            "country_code": row.get("country_code", ""),
            "country": row.get("country", ""),
            "state": row.get("state", ""),
            "province": row.get("county", ""),
            "county": row.get("county", ""),
            "city": row.get("city", ""),
            "city_district": row.get("district", ""),
            "suburb": row.get("suburb", ""),
            "road": row.get("street", ""),
            "house_number": row.get("housenumber", ""),
            "postcode": row.get("postcode", ""),
        },
        "provider_confidence": rank.get("confidence"),
        "match_type": rank.get("match_type", ""),
        "place_id": row.get("place_id", ""),
    }


def choose_candidate(query: str, location: dict, candidates: list) -> tuple[dict | None, float]:
    scored = []
    for candidate in candidates:
        try:
            lat, lon = float(candidate["lat"]), float(candidate["lon"])
        except (KeyError, TypeError, ValueError):
            continue
        if not within_peru(lat, lon):
            continue
        scored.append((score_candidate(query, location, candidate), candidate))
    if not scored:
        return None, 0
    score, candidate = max(scored, key=lambda item: item[0])
    return candidate, score


def apply_candidate(location: dict, query: str, candidate: dict, score: float, provider: str):
    address = candidate.get("address") or {}
    location["latitud"] = float(candidate["lat"])
    location["longitud"] = float(candidate["lon"])
    location["direccion_geocodificada"] = candidate.get("display_name", "")
    location["consulta_geocodificacion"] = query
    location["puntaje_geocodificacion"] = score
    location["fuente_coordenada"] = PROVIDER_LABELS.get(provider, provider)
    location["precision_ubicacion"] = "Geocodificada"
    location["confianza_ubicacion"] = "Alta" if score >= 0.70 else "Media"
    if not location.get("distrito"):
        location["distrito"] = address.get("city_district") or address.get("suburb") or address.get("quarter") or ""
    if not location.get("provincia"):
        location["provincia"] = address.get("province") or address.get("county") or address.get("city") or ""
    if not location.get("departamento"):
        location["departamento"] = address.get("state") or ""


def import_previous_resolutions(cache: Cache, output_path: Path) -> int:
    """Migra resultados de una ejecución anterior al historial permanente."""
    if not output_path.exists():
        return 0
    try:
        previous = json.loads(output_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return 0
    imported = 0
    for row in previous:
        query = row.get("consulta_geocodificacion")
        if not query or row.get("latitud") is None or row.get("longitud") is None:
            continue
        source = fold(row.get("fuente_coordenada", ""))
        provider = "geoapify" if "geoapify" in source else "nominatim"
        candidate = {
            "lat": row["latitud"],
            "lon": row["longitud"],
            "display_name": row.get("direccion_geocodificada", ""),
            "addresstype": "building" if row.get("confianza_ubicacion") == "Alta" else "street",
            "address": {
                "country_code": "pe",
                "state": row.get("departamento", ""),
                "province": row.get("provincia", ""),
                "city_district": row.get("distrito", ""),
            },
        }
        cache.put_resolution(query, candidate, float(row.get("puntaje_geocodificacion") or 0.42), provider)
        imported += 1
    return imported


def main():
    parser = argparse.ArgumentParser(description="Geocodifica los locales consolidados con caché y revisión de confianza.")
    parser.add_argument("--entrada", default="salida/locales.json")
    parser.add_argument("--salida", default="salida/locales_geocodificados.json")
    parser.add_argument("--revision", default="salida/revision_geocodificacion.json")
    parser.add_argument("--geojson", default="salida/locales_mapa.geojson", help="Archivo geográfico para la plataforma web.")
    parser.add_argument("--cache", default="salida/geocodificacion_cache.sqlite")
    parser.add_argument("--lugares", default=str(Path(__file__).with_name("lugares_conocidos.json")), help="Diccionario JSON de centros comerciales y locales conocidos.")
    parser.add_argument("--proveedor", choices=["geoapify", "nominatim"], default="geoapify", help="Servicio usado solo para direcciones aún no guardadas.")
    parser.add_argument("--api-key", default=os.getenv("GEOAPIFY_API_KEY", ""), help="Clave de Geoapify; preferiblemente use GEOAPIFY_API_KEY.")
    parser.add_argument("--limite", type=int, default=0, help="Máximo de consultas nuevas. Cero procesa todas.")
    parser.add_argument("--pausa", type=float, default=None, help="Segundos entre consultas. Por defecto: 0.22 Geoapify y 1.1 Nominatim.")
    parser.add_argument("--solo-cache", action="store_true", help="No realiza consultas de red; aplica solo resultados guardados.")
    parser.add_argument("--acepto-politica", action="store_true", help="Confirma que se revisó la política pública de Nominatim.")
    parser.add_argument("--email", default=os.getenv("GEOCODER_EMAIL", ""), help="Correo de contacto opcional; también admite GEOCODER_EMAIL.")
    parser.add_argument("--umbral", type=float, default=0.70, help="Puntaje mínimo para aceptar automáticamente una coordenada (predeterminado: 0.70).")
    parser.add_argument("--reintentos", type=int, default=3, help="Consultas alternativas por dirección dudosa (predeterminado: 3).")
    parser.add_argument("--desde-cero", action="store_true", help="Ignora resultados y respuestas anteriores; realiza consultas nuevas. Use además un archivo --cache nuevo para separar completamente la prueba.")
    args = parser.parse_args()
    args.pausa = args.pausa if args.pausa is not None else (0.22 if args.proveedor == "geoapify" else 1.1)
    if not args.solo_cache and args.proveedor == "geoapify" and not args.api_key:
        parser.error("Falta GEOAPIFY_API_KEY. Créela gratis en https://myprojects.geoapify.com/ y defínala como variable de entorno.")
    if not args.solo_cache and args.proveedor == "nominatim" and not args.acepto_politica:
        parser.error("Para consultar Nominatim debe usar --acepto-politica después de revisar https://operations.osmfoundation.org/policies/nominatim/")
    if not args.solo_cache and args.proveedor == "nominatim" and args.pausa < 1.1:
        parser.error("La pausa mínima permitida por este programa es 1.1 segundos.")
    if args.pausa < 0:
        parser.error("La pausa no puede ser negativa.")
    if args.reintentos < 0 or args.reintentos > 4:
        parser.error("--reintentos debe estar entre 0 y 4.")

    locations = json.loads(Path(args.entrada).read_text(encoding="utf-8"))
    known_places = load_known_places(Path(args.lugares))
    cache = Cache(Path(args.cache))
    imported_count = 0 if args.desde_cero else import_previous_resolutions(cache, Path(args.salida))
    reviews, queried, accepted, cached_count, persistent_count = [], 0, 0, 0, 0
    user_agent = "PromosPeruGeocoder/1.0 (+https://github.com/; address-cleaning-project)"
    unique_results = {}

    def eligible(row: dict) -> bool:
        if row.get("latitud") is not None:
            return False
        text = " ".join(str(row.get(field, "")) for field in ("nombre_local", "direccion_limpia", "direccion_original"))
        return bool(
            row.get("confianza_ubicacion") != "Revisar"
            or re.search(r"\b(?:Av\.?|Avenida|Jr\.?|Jir[oó]n|Calle|Carretera|Prolongaci[oó]n|Pasaje|Malec[oó]n)\b", text, flags=re.I)
            or re.search(r"\d{2,}", text)
            or place_match(row, known_places)
        )

    candidates = [row for row in locations if eligible(row)]
    total = len(candidates)
    for position, location in enumerate(candidates, 1):
        known_place = place_match(location, known_places)
        scoring_location = enrich_from_known_place(location, known_place)
        if known_place:
            for field in ("distrito", "provincia", "departamento"):
                if not location.get(field) and scoring_location.get(field):
                    location[field] = scoring_location[field]
        variants = query_variants(scoring_location, known_place)[:1 + args.reintentos]
        primary_query = variants[0]
        primary_key = fold(primary_query)
        saved = None if args.desde_cero else cache.get_resolution(primary_query)
        if saved and saved[1] >= args.umbral:
            candidate, score, saved_provider = saved
            apply_candidate(location, primary_query, candidate, score, saved_provider)
            persistent_count += 1
            accepted += 1
            unique_results[primary_key] = [candidate]
            print(f"[{position}/{total}] reutilizada: {primary_query[:65]}")
            continue
        best_candidate, best_score, best_query = None, 0.0, primary_query
        last_error = ""
        attempted = []
        for query in variants:
            query_key = fold(query)
            response = unique_results.get(query_key)
            from_cache = False
            if response is None:
                response = None if args.desde_cero else cache.get_response(args.proveedor, query)
                from_cache = response is not None
            if response is None:
                if args.solo_cache or (args.limite and queried >= args.limite):
                    continue
                try:
                    if args.proveedor == "geoapify":
                        response = request_geoapify(query, args.api_key, args.pausa)
                    else:
                        response = request_nominatim(query, user_agent, args.email, args.pausa)
                    cache.put_response(args.proveedor, query, response)
                    queried += 1
                except Exception as error:
                    last_error = f"Error de consulta: {type(error).__name__}"
                    continue
            elif from_cache:
                cached_count += 1
            unique_results[query_key] = response
            attempted.append(query)
            candidate, score = choose_candidate(query, scoring_location, response)
            if candidate and score > best_score:
                best_candidate, best_score, best_query = candidate, score, query
            if candidate and score >= args.umbral:
                break

        if best_candidate and best_score >= args.umbral:
            apply_candidate(location, best_query, best_candidate, best_score, args.proveedor)
            cache.put_resolution(primary_query, best_candidate, best_score, args.proveedor)
            accepted += 1
            suffix = " reintento" if best_query != primary_query else ""
            status = f"aceptada ({best_score:.2f}){suffix}"
        else:
            if not attempted and (args.solo_cache or (args.limite and queried >= args.limite)):
                reviews.append({
                    "id_local": location["id_local"], "consulta": primary_query,
                    "consultas_intentadas": [],
                    "motivo": "No procesada: se alcanzó el límite de consultas" if args.limite and queried >= args.limite else "Sin resultado disponible en caché",
                    "puntaje": 0, "resultado": "",
                    "direccion_original": location.get("direccion_original", ""),
                    "direccion_limpia": location.get("direccion_limpia", ""),
                    "nombre_local": location.get("nombre_local", ""),
                    "distrito": scoring_location.get("distrito", ""),
                    "provincia": scoring_location.get("provincia", ""),
                    "departamento": scoring_location.get("departamento", ""),
                    "lugar_conocido": (known_place or {}).get("nombre", ""),
                })
                continue
            reviews.append({
                "id_local": location["id_local"], "consulta": primary_query,
                "consultas_intentadas": attempted,
                "motivo": last_error or "Sin coincidencia suficientemente confiable",
                "puntaje": best_score,
                "resultado": best_candidate.get("display_name", "") if best_candidate else "",
                "direccion_original": location.get("direccion_original", ""),
                "direccion_limpia": location.get("direccion_limpia", ""),
                "nombre_local": location.get("nombre_local", ""),
                "distrito": scoring_location.get("distrito", ""),
                "provincia": scoring_location.get("provincia", ""),
                "departamento": scoring_location.get("departamento", ""),
                "lugar_conocido": (known_place or {}).get("nombre", ""),
            })
            status = f"revisar ({best_score:.2f})"
        print(f"[{position}/{total}] {status}: {primary_query[:65]}")

    resolution_count = cache.count_resolutions()
    cache.close()
    Path(args.salida).write_text(json.dumps(locations, ensure_ascii=False, indent=2), encoding="utf-8")
    Path(args.revision).write_text(json.dumps(reviews, ensure_ascii=False, indent=2), encoding="utf-8")
    features = []
    for location in locations:
        if location.get("latitud") is None or location.get("longitud") is None:
            continue
        features.append({
            "type": "Feature",
            "id": location["id_local"],
            "geometry": {"type": "Point", "coordinates": [location["longitud"], location["latitud"]]},
            "properties": {
                "id_local": location["id_local"], "id_promocion": location["id_promocion"],
                "nombre_local": location.get("nombre_local", ""), "direccion": location.get("direccion_limpia", ""),
                "distrito": location.get("distrito", ""), "precision": location.get("precision_ubicacion", ""),
                "confianza": location.get("confianza_ubicacion", ""), "fuente": location.get("fuente_coordenada", ""),
            },
        })
    Path(args.geojson).write_text(json.dumps({"type": "FeatureCollection", "features": features}, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Consultas nuevas: {queried}")
    if imported_count:
        print(f"Resultados anteriores incorporados al historial: {imported_count}")
    print(f"Coordenadas reutilizadas del historial: {persistent_count}")
    print(f"Resultados desde caché: {cached_count}")
    print(f"Coordenadas aceptadas: {accepted}")
    print(f"Direcciones únicas guardadas permanentemente: {resolution_count}")
    print(f"Casos para revisión: {len(reviews)}")
    print(f"Puntos disponibles para el mapa: {len(features)}")


if __name__ == "__main__":
    main()
