from __future__ import annotations

import argparse
import gzip
import hashlib
import json
from datetime import datetime, timezone
from pathlib import Path


PROMOTION_FIELDS = (
    "bank", "type", "name", "merchant", "detail", "benefit", "discount",
    "category", "subcategory", "sourceCategory", "startDate", "endDate",
    "status", "paymentMethod", "restrictions", "terms", "link", "image",
    "sourceFile",
)


def compact_catalog(rows: list[dict]) -> list[dict]:
    promotions: dict[str, dict] = {}
    location_keys: dict[str, set[tuple]] = {}
    for row in rows:
        promotion_id = str(row.get("promotionId") or row.get("id") or "").strip()
        if not promotion_id:
            continue
        if promotion_id not in promotions:
            item = {"id": promotion_id}
            item.update({field: row.get(field, "") for field in PROMOTION_FIELDS})
            item["locations"] = []
            promotions[promotion_id] = item
            location_keys[promotion_id] = set()

        address = str(row.get("location") or "").strip()
        district = str(row.get("district") or "").strip()
        lat, lng = row.get("latitude"), row.get("longitude")
        has_location = bool(row.get("locationId") or address or district or lat is not None or lng is not None)
        if not has_location:
            continue
        key = (address, district, lat, lng)
        if key in location_keys[promotion_id]:
            continue
        location_keys[promotion_id].add(key)
        promotions[promotion_id]["locations"].append({
            "id": row.get("locationId") or f"{promotion_id}-{len(location_keys[promotion_id])}",
            "address": address,
            "district": district,
            "province": row.get("province") or "",
            "department": row.get("department") or "",
            "lat": lat,
            "lng": lng,
            "precision": row.get("coordinatePrecision") or "none",
            "mapStatus": row.get("mapStatus") or "",
        })
    return sorted(promotions.values(), key=lambda p: (str(p.get("bank")), str(p.get("merchant")), str(p.get("name"))))


def main() -> int:
    parser = argparse.ArgumentParser(description="Normaliza y divide el catálogo comprimido para GitHub Pages.")
    parser.add_argument("--entrada", required=True)
    parser.add_argument("--salida", default="site/data")
    parser.add_argument("--tamano-parte", type=int, default=96_000)
    args = parser.parse_args()

    rows = json.loads(Path(args.entrada).read_text(encoding="utf-8"))
    catalog = compact_catalog(rows)
    raw = json.dumps(catalog, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    compressed = gzip.compress(raw, compresslevel=9, mtime=0)
    output = Path(args.salida)
    parts_dir = output / "catalog.parts"
    parts_dir.mkdir(parents=True, exist_ok=True)
    for previous in parts_dir.glob("part-*.bin"):
        previous.unlink()

    names = []
    for index, start in enumerate(range(0, len(compressed), args.tamano_parte), 1):
        name = f"part-{index:03d}.bin"
        (parts_dir / name).write_bytes(compressed[start:start + args.tamano_parte])
        names.append(name)

    mapped = sum(1 for p in catalog for loc in p["locations"] if loc["lat"] is not None and loc["lng"] is not None)
    manifest = {
        "version": 1,
        "generatedAt": datetime.now(timezone.utc).isoformat(),
        "promotions": len(catalog),
        "mappedLocations": mapped,
        "uncompressedBytes": len(raw),
        "compressedBytes": len(compressed),
        "sha256": hashlib.sha256(compressed).hexdigest(),
        "parts": names,
    }
    (output / "catalog-manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(manifest, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
