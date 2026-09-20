from __future__ import annotations

import argparse
import gzip
import json
from pathlib import Path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Valida un catálogo antes de publicarlo.")
    parser.add_argument("--candidate", type=Path, default=Path("site/data"))
    parser.add_argument("--current", type=Path, default=Path("site/data"))
    parser.add_argument("--master", type=Path)
    parser.add_argument("--min-promotions", type=int, default=1500)
    parser.add_argument("--min-mapped", type=int, default=2000)
    parser.add_argument("--min-mapped-ratio", type=float, default=0.90)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    manifest_path = args.candidate / "catalog-manifest.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    if args.master and args.master.suffix == ".gz":
        with gzip.open(args.master, mode="rt", encoding="utf-8") as fh:
            master = json.load(fh)
    else:
        master = json.loads(args.master.read_text(encoding="utf-8")) if args.master else {}
    banks = {p.get("banco") for p in master.get("promociones", []) if p.get("banco")}

    mapped_floor = args.min_mapped
    current_manifest_path = args.current / "catalog-manifest.json"
    if current_manifest_path.exists() and current_manifest_path.resolve() != manifest_path.resolve():
        current = json.loads(current_manifest_path.read_text(encoding="utf-8"))
        current_mapped = int(current.get("mappedLocations", 0))
        mapped_floor = max(mapped_floor, int(current_mapped * args.min_mapped_ratio))
        print(
            f"Referencia vigente: {current_mapped:,} locales mapeados; "
            f"mínimo permitido: {mapped_floor:,}."
        )

    checks = {
        f"promociones >= {args.min_promotions}": manifest["promotions"] >= args.min_promotions,
        f"locales mapeados >= {mapped_floor}": manifest["mappedLocations"] >= mapped_floor,
        "bancos >= 5": len(banks) >= 5,
        "partes del catálogo": all(
            (args.candidate / "catalog.parts" / name).exists() for name in manifest["parts"]
        ),
    }
    for label, passed in checks.items():
        print(f"{'OK' if passed else 'ERROR'}: {label}")
    return 0 if all(checks.values()) else 2


if __name__ == "__main__":
    raise SystemExit(main())
