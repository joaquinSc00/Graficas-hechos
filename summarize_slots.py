#!/usr/bin/env python3
"""Summaries slot-sized rectangles from a slot report JSON file."""

from __future__ import annotations

import argparse
import csv
import json
from pathlib import Path
from typing import Dict, Iterator, List, Optional, Sequence, Tuple


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Resume por página los slots detectados en un reporte generado por"
            " generate_slot_report.py."
        )
    )
    parser.add_argument(
        "input",
        help="Ruta al archivo *_slots_report.json generado por generate_slot_report.py",
    )
    parser.add_argument(
        "--csv",
        dest="csv_path",
        default=None,
        help="Si se especifica, guarda el resumen por página en un CSV",
    )
    parser.add_argument(
        "--only-slots",
        action="store_true",
        help=(
            "Filtra únicamente los elementos que tengan etiqueta o estilo compatible"
            " con slots (root, SLOT_*)"
        ),
    )
    parser.add_argument(
        "--min-area-cm2",
        type=float,
        default=0.0,
        help="Descarta slots con área menor a este umbral (en cm²)",
    )
    return parser.parse_args(argv)


def load_payload(path: str) -> Dict[str, object]:
    with open(path, "r", encoding="utf-8") as fh:
        return json.load(fh)


def iter_page_items(payload: Dict[str, object]) -> Iterator[Tuple[Dict[str, object], Dict[str, object]]]:
    pages = payload.get("pages", [])
    if not isinstance(pages, list):
        return iter(())
    for page in pages:
        items = page.get("items", [])
        if not isinstance(items, list):
            continue
        for item in items:
            yield page, item


def looks_like_slot(item: Dict[str, object]) -> bool:
    label = str(item.get("label", "") or "").strip().lower()
    if label in {"slot", "root"}:
        return True
    if label.startswith("slot") or "slot" in label.split():
        return True

    style = str(item.get("object_style", "") or "").strip().lower()
    if style:
        style_name = style.split("/")[-1]
        if style_name.startswith("slot") or style_name in {"root", "slot"}:
            return True
    return False


def bounds_from_item(item: Dict[str, object]) -> Optional[Dict[str, float]]:
    bounds = item.get("bounds_page") or item.get("bounds_spread")
    if not isinstance(bounds, dict):
        return None
    try:
        width_mm = float(bounds.get("width_mm", 0.0))
        height_mm = float(bounds.get("height_mm", 0.0))
        left_mm = float(bounds.get("left_mm", 0.0))
        top_mm = float(bounds.get("top_mm", 0.0))
    except (TypeError, ValueError):
        return None
    return {
        "width_mm": width_mm,
        "height_mm": height_mm,
        "left_mm": left_mm,
        "top_mm": top_mm,
    }


def area_cm2(bounds: Dict[str, float]) -> float:
    return (bounds["width_mm"] * bounds["height_mm"]) / 100.0


def summarize(payload: Dict[str, object], only_slots: bool, min_area_cm2: float) -> List[Dict[str, object]]:
    summary: Dict[str, Dict[str, object]] = {}
    for page, item in iter_page_items(payload):
        page_key = str(page.get("page_id", ""))
        summary.setdefault(
            page_key,
            {
                "page_index": page.get("page_index"),
                "page_id": page.get("page_id"),
                "page_name": page.get("page_name"),
                "slots": [],
            },
        )

        if only_slots and not looks_like_slot(item):
            continue

        bounds = bounds_from_item(item)
        if not bounds:
            continue
        slot_area = area_cm2(bounds)
        if slot_area < min_area_cm2:
            continue

        summary[page_key]["slots"].append(
            {
                "item_id": item.get("item_id"),
                "label": item.get("label"),
                "object_style": item.get("object_style"),
                "width_mm": bounds["width_mm"],
                "height_mm": bounds["height_mm"],
                "area_cm2": slot_area,
            }
        )

    ordered = sorted(summary.values(), key=lambda entry: (entry["page_index"] or 0))
    for entry in ordered:
        slots = entry["slots"]
        total_area = sum(slot["area_cm2"] for slot in slots)
        entry["slot_count"] = len(slots)
        entry["total_area_cm2"] = total_area
        entry["average_area_cm2"] = total_area / len(slots) if slots else 0.0
    return ordered


def print_summary(payload: Dict[str, object], summary: Sequence[Dict[str, object]]) -> None:
    document = payload.get("document", {})
    doc_name = document.get("name", "(sin nombre)")
    total_pages = document.get("pages", len(summary))
    print(f"Documento: {doc_name}")
    print(f"Páginas reportadas: {total_pages}")
    print()
    for entry in summary:
        page_index = entry.get("page_index")
        page_name = entry.get("page_name")
        label = f" (pág. {page_name})" if page_name not in (None, "", page_index) else ""
        slot_count = entry.get("slot_count", 0)
        total_area = entry.get("total_area_cm2", 0.0)
        average_area = entry.get("average_area_cm2", 0.0)
        print(
            f"• Página {page_index}{label}: {slot_count} slots"
            f" — área total {total_area:.1f} cm² (promedio {average_area:.1f} cm²)"
        )
    print()


def write_csv(csv_path: str, summary: Sequence[Dict[str, object]]) -> None:
    fieldnames = [
        "page_index",
        "page_id",
        "page_name",
        "slot_count",
        "total_area_cm2",
        "average_area_cm2",
    ]
    with open(csv_path, "w", encoding="utf-8", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()
        for entry in summary:
            writer.writerow(
                {
                    "page_index": entry.get("page_index"),
                    "page_id": entry.get("page_id"),
                    "page_name": entry.get("page_name"),
                    "slot_count": entry.get("slot_count"),
                    "total_area_cm2": f"{entry.get('total_area_cm2', 0.0):.2f}",
                    "average_area_cm2": f"{entry.get('average_area_cm2', 0.0):.2f}",
                }
            )


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)
    payload = load_payload(args.input)
    summary = summarize(payload, args.only_slots, args.min_area_cm2)
    print_summary(payload, summary)
    if args.csv_path:
        write_csv(args.csv_path, summary)
        print(f"Resumen guardado en: {Path(args.csv_path).resolve()}")
    return 0


if __name__ == "__main__":  # pragma: no cover - script entry point
    raise SystemExit(main())
