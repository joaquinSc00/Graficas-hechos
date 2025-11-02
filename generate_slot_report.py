#!/usr/bin/env python3
"""Analiza un documento IDML y genera un reporte de rectángulos por página."""


from __future__ import annotations

import argparse
import csv
import json
import math
import os
import sys
import zipfile
from dataclasses import dataclass, field
from typing import Dict, Iterable, List, Optional, Sequence, Tuple
import xml.etree.ElementTree as ET

IDPKG_NS = "{http://ns.adobe.com/AdobeInDesign/idml/1.0/packaging}"
POINT_TO_MM = 25.4 / 72.0


@dataclass
class Transform:
    a: float = 1.0
    b: float = 0.0
    c: float = 0.0
    d: float = 1.0
    tx: float = 0.0
    ty: float = 0.0

    @classmethod
    def from_string(cls, value: Optional[str]) -> "Transform":
        if not value:
            return cls()
        parts: List[float] = []
        for chunk in value.replace(",", " ").split():
            if not chunk:
                continue
            try:
                parts.append(float(chunk))
            except ValueError:
                continue
        while len(parts) < 6:
            parts.append(0.0)
        return cls(*parts[:6])

    def apply(self, x: float, y: float) -> Tuple[float, float]:
        return (
            self.a * x + self.c * y + self.tx,
            self.b * x + self.d * y + self.ty,
        )

    def combine(self, other: "Transform") -> "Transform":
        # self ∘ other (apply "other" first, then "self")
        return Transform(
            a=self.a * other.a + self.c * other.b,
            b=self.b * other.a + self.d * other.b,
            c=self.a * other.c + self.c * other.d,
            d=self.b * other.c + self.d * other.d,
            tx=self.a * other.tx + self.c * other.ty + self.tx,
            ty=self.b * other.tx + self.d * other.ty + self.ty,
        )

    def inverse(self) -> Optional["Transform"]:
        det = self.a * self.d - self.b * self.c
        if abs(det) < 1e-9:
            return None
        inv_a = self.d / det
        inv_b = -self.b / det
        inv_c = -self.c / det
        inv_d = self.a / det
        inv_tx = -(inv_a * self.tx + inv_c * self.ty)
        inv_ty = -(inv_b * self.tx + inv_d * self.ty)
        return Transform(inv_a, inv_b, inv_c, inv_d, inv_tx, inv_ty)

    def rotation_deg(self) -> float:
        return math.degrees(math.atan2(self.b, self.a))


@dataclass
class Bounds:
    top: float
    left: float
    bottom: float
    right: float

    @classmethod
    def from_string(cls, value: Optional[str]) -> Optional["Bounds"]:
        if not value:
            return None
        numbers: List[float] = []
        for chunk in value.replace(",", " ").split():
            if not chunk:
                continue
            try:
                numbers.append(float(chunk))
            except ValueError:
                continue
        if len(numbers) != 4:
            return None
        return cls(top=numbers[0], left=numbers[1], bottom=numbers[2], right=numbers[3])

    @property
    def width(self) -> float:
        return self.right - self.left

    @property
    def height(self) -> float:
        return self.bottom - self.top


@dataclass
class PageInfo:
    index: int
    spread_index: int
    spread_id: str
    page_id: str
    name: str
    transform: Transform
    bounds: Bounds
    inverse_transform: Optional[Transform] = field(init=False)

    def __post_init__(self) -> None:
        self.inverse_transform = self.transform.inverse()

    def contains_spread_point(self, x: float, y: float, tolerance: float = 0.5) -> bool:
        if not self.inverse_transform or not self.bounds:
            return False
        local_x, local_y = self.inverse_transform.apply(x, y)
        return (
            self.bounds.left - tolerance <= local_x <= self.bounds.right + tolerance
            and self.bounds.top - tolerance <= local_y <= self.bounds.bottom + tolerance
        )

    def to_page_coords(self, points: Sequence[Tuple[float, float]]) -> List[Tuple[float, float]]:
        if not self.inverse_transform:
            return []
        return [self.inverse_transform.apply(x, y) for (x, y) in points]


@dataclass
class ItemReport:
    item_id: str
    item_type: str
    label: str
    object_style: str
    layer: str
    spread_id: str
    spread_index: int
    page_id: Optional[str]
    page_index: Optional[int]
    page_name: Optional[str]
    bounds_spread: Bounds
    bounds_page: Optional[Bounds]
    rotation_deg: float
    path_points_spread: List[Tuple[float, float]]
    path_points_page: Optional[List[Tuple[float, float]]]

    def as_dict(self) -> Dict[str, object]:
        def bounds_to_dict(bounds: Optional[Bounds]) -> Optional[Dict[str, object]]:
            if not bounds:
                return None
            return {
                "top_pt": bounds.top,
                "left_pt": bounds.left,
                "bottom_pt": bounds.bottom,
                "right_pt": bounds.right,
                "width_pt": bounds.width,
                "height_pt": bounds.height,
                "top_mm": bounds.top * POINT_TO_MM,
                "left_mm": bounds.left * POINT_TO_MM,
                "bottom_mm": bounds.bottom * POINT_TO_MM,
                "right_mm": bounds.right * POINT_TO_MM,
                "width_mm": bounds.width * POINT_TO_MM,
                "height_mm": bounds.height * POINT_TO_MM,
                "top_cm": bounds.top * POINT_TO_MM / 10.0,
                "left_cm": bounds.left * POINT_TO_MM / 10.0,
                "bottom_cm": bounds.bottom * POINT_TO_MM / 10.0,
                "right_cm": bounds.right * POINT_TO_MM / 10.0,
                "width_cm": bounds.width * POINT_TO_MM / 10.0,
                "height_cm": bounds.height * POINT_TO_MM / 10.0,
            }

        return {
            "item_id": self.item_id,
            "item_type": self.item_type,
            "label": self.label,
            "object_style": self.object_style,
            "layer": self.layer,
            "spread_id": self.spread_id,
            "spread_index": self.spread_index,
            "page_id": self.page_id,
            "page_index": self.page_index,
            "page_name": self.page_name,
            "rotation_deg": self.rotation_deg,
            "bounds_spread": bounds_to_dict(self.bounds_spread),
            "bounds_page": bounds_to_dict(self.bounds_page),
            "path_points_spread": [
                {"x": x, "y": y} for (x, y) in self.path_points_spread
            ],
            "path_points_page": None
            if not self.path_points_page
            else [{"x": x, "y": y} for (x, y) in self.path_points_page],
        }


TARGET_TAGS = {"Rectangle", "Polygon", "TextFrame"}
CONTAINER_TAGS = {
    "Spread",
    "Page",
    "Group",
    "Button",
    "State",
    "MultiStateObject",
    "EPSText",
    "EPSGraphic",
    "Rectangle",
    "Polygon",
    "TextFrame",
}


def strip_namespace(tag: str) -> str:
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


def extract_path_points(element: ET.Element) -> List[Tuple[float, float]]:
    points: List[Tuple[float, float]] = []
    for path_point in element.findall(".//PathPointType"):
        anchor = path_point.attrib.get("Anchor")
        if not anchor:
            continue
        coords: List[float] = []
        for chunk in anchor.replace(",", " ").split():
            if not chunk:
                continue
            try:
                coords.append(float(chunk))
            except ValueError:
                continue
        if len(coords) >= 2:
            points.append((coords[0], coords[1]))
    return points


def extract_label(element: ET.Element) -> str:
    properties = element.find("Properties")
    if properties is not None:
        label_elem = properties.find("Label")
        if label_elem is not None:
            content = label_elem.findtext("Content")
            if content:
                return content
            key_values = label_elem.findall("KeyValuePair")
            for kv in key_values:
                key = kv.attrib.get("Key")
                value = kv.attrib.get("Value")
                if key == "__default__" and value:
                    return value
                if value:
                    return value
    return element.attrib.get("MarkupTag", "")


def make_bounds_from_points(points: Sequence[Tuple[float, float]]) -> Optional[Bounds]:
    if not points:
        return None
    xs = [p[0] for p in points]
    ys = [p[1] for p in points]
    return Bounds(top=min(ys), left=min(xs), bottom=max(ys), right=max(xs))


def iter_items_with_transforms(root: ET.Element, base_transform: Transform) -> Iterable[Tuple[ET.Element, Transform]]:
    stack: List[Tuple[ET.Element, Transform]] = [(root, base_transform)]
    while stack:
        element, transform = stack.pop()
        tag = strip_namespace(element.tag)
        if tag in TARGET_TAGS and element is not root:
            yield element, transform
        for child in list(element):
            child_tag = strip_namespace(child.tag)
            if child_tag == "Properties":
                continue
            child_transform_attr = child.attrib.get("ItemTransform")
            if child_transform_attr:
                child_transform = transform.combine(Transform.from_string(child_transform_attr))
            else:
                child_transform = transform
            if (
                child_tag in CONTAINER_TAGS
                or child_transform_attr
                or child_tag in TARGET_TAGS
            ):
                stack.append((child, child_transform))


def parse_spread(zip_file: zipfile.ZipFile, spread_src: str, spread_index: int, page_counter: int) -> Tuple[List[PageInfo], List[ItemReport]]:
    with zip_file.open(spread_src) as fh:
        tree = ET.parse(fh)
    root = tree.getroot()
    spread_elem = root.find("Spread")
    if spread_elem is None:
        return [], []

    spread_id = spread_elem.attrib.get("Self", spread_src)
    spread_transform = Transform.from_string(spread_elem.attrib.get("ItemTransform"))

    pages: List[PageInfo] = []
    page_reports: Dict[str, List[ItemReport]] = {}
    items: List[ItemReport] = []

    page_elements = spread_elem.findall("Page")
    for page_elem in page_elements:
        bounds = Bounds.from_string(page_elem.attrib.get("GeometricBounds"))
        transform = Transform.from_string(page_elem.attrib.get("ItemTransform"))
        page_counter += 1
        page_info = PageInfo(
            index=page_counter,
            spread_index=spread_index,
            spread_id=spread_id,
            page_id=page_elem.attrib.get("Self", f"page_{page_counter}"),
            name=page_elem.attrib.get("Name", str(page_counter)),
            transform=spread_transform.combine(transform),
            bounds=bounds if bounds else Bounds(0.0, 0.0, 0.0, 0.0),
        )
        pages.append(page_info)
        page_reports[page_info.page_id] = []

    items_raw: List[Tuple[ET.Element, Transform]] = list(
        iter_items_with_transforms(spread_elem, spread_transform)
    )

    for element, transform in items_raw:
        tag = strip_namespace(element.tag)
        if tag not in TARGET_TAGS:
            continue
        points_local = extract_path_points(element)
        if not points_local:
            continue
        points_spread = [transform.apply(x, y) for (x, y) in points_local]
        bounds_spread = make_bounds_from_points(points_spread)
        if not bounds_spread:
            continue
        center_x = (bounds_spread.left + bounds_spread.right) / 2.0
        center_y = (bounds_spread.top + bounds_spread.bottom) / 2.0

        assigned_page: Optional[PageInfo] = None
        for page in pages:
            if page.contains_spread_point(center_x, center_y):
                assigned_page = page
                break

        path_points_page: Optional[List[Tuple[float, float]]] = None
        bounds_page: Optional[Bounds] = None
        page_id: Optional[str] = None
        page_index: Optional[int] = None
        page_name: Optional[str] = None
        if assigned_page is not None:
            page_id = assigned_page.page_id
            page_index = assigned_page.index
            page_name = assigned_page.name
            path_points_page = assigned_page.to_page_coords(points_spread)
            bounds_page = make_bounds_from_points(path_points_page or [])

        report = ItemReport(
            item_id=element.attrib.get("Self", ""),
            item_type=tag,
            label=extract_label(element),
            object_style=element.attrib.get("AppliedObjectStyle", ""),
            layer=element.attrib.get("ItemLayer", ""),
            spread_id=spread_id,
            spread_index=spread_index,
            page_id=page_id,
            page_index=page_index,
            page_name=page_name,
            bounds_spread=bounds_spread,
            bounds_page=bounds_page,
            rotation_deg=transform.rotation_deg(),
            path_points_spread=points_spread,
            path_points_page=path_points_page,
        )
        items.append(report)
        if page_id is not None:
            page_reports[page_id].append(report)

    return pages, items


def analyze_document(path: str) -> Tuple[Dict[str, object], List[PageInfo], List[ItemReport]]:
    if not os.path.exists(path):
        raise FileNotFoundError(f"IDML document not found: {path}")

    with zipfile.ZipFile(path, "r") as zf:
        with zf.open("designmap.xml") as fh:
            designmap_root = ET.parse(fh).getroot()

        spread_elems = designmap_root.findall(f"{IDPKG_NS}Spread")
        spreads_order = [elem.attrib.get("src") for elem in spread_elems if elem.attrib.get("src")]

        document_info = {
            "name": designmap_root.attrib.get("Name", os.path.basename(path)),
            "spreads": len(spreads_order),
            "pages": 0,
        }

        all_pages: List[PageInfo] = []
        all_items: List[ItemReport] = []
        page_counter = 0

        for spread_index, spread_src in enumerate(spreads_order, start=1):
            if not spread_src:
                continue
            if spread_src not in zf.namelist():
                continue
            pages, items = parse_spread(zf, spread_src, spread_index, page_counter)
            if pages:
                page_counter = pages[-1].index
            all_pages.extend(pages)
            all_items.extend(items)

        document_info["pages"] = len(all_pages)

    return document_info, all_pages, all_items


def build_page_output(pages: Sequence[PageInfo], items: Sequence[ItemReport]) -> List[Dict[str, object]]:
    items_by_page: Dict[str, List[ItemReport]] = {}
    for item in items:
        if item.page_id is None:
            continue
        items_by_page.setdefault(item.page_id, []).append(item)

    output: List[Dict[str, object]] = []
    for page in pages:
        page_items = items_by_page.get(page.page_id, [])
        output.append(
            {
                "page_index": page.index,
                "page_id": page.page_id,
                "page_name": page.name,
                "spread_index": page.spread_index,
                "spread_id": page.spread_id,
                "bounds": {
                    "top_pt": page.bounds.top,
                    "left_pt": page.bounds.left,
                    "bottom_pt": page.bounds.bottom,
                    "right_pt": page.bounds.right,
                    "width_pt": page.bounds.width,
                    "height_pt": page.bounds.height,
                    "top_mm": page.bounds.top * POINT_TO_MM,
                    "left_mm": page.bounds.left * POINT_TO_MM,
                    "bottom_mm": page.bounds.bottom * POINT_TO_MM,
                    "right_mm": page.bounds.right * POINT_TO_MM,
                    "width_mm": page.bounds.width * POINT_TO_MM,
                    "height_mm": page.bounds.height * POINT_TO_MM,
                    "top_cm": page.bounds.top * POINT_TO_MM / 10.0,
                    "left_cm": page.bounds.left * POINT_TO_MM / 10.0,
                    "bottom_cm": page.bounds.bottom * POINT_TO_MM / 10.0,
                    "right_cm": page.bounds.right * POINT_TO_MM / 10.0,
                    "width_cm": page.bounds.width * POINT_TO_MM / 10.0,
                    "height_cm": page.bounds.height * POINT_TO_MM / 10.0,
                },
                "items": [item.as_dict() for item in page_items],
            }
        )
    return output


def write_json(output_path: str, payload: Dict[str, object]) -> None:
    with open(output_path, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, indent=2, ensure_ascii=False)
        fh.write("\n")


def write_csv(output_path: str, items: Sequence[ItemReport]) -> None:
    fieldnames = [
        "spread_index",
        "spread_id",
        "page_index",
        "page_id",
        "page_name",
        "item_type",
        "item_id",
        "label",
        "object_style",
        "layer",
        "rotation_deg",
        "left_pt",
        "top_pt",
        "right_pt",
        "bottom_pt",
        "width_pt",
        "height_pt",
    ]
    with open(output_path, "w", encoding="utf-8", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()
        for item in items:
            bounds = item.bounds_page or item.bounds_spread
            writer.writerow(
                {
                    "spread_index": item.spread_index,
                    "spread_id": item.spread_id,
                    "page_index": item.page_index,
                    "page_id": item.page_id,
                    "page_name": item.page_name,
                    "item_type": item.item_type,
                    "item_id": item.item_id,
                    "label": item.label,
                    "object_style": item.object_style,
                    "layer": item.layer,
                    "rotation_deg": round(item.rotation_deg, 6),
                    "left_pt": round(bounds.left, 6) if bounds else "",
                    "top_pt": round(bounds.top, 6) if bounds else "",
                    "right_pt": round(bounds.right, 6) if bounds else "",
                    "bottom_pt": round(bounds.bottom, 6) if bounds else "",
                    "width_pt": round(bounds.width, 6) if bounds else "",
                    "height_pt": round(bounds.height, 6) if bounds else "",
                }
            )


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Analiza un documento IDML y exporta un reporte con las posiciones de los rectángulos por página."
        )
    )
    parser.add_argument(
        "input",
        help="Ruta al archivo .idml de la maqueta",
    )
    parser.add_argument(
        "--out-json",
        dest="out_json",
        default=None,
        help="Ruta donde guardar el reporte JSON (por defecto <nombre>_slots_report.json)",
    )
    parser.add_argument(
        "--out-csv",
        dest="out_csv",
        default=None,
        help="Ruta donde guardar un resumen CSV (opcional)",
    )
    parser.add_argument(
        "--pretty",
        action="store_true",
        help="Imprime un resumen legible en consola además de los archivos",
    )
    return parser.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)
    try:
        document_info, pages, items = analyze_document(args.input)
    except Exception as exc:  # pylint: disable=broad-except
        print(f"Error analizando '{args.input}': {exc}", file=sys.stderr)
        return 1

    payload = {
        "document": document_info,
        "pages": build_page_output(pages, items),
        "items": [item.as_dict() for item in items],
    }

    base_name = os.path.splitext(os.path.basename(args.input))[0]
    out_json = args.out_json or f"{base_name}_slots_report.json"
    write_json(out_json, payload)

    if args.out_csv:
        write_csv(args.out_csv, items)

    if args.pretty:
        print(f"Documento: {document_info['name']}")
        print(f"Páginas: {document_info['pages']}  —  Spreads: {document_info['spreads']}")
        per_page_counts = [
            (page['page_index'], page['page_name'], len(page['items']))
            for page in payload['pages']
        ]
        for idx, name, count in per_page_counts:
            label = f" (pág. {name})" if name else ""
            print(f"  • Página {idx}{label}: {count} rectángulos/polígonos detectados")
        print(f"Reporte JSON: {out_json}")
        if args.out_csv:
            print(f"Resumen CSV: {args.out_csv}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
