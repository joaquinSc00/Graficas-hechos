"""Prediseño automatizado.

Este módulo implementa el esqueleto funcional descrito en el diseño
interactivo.  Las funciones están divididas en etapas claramente
identificadas para permitir que diferentes miembros del equipo puedan
profundizar en cada paso sin interferir con los demás.  La meta a
futuro es procesar directamente un IDML para generar el plan de bloques
de notas junto con los archivos de texto que utilizará el JSX de
automatización.

La implementación real de parsing y geometría todavía no existe en este
archivo: cada función relevante expone su contrato y lanza
``NotImplementedError`` con un mensaje orientativo.  De esta forma se
garantiza que el script pueda ejecutarse en modo "stub" y que los
desarrolladores obtengan retroalimentación inmediata sobre qué etapas
falta completar.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

import csv
import json
import logging
import sys

try:  # pragma: no cover - dependencia opcional en tiempo de ejecución
    from docx import Document  # type: ignore
except Exception:  # pragma: no cover - la librería puede no estar instalada
    Document = None  # type: ignore

try:
    from generate_slot_report import POINT_TO_MM, analyze_document as analyze_idml_document
except Exception:  # pragma: no cover - import defensivo
    # Si la importación directa falla (por ejemplo al ejecutar el script de forma
    # independiente) reutilizamos la implementación probada del analizador IDML.
    import math
    import os
    import zipfile
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

    def analyze_idml_document(path: str):
        return analyze_document(path)


# ---------------------------------------------------------------------------
# Modelos de datos


@dataclass
class Config:
    """Parámetros capturados desde la consola."""

    idml_path: Path
    pages: List[int]
    cierre_root: Path
    output_dir: Path
    slot_selector: Dict[str, str] = field(default_factory=dict)


@dataclass
class DocumentInfo:
    """Información general del documento IDML."""

    page_size_mm: Tuple[float, float]
    bleed_mm: Tuple[float, float, float, float]


@dataclass
class PageGeometry:
    """Información geométrica deducida del IDML."""

    page_number: int
    name: str
    usable_rect_mm: Tuple[float, float, float, float]
    columns: int
    gutter_mm: float
    column_width_mm: float
    margins_mm: Dict[str, float]


@dataclass
class Block:
    """Bloque de 1 columna generado a partir de la retícula."""

    page: int
    column_index: int
    span: int
    x_mm: float
    y_mm: float
    w_mm: float
    h_mm: float


@dataclass
class Note:
    """Nota extraída desde DOCX."""

    note_id: int
    title: str
    body: str
    chars_title: int
    chars_body: int
    words: int
    source: Optional[Path] = None


@dataclass
class PipelineStats:
    """Estadísticas básicas para el log final."""

    pages_processed: int = 0
    total_blocks: int = 0
    notes_processed: int = 0
    warnings: List[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Entrada por consola


def prompt_inputs() -> Config:
    """Solicita los parámetros básicos para ejecutar el pipeline.

    Los valores ingresados se normalizan a rutas absolutas y se aplican
    los defaults solicitados en el diseño.
    """

    default_pages = "2,3,5,6,7,8,9,10,11,12,13,14,15,16,17,18"

    def _clean_path(raw: str) -> Path:
        """Normaliza rutas ingresadas manualmente.

        Es habitual que, al copiar y pegar desde el explorador de
        archivos, Windows incluya comillas dobles alrededor de la ruta.
        Esto provoca que ``Path`` busque un archivo con comillas
        literales en su nombre.  Para evitar el fallo se eliminan las
        comillas de apertura y cierre (simples o dobles) antes de
        convertir a :class:`Path`.
        """

        cleaned = raw.strip().strip('"\'')
        return Path(cleaned).expanduser().resolve()

    idml_raw = input("Ruta del IDML: ")
    pages_raw = input(
        f"Páginas a procesar (ENTER para {default_pages}): "
    ).strip()
    cierre_raw = input("Carpeta del cierre: ")
    output_raw = input("Carpeta de salida: ")

    if not pages_raw:
        pages_raw = default_pages

    pages = [int(p.strip()) for p in pages_raw.split(",") if p.strip()]

    cfg = Config(
        idml_path=_clean_path(idml_raw),
        pages=pages,
        cierre_root=_clean_path(cierre_raw),
        output_dir=_clean_path(output_raw),
        slot_selector={"layer": "ESPACIO_NOTAS"},
    )

    return cfg


# ---------------------------------------------------------------------------
# Orquestación


def run_pipeline(cfg: Config) -> None:
    """Ejecuta el flujo completo definido en el diseño."""

    logging.info("Iniciando pipeline con IDML %s", cfg.idml_path)

    document_info, pages, items = parse_idml_document(cfg.idml_path)
    usable_by_page = extract_usable_rect_by_page(pages, items, cfg.slot_selector)
    grid_by_page = extract_grid_by_page(pages)
    page_geometries = merge_page_geometry(pages, usable_by_page, grid_by_page)

    target_pages = filter_target_pages(page_geometries, cfg.pages)

    blocks: List[Block] = []
    stats = PipelineStats()

    for page_geom in target_pages:
        issues = validate_page_geometry(page_geom)
        if issues:
            stats.warnings.extend(issues)

        page_blocks = build_onecol_blocks(page_geom)
        blocks.extend(page_blocks)

        stats.pages_processed += 1
        stats.total_blocks += len(page_blocks)

    page_map = scan_closure_folder(cfg.cierre_root)

    for page_geom in target_pages:
        notes = extract_notes_for_page(page_map.get(page_geom.name))
        write_out_txt(page_geom.name, notes, cfg.output_dir)
        stats.notes_processed += len(notes)

    save_plan_blocks_json(blocks, cfg.output_dir)
    save_plan_blocks_csv(blocks, cfg.output_dir)
    log_global_summary(stats)


# ---------------------------------------------------------------------------
# 1) IDML → Auditor interno


def parse_idml_document(idml_path: Path):
    """Extrae información geométrica del IDML.

    Esta función debe replicar la lógica probada en ``generate_slot_report.py``
    para garantizar que las coordenadas resultantes estén expresadas en el
    sistema de referencia de página.  Retorna la metainformación general del
    documento, la lista de páginas con su geometría básica y la colección de
    items (rectángulos, polígonos, textframes) encontrados.
    """

    logging.debug("Analizando documento IDML %s", idml_path)

    document_info_raw, pages, items = analyze_idml_document(str(idml_path))

    if pages:
        first_bounds = pages[0].bounds
        page_size_mm = (
            first_bounds.width * POINT_TO_MM,
            first_bounds.height * POINT_TO_MM,
        )
    else:
        page_size_mm = (0.0, 0.0)

    doc_info = DocumentInfo(page_size_mm=page_size_mm, bleed_mm=(0.0, 0.0, 0.0, 0.0))

    logging.info(
        "Documento '%s' — páginas: %s — spreads: %s",
        document_info_raw.get("name"),
        document_info_raw.get("pages"),
        document_info_raw.get("spreads"),
    )

    return doc_info, pages, items


def extract_usable_rect_by_page(pages, items, selector: Dict[str, str]):
    """Selecciona el rectángulo utilizable por página."""

    selector_normalized = {
        key: (value.strip().lower() if isinstance(value, str) else value)
        for key, value in selector.items()
        if value is not None
    }

    candidates: Dict[str, List[Tuple[bool, float, Tuple[float, float, float, float]]]] = {}

    for item in items:
        if item.page_id is None:
            continue

        if item.item_type not in {"Rectangle", "Polygon", "TextFrame"}:
            continue

        match = True
        for key, expected in selector_normalized.items():
            actual = getattr(item, key, None)
            if actual is None:
                actual = item.as_dict().get(key)  # type: ignore[arg-type]
            if isinstance(expected, str):
                actual_str = str(actual or "").strip().lower()
                if actual_str != expected:
                    match = False
                    break
            else:
                if actual != expected:
                    match = False
                    break

        bounds = item.bounds_page or item.bounds_spread
        if not bounds:
            continue

        rect = (
            bounds.left * POINT_TO_MM,
            bounds.top * POINT_TO_MM,
            bounds.width * POINT_TO_MM,
            bounds.height * POINT_TO_MM,
        )
        area = rect[2] * rect[3]

        match_flag = bool(selector_normalized == {} or match)
        hard_match = bool(match and selector_normalized)

        if not match_flag and selector_normalized:
            # Guardamos como candidato de fallback si no encontramos coincidencias estrictas.
            candidates.setdefault(item.page_id, []).append((False, area, rect))
            continue

        candidates.setdefault(item.page_id, []).append((hard_match or not selector_normalized, area, rect))

    usable: Dict[str, Tuple[float, float, float, float]] = {}
    for page_id, rects in candidates.items():
        if not rects:
            continue
        rects.sort(key=lambda entry: (entry[0], entry[1]), reverse=True)
        usable[page_id] = rects[0][2]

    return usable


def extract_grid_by_page(pages):
    """Determina la retícula de columnas por página."""

    grid: Dict[str, Dict[str, float]] = {}
    for page in pages:
        width_mm = page.bounds.width * POINT_TO_MM
        columns = 5 if width_mm >= 250.0 else 4
        gutter_mm = 4.0
        grid[page.page_id] = {
            "columns": float(columns),
            "gutter_mm": gutter_mm,
        }
    return grid


def merge_page_geometry(pages, usable_rects, grids) -> List[PageGeometry]:
    """Combina la información de rectángulos y retícula en objetos completos."""

    geometries: List[PageGeometry] = []

    for page in pages:
        usable = usable_rects.get(page.page_id)
        if usable is None:
            usable = (0.0, 0.0, 0.0, 0.0)

        left_mm, top_mm, width_mm, height_mm = usable

        page_width_mm = page.bounds.width * POINT_TO_MM
        page_height_mm = page.bounds.height * POINT_TO_MM

        margins = {
            "left": max(left_mm, 0.0),
            "top": max(top_mm, 0.0),
            "right": max(page_width_mm - (left_mm + width_mm), 0.0),
            "bottom": max(page_height_mm - (top_mm + height_mm), 0.0),
        }

        grid_info = grids.get(page.page_id, {})
        columns = int(grid_info.get("columns", 5))
        gutter_mm = float(grid_info.get("gutter_mm", 4.0))

        if columns > 0:
            inner_width = max(width_mm - gutter_mm * (columns - 1), 0.0)
            column_width_mm = inner_width / columns if columns else 0.0
        else:
            column_width_mm = 0.0

        geometries.append(
            PageGeometry(
                page_number=page.index,
                name=page.name,
                usable_rect_mm=(left_mm, top_mm, width_mm, height_mm),
                columns=columns,
                gutter_mm=gutter_mm,
                column_width_mm=column_width_mm,
                margins_mm=margins,
            )
        )

    return geometries


# ---------------------------------------------------------------------------
# 2) Filtro y verificación


def filter_target_pages(pages: Sequence[PageGeometry], wanted: Iterable[int]) -> List[PageGeometry]:
    """Filtra y ordena las páginas solicitadas."""

    wanted_set = set(wanted)
    selected = [p for p in pages if p.page_number in wanted_set]
    return sorted(selected, key=lambda p: p.page_number)


def validate_page_geometry(page_geom: PageGeometry) -> List[str]:
    """Verifica que la geometría cumpla con las reglas básicas."""

    issues: List[str] = []

    if not page_geom.usable_rect_mm:
        issues.append(f"Página {page_geom.page_number}: no se encontró rectángulo utilizable")

    if page_geom.columns not in {4, 5}:
        issues.append(
            f"Página {page_geom.page_number}: columnas esperadas 4/5, se obtuvo {page_geom.columns}"
        )

    if page_geom.gutter_mm <= 0:
        issues.append(f"Página {page_geom.page_number}: gutter <= 0")

    if page_geom.column_width_mm <= 0:
        issues.append(f"Página {page_geom.page_number}: ancho de columna <= 0")

    return issues


# ---------------------------------------------------------------------------
# 3) Generación de bloques


def build_onecol_blocks(page_geom: PageGeometry) -> List[Block]:
    """Genera bloques de span=1 dentro del rectángulo utilizable."""

    blocks: List[Block] = []

    left_mm, top_mm, width_mm, height_mm = page_geom.usable_rect_mm
    if width_mm <= 0 or height_mm <= 0 or page_geom.columns <= 0:
        return blocks

    column_width = page_geom.column_width_mm
    gutter = page_geom.gutter_mm

    for idx in range(page_geom.columns):
        x_mm = left_mm + idx * (column_width + gutter)
        block = Block(
            page=page_geom.page_number,
            column_index=idx + 1,
            span=1,
            x_mm=x_mm,
            y_mm=top_mm,
            w_mm=column_width,
            h_mm=height_mm,
        )
        blocks.append(block)

    return blocks


# ---------------------------------------------------------------------------
# 4) Parsing de notas


def scan_closure_folder(root: Path) -> Dict[str, Path]:
    """Detecta las carpetas asociadas a cada página en el cierre."""

    if root is None:
        raise ValueError("scan_closure_folder requiere una ruta raíz válida")

    mapping: Dict[str, Path] = {}

    if not root.exists():
        logging.warning("La carpeta de cierre %s no existe", root)
        return mapping

    for entry in root.iterdir():
        if not entry.is_dir():
            continue
        name = entry.name.strip()
        if not name:
            continue
        mapping.setdefault(name, entry)

        digits = "".join(ch for ch in name if ch.isdigit())
        if digits:
            mapping.setdefault(digits, entry)
            mapping.setdefault(f"Página {int(digits)}", entry)

    return mapping


def extract_notes_for_page(page_dir: Optional[Path]) -> List[Note]:
    """Procesa los DOCX de la página indicada."""

    if page_dir is None:
        logging.warning("No se encontró carpeta de cierre para la página")
        return []

    notes: List[Note] = []

    if not page_dir.exists():
        logging.warning("La carpeta %s no existe", page_dir)
        return notes

    docx_files = sorted(page_dir.glob("*.docx"))
    txt_files = sorted(page_dir.glob("*.txt")) if not docx_files else []

    counter = 1

    if docx_files and Document is None:
        logging.warning("python-docx no está disponible; se omiten los DOCX en %s", page_dir)
        docx_files = []

    for path in docx_files:
        try:
            document = Document(path)
        except Exception as exc:  # pragma: no cover - lectura de terceros
            logging.warning("No se pudo abrir %s: %s", path, exc)
            continue

        paragraphs = [para.text.strip() for para in document.paragraphs if para.text.strip()]
        if not paragraphs:
            continue

        title = paragraphs[0]
        body = "\n".join(paragraphs[1:]) if len(paragraphs) > 1 else ""
        body_words = body.split()
        title_words = title.split()

        notes.append(
            Note(
                note_id=counter,
                title=title,
                body=body,
                chars_title=len(title),
                chars_body=len(body),
                words=len(title_words) + len(body_words),
                source=path,
            )
        )
        counter += 1

    if not notes:
        for path in txt_files:
            try:
                content = path.read_text(encoding="utf-8").strip()
            except Exception as exc:  # pragma: no cover - lectura IO
                logging.warning("No se pudo leer %s: %s", path, exc)
                continue
            if not content:
                continue
            lines = [line.strip() for line in content.splitlines() if line.strip()]
            if not lines:
                continue
            title = lines[0]
            body = "\n".join(lines[1:]) if len(lines) > 1 else ""
            notes.append(
                Note(
                    note_id=counter,
                    title=title,
                    body=body,
                    chars_title=len(title),
                    chars_body=len(body),
                    words=len((title + " " + body).split()),
                    source=path,
                )
            )
            counter += 1

    return notes


def write_out_txt(page_name: str, notes: Sequence[Note], out_dir: Path) -> None:
    """Escribe los archivos .txt y meta.json solicitados."""

    page_dir = out_dir / "out_txt" / page_name
    page_dir.mkdir(parents=True, exist_ok=True)

    meta = []

    for note in notes:
        title_path = page_dir / f"{note.note_id:02d}_title.txt"
        body_path = page_dir / f"{note.note_id:02d}_body.txt"

        title_path.write_text(note.title, encoding="utf-8")
        body_path.write_text(note.body, encoding="utf-8")

        meta.append(
            {
                "note_id": note.note_id,
                "chars_title": note.chars_title,
                "chars_body": note.chars_body,
                "words": note.words,
                "title_path": title_path.name,
                "body_path": body_path.name,
            }
        )

    meta_path = page_dir / "meta.json"
    meta_path.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")


# ---------------------------------------------------------------------------
# 5) Serialización


def save_plan_blocks_json(blocks: Sequence[Block], out_dir: Path) -> None:
    """Guarda el plan de bloques como JSON."""

    data = [block.__dict__ for block in blocks]
    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / "plan_bloques.json").write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def save_plan_blocks_csv(blocks: Sequence[Block], out_dir: Path) -> None:
    """Guarda el plan de bloques como CSV."""

    out_dir.mkdir(parents=True, exist_ok=True)
    csv_path = out_dir / "plan_bloques.csv"
    with csv_path.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh, delimiter=";")
        writer.writerow(["page", "i", "span", "x_mm", "y_mm", "w_mm", "h_mm"])
        for block in blocks:
            writer.writerow(
                [
                    block.page,
                    block.column_index,
                    block.span,
                    f"{block.x_mm:.2f}",
                    f"{block.y_mm:.2f}",
                    f"{block.w_mm:.2f}",
                    f"{block.h_mm:.2f}",
                ]
            )


# ---------------------------------------------------------------------------
# 6) Logging


def log_page_summary(page_name: str, n_blocks: int, issues: Sequence[str]) -> None:
    """Escribe un resumen por página en el log principal."""

    if issues:
        for issue in issues:
            logging.warning(issue)
    logging.info("Página %s: %s bloques generados", page_name, n_blocks)


def log_global_summary(stats: PipelineStats) -> None:
    """Registra las estadísticas generales y configura el archivo de log."""

    log_dir = Path("logs")
    log_dir.mkdir(exist_ok=True)
    log_path = log_dir / "prediseno.log"

    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    formatter = logging.Formatter("%(levelname)s\t%(message)s")
    file_handler.setFormatter(formatter)
    root_logger = logging.getLogger()
    root_logger.addHandler(file_handler)

    logging.info("Páginas procesadas: %s", stats.pages_processed)
    logging.info("Bloques generados: %s", stats.total_blocks)
    logging.info("Notas procesadas: %s", stats.notes_processed)

    for warning in stats.warnings:
        logging.warning(warning)


# ---------------------------------------------------------------------------
# Punto de entrada CLI


def main(argv: Optional[Sequence[str]] = None) -> int:
    """Punto de entrada principal para la línea de comandos."""

    logging.basicConfig(level=logging.INFO, format="%(levelname)s\t%(message)s")

    cfg = prompt_inputs()
    try:
        run_pipeline(cfg)
    except NotImplementedError as exc:
        logging.error("Funcionalidad no implementada: %s", exc)
        return 1
    except Exception as exc:  # pragma: no cover - guardia general
        logging.exception("Fallo inesperado del pipeline: %s", exc)
        return 1

    return 0


if __name__ == "__main__":  # pragma: no cover - ejecución directa
    raise SystemExit(main(sys.argv[1:]))

