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
from typing import Any, Dict, Iterable, List, Mapping, Optional, Sequence, Tuple

import csv
import json
import logging
import sys

from solver import PlanSettings, solve_page_layout, summarize_outcome

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
class TextStyle:
    """Modelo tipográfico básico.

    El estilo captura la cantidad promedio de caracteres por línea, la densidad de
    líneas por milímetro y el costo vertical de los títulos por nivel.  Estos
    parámetros se utilizan para estimar capacidades y reservar espacio en los
    cálculos de variantes.
    """

    chars_per_line: int
    lines_per_mm: float
    title_heights_mm: Dict[int, float]

    def capacity_for_height(self, height_mm: float) -> int:
        """Calcula la capacidad de texto disponible para una altura dada."""

        effective_height = max(height_mm, 0.0)
        lines = effective_height * self.lines_per_mm
        return int(lines * self.chars_per_line)

    def title_cost(self, level: Optional[int], span: int = 1) -> float:
        """Devuelve la altura en mm reservada por un título."""

        if level is None:
            return 0.0
        base = self.title_heights_mm.get(level)
        if base is None:
            return 0.0
        return base * max(span, 1)

    @classmethod
    def defaults(cls) -> "TextStyle":
        """Genera un set de parámetros tipográficos por defecto."""

        return cls(chars_per_line=32, lines_per_mm=0.38, title_heights_mm={1: 16.0, 2: 12.0, 3: 8.0})

    @classmethod
    def from_config(cls, data: Mapping[str, Any]) -> "TextStyle":
        """Construye la instancia a partir de un diccionario de configuración."""

        if not data:
            return cls.defaults()

        defaults = cls.defaults()
        chars = int(
            data.get(
                "chars_per_line",
                data.get("chars_por_linea", data.get("caracteres_por_linea", defaults.chars_per_line)),
            )
        )
        lines_per_mm = float(
            data.get(
                "lines_per_mm",
                data.get("lineas_por_mm", data.get("líneas_por_mm", defaults.lines_per_mm)),
            )
        )

        raw_heights = data.get("title_heights_mm") or data.get("alturas_titulos_mm") or {}
        title_heights: Dict[int, float] = {}
        if isinstance(raw_heights, Mapping):
            for key, value in raw_heights.items():
                try:
                    level = int(key)
                    title_heights[level] = float(value)
                except (TypeError, ValueError):
                    continue

        if not title_heights:
            title_heights = dict(defaults.title_heights_mm)

        return cls(chars_per_line=chars, lines_per_mm=lines_per_mm, title_heights_mm=title_heights)


@dataclass
class ImagePreset:
    """Preset de imagen parametrizable por span."""

    name: str
    span: int
    height_mm: float

    def cost(self, span: Optional[int] = None) -> float:
        """Calcula el costo vertical en función del span solicitado."""

        if span is None or span == self.span:
            return self.height_mm

        base_span = max(self.span, 1)
        requested_span = max(span, 1)
        return self.height_mm * requested_span / base_span

    @classmethod
    def default_presets(cls) -> Dict[str, "ImagePreset"]:
        """Presets usados habitualmente en el flujo."""

        return {
            "horizontal": cls(name="horizontal", span=2, height_mm=43.0),
            "vertical": cls(name="vertical", span=1, height_mm=60.0),
        }

    @classmethod
    def from_config(cls, name: str, data: Mapping[str, Any]) -> "ImagePreset":
        """Construye un preset a partir de un diccionario de configuración."""

        defaults = cls.default_presets().get(name)
        span = int(data.get("span", data.get("columnas", defaults.span if defaults else 1)))

        if "height_mm" in data:
            height_mm = float(data["height_mm"])
        elif "height_cm" in data:
            height_mm = float(data["height_cm"]) * 10.0
        elif "alto_cm" in data:
            height_mm = float(data["alto_cm"]) * 10.0
        else:
            height_mm = defaults.height_mm if defaults else 0.0

        return cls(name=name, span=span, height_mm=height_mm)

    @classmethod
    def dict_from_config(cls, data: Mapping[str, Any]) -> Dict[str, "ImagePreset"]:
        """Genera un diccionario de presets a partir de una configuración genérica."""

        presets = {}
        if not isinstance(data, Mapping):
            return cls.default_presets()

        for name, raw in data.items():
            if not isinstance(raw, Mapping):
                continue
            presets[name] = cls.from_config(name, raw)

        if not presets:
            presets = cls.default_presets()

        return presets


@dataclass
class CapacityModel:
    """Modelo que encapsula los cálculos de capacidad y costos adicionales."""

    text_style: TextStyle
    image_presets: Dict[str, ImagePreset]

    def capacity_per_column(
        self,
        column_height_mm: float,
        span: int = 1,
        title_level: Optional[int] = None,
        image_preset: Optional[str] = None,
    ) -> int:
        """Calcula la capacidad disponible por columna considerando reservas."""

        reserved_height = self.text_style.title_cost(title_level, span)

        if image_preset:
            try:
                reserved_height += self.image_cost(image_preset, span)
            except KeyError:
                logging.warning("Preset de imagen '%s' no encontrado", image_preset)

        available_height = max(column_height_mm - reserved_height, 0.0)
        per_column = self.text_style.capacity_for_height(available_height)
        return per_column * max(span, 1)

    def title_cost(self, level: Optional[int], span: int = 1) -> float:
        """Interfaz directa para obtener el costo de título."""

        return self.text_style.title_cost(level, span)

    def image_cost(self, name: str, span: Optional[int] = None) -> float:
        """Devuelve el costo vertical del preset solicitado."""

        preset = self.image_presets.get(name)
        if preset is None:
            raise KeyError(name)
        return preset.cost(span if span is not None else preset.span)


@dataclass
class Config:
    """Parámetros capturados desde la consola."""

    layout_path: Path
    pages: List[int]
    cierre_root: Path
    output_dir: Path
    slot_selector: Dict[str, str] = field(default_factory=dict)
    typography: TextStyle = field(default_factory=TextStyle.defaults)
    image_presets: Dict[str, ImagePreset] = field(default_factory=ImagePreset.default_presets)
    config_path: Optional[Path] = None

    def build_capacity_model(self) -> CapacityModel:
        """Crea una instancia de :class:`CapacityModel` con los parámetros vigentes."""

        return CapacityModel(text_style=self.typography, image_presets=dict(self.image_presets))


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
    slots_mm: Tuple[Tuple[float, float, float, float], ...] = field(default_factory=tuple)


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
    note_id: Optional[str] = None
    title_height_mm: float = 0.0
    body_height_mm: float = 0.0
    image_height_mm: float = 0.0
    img_mode: str = "none"
    body_chars_fit: int = 0
    body_chars_overflow: int = 0


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
class ColumnFootprint:
    """Huella de columnas estimada para una variante de nota."""

    span: int
    column_heights_mm: Tuple[float, ...]
    penalties: Dict[str, float]
    image_priority: int
    image_preset: Optional[str] = None


@dataclass
class NoteVariant:
    """Descripción de una configuración posible para una nota."""

    note_id: int
    title_span: int
    span: int
    image_preset: Optional[str]
    total_height_mm: float
    text_height_mm: float
    footprint: ColumnFootprint

    def preference_key(self) -> Tuple[int, int, float]:
        """Genera la llave de ordenamiento por preferencia."""

        return (
            self.title_span,
            _image_order_index(self.image_preset),
            self.total_height_mm,
        )


def _image_order_index(image: Optional[str]) -> int:
    """Define el orden de preferencia para las imágenes."""

    if image is None:
        return 2
    if image.lower() == "vertical":
        return 1
    return 0


def _image_priority(image: Optional[str]) -> int:
    """Asigna prioridad a la preservación de la imagen."""

    if image is None:
        return 0
    if image.lower() == "vertical":
        return 1
    return 2


def estimate_text_height_mm(note: Note, model: CapacityModel, span: int) -> float:
    """Calcula la altura necesaria para alojar el cuerpo de la nota."""

    span = max(span, 1)
    chars_per_column = model.text_style.chars_per_line * span
    if chars_per_column <= 0:
        return 0.0
    estimated_lines = note.chars_body / float(chars_per_column)
    return estimated_lines / model.text_style.lines_per_mm if model.text_style.lines_per_mm else 0.0


def generate_note_variants(
    note: Note,
    model: CapacityModel,
    title_spans: Sequence[int] = (1, 2),
    title_level: int = 1,
) -> List[NoteVariant]:
    """Calcula variantes ordenadas por preferencia para la nota dada."""

    available_presets = {
        name: preset for name, preset in model.image_presets.items()
    }
    image_options: List[Optional[str]] = [None]

    if "vertical" in available_presets:
        image_options.append("vertical")
    if "horizontal" in available_presets:
        image_options.append("horizontal")
    else:
        for name, preset in available_presets.items():
            if preset.span >= 2:
                image_options.append(name)
                break

    variants: List[NoteVariant] = []

    for title_span in title_spans:
        if title_span <= 0:
            continue
        title_height_total = model.title_cost(title_level, span=title_span)
        title_height_per_column = title_height_total / float(title_span)

        for image_name in image_options:
            preset = available_presets.get(image_name) if image_name else None
            image_span = preset.span if preset else 0
            span = max(title_span, image_span, 1)

            text_height = estimate_text_height_mm(note, model, span)

            image_height_total = model.image_cost(image_name, span) if image_name else 0.0
            image_height_per_column = image_height_total / float(image_span) if image_span else 0.0

            column_heights: List[float] = []
            for col in range(span):
                col_height = text_height
                if col < title_span:
                    col_height += title_height_per_column
                if image_span and col < image_span:
                    col_height += image_height_per_column
                column_heights.append(col_height)

            total_height = max(column_heights) if column_heights else 0.0
            available_chars = model.capacity_per_column(
                total_height,
                span=span,
                title_level=title_level,
                image_preset=image_name,
            )
            overflow = max(note.chars_body - available_chars, 0)
            slack = max(available_chars - note.chars_body, 0)

            penalties: Dict[str, float] = {}
            if overflow:
                penalties["overflow_chars"] = float(overflow)
            if slack:
                penalties["unused_capacity"] = float(slack)
            if image_name is None:
                penalties.setdefault("missing_image", 1.0)

            footprint = ColumnFootprint(
                span=span,
                column_heights_mm=tuple(column_heights),
                penalties=penalties,
                image_priority=_image_priority(image_name),
                image_preset=image_name,
            )

            variants.append(
                NoteVariant(
                    note_id=note.note_id,
                    title_span=title_span,
                    span=span,
                    image_preset=image_name,
                    total_height_mm=total_height,
                    text_height_mm=text_height,
                    footprint=footprint,
                )
            )

    variants.sort(key=lambda variant: variant.preference_key())
    return variants


def generate_variants_for_notes(
    notes: Sequence[Note],
    model: CapacityModel,
    title_spans: Sequence[int] = (1, 2),
    title_level: int = 1,
) -> Dict[int, List[NoteVariant]]:
    """Calcula variantes para todas las notas en la colección."""

    variants: Dict[int, List[NoteVariant]] = {}
    for note in notes:
        variants[note.note_id] = generate_note_variants(
            note,
            model,
            title_spans=title_spans,
            title_level=title_level,
        )
    return variants


@dataclass
class PipelineStats:
    """Estadísticas básicas para el log final."""

    pages_processed: int = 0
    total_blocks: int = 0
    notes_processed: int = 0
    warnings: List[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Entrada por consola


def load_capacity_settings_from_file(path: Path) -> Tuple[TextStyle, Dict[str, ImagePreset]]:
    """Carga parámetros tipográficos e imágenes desde un archivo JSON."""

    raw_text = path.read_text(encoding="utf-8")
    payload = json.loads(raw_text)

    typography_data = payload.get("typography") or payload.get("tipografia") or {}
    image_data = payload.get("image_presets") or payload.get("imagenes") or payload.get("imágenes") or {}

    typography = TextStyle.from_config(typography_data)
    image_presets = ImagePreset.dict_from_config(image_data)

    return typography, image_presets


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

    layout_raw = input("Ruta del layout_slots.json: ")
    pages_raw = input(
        f"Páginas a procesar (ENTER para {default_pages}): "
    ).strip()
    cierre_raw = input("Carpeta del cierre: ")
    output_raw = input("Carpeta de salida: ")
    config_file_raw = input(
        "Archivo de configuración tipográfica (ENTER para defaults): "
    ).strip()

    if not pages_raw:
        pages_raw = default_pages

    pages = [int(p.strip()) for p in pages_raw.split(",") if p.strip()]

    config_path = _clean_path(config_file_raw) if config_file_raw else None

    typography = TextStyle.defaults()
    image_presets = ImagePreset.default_presets()

    if config_path:
        try:
            typography, image_presets = load_capacity_settings_from_file(config_path)
        except Exception as exc:  # pragma: no cover - errores de IO/formato
            logging.warning("No se pudo cargar la configuración %s: %s", config_path, exc)
            config_path = None

    cfg = Config(
        layout_path=_clean_path(layout_raw),
        pages=pages,
        cierre_root=_clean_path(cierre_raw),
        output_dir=_clean_path(output_raw),
        slot_selector={"layer": "ESPACIO_NOTAS"},
        typography=typography,
        image_presets=image_presets,
        config_path=config_path,
    )

    return cfg


# ---------------------------------------------------------------------------
# Orquestación


def run_pipeline(cfg: Config) -> None:
    """Ejecuta el flujo completo definido en el diseño."""

    logging.info("Iniciando pipeline con layout %s", cfg.layout_path)

    page_geometries = load_page_geometry_from_layout_json(cfg.layout_path, cfg.slot_selector)

    target_pages = filter_target_pages(page_geometries, cfg.pages)

    capacity_model = cfg.build_capacity_model()
    logging.debug(
        "Modelo de capacidad inicializado — chars/linea=%s líneas/mm=%s presets=%s",
        capacity_model.text_style.chars_per_line,
        capacity_model.text_style.lines_per_mm,
        list(capacity_model.image_presets.keys()),
    )

    capacity_summary: Dict[int, Dict[str, int]] = {}

    blocks: List[Block] = []
    stats = PipelineStats()
    plan_settings = PlanSettings()

    page_map = scan_closure_folder(cfg.cierre_root)
    variants_by_page: Dict[str, Dict[int, List[NoteVariant]]] = {}

    for page_geom in target_pages:
        issues = validate_page_geometry(page_geom)
        if issues:
            stats.warnings.extend(issues)

        column_height = page_geom.usable_rect_mm[3]
        base_capacity = capacity_model.capacity_per_column(column_height)
        capacity_summary[page_geom.page_number] = {"span1": base_capacity}

        notes = extract_notes_for_page(page_map.get(page_geom.name))
        write_out_txt(page_geom.name, notes, cfg.output_dir)
        stats.notes_processed += len(notes)

        outcome = solve_page_layout(page_geom, notes, capacity_model, plan_settings)

        left_mm, top_mm, _, _ = page_geom.usable_rect_mm
        column_step = page_geom.column_width_mm + page_geom.gutter_mm
        page_blocks: List[Block] = []
        for assignment in outcome.assignments:
            x_mm = left_mm + assignment.column_index * column_step
            y_mm = top_mm + assignment.start_mm
            note_identifier = f"{page_geom.name}#{getattr(assignment.note, 'note_id', '?')}"
            page_blocks.append(
                Block(
                    page=page_geom.page_number,
                    column_index=assignment.column_index + 1,
                    span=1,
                    x_mm=x_mm,
                    y_mm=y_mm,
                    w_mm=page_geom.column_width_mm,
                    h_mm=assignment.used_height_mm,
                    note_id=note_identifier,
                    title_height_mm=assignment.title_height_mm,
                    body_height_mm=assignment.body_height_mm,
                    image_height_mm=assignment.image_height_mm,
                    img_mode=assignment.img_mode,
                    body_chars_fit=assignment.body_chars_fit,
                    body_chars_overflow=assignment.body_chars_overflow,
                )
            )

        blocks.extend(page_blocks)
        stats.pages_processed += 1
        stats.total_blocks += len(page_blocks)

        log_page_summary(page_geom.name, len(page_blocks), issues)
        logging.info("Página %s · solver: %s", page_geom.name, summarize_outcome(outcome))
        overflow_blocks = [block for block in page_blocks if block.body_chars_overflow > 0]
        if overflow_blocks:
            total_overflow = sum(block.body_chars_overflow for block in overflow_blocks)
            logging.info(
                "Página %s: overflow total=%s chars en %s bloques",
                page_geom.name,
                total_overflow,
                len(overflow_blocks),
            )
        for entry in outcome.logs:
            logging.debug("Página %s · %s", page_geom.name, entry)
        if outcome.dropped_notes:
            dropped_titles = [
                getattr(note, "title", f"Nota {getattr(note, 'note_id', '?')}") or f"Nota {getattr(note, 'note_id', '?')}"
                for note in outcome.dropped_notes
            ]
            logging.warning(
                "Página %s: %s notas sin ubicar (%s)",
                page_geom.name,
                len(dropped_titles),
                "; ".join(dropped_titles),
            )
        if notes:
            note_variants = generate_variants_for_notes(notes, capacity_model)
            variants_by_page[page_geom.name] = note_variants
            logging.debug(
                "Variantes generadas para %s: %s",
                page_geom.name,
                {
                    note_id: [
                        {
                            "title_span": variant.title_span,
                            "span": variant.span,
                            "image": variant.image_preset,
                            "height_mm": round(variant.total_height_mm, 2),
                        }
                        for variant in variants
                    ]
                    for note_id, variants in note_variants.items()
                },
            )

    logging.debug("Resumen de capacidad por página: %s", capacity_summary)

    save_plan_blocks_json(blocks, cfg.output_dir)
    save_plan_blocks_csv(blocks, cfg.output_dir)
    log_global_summary(stats)

# ---------------------------------------------------------------------------
# 1) Layout JSON → geometrías
#

def load_page_geometry_from_layout_json(
    layout_path: Path, slot_selector: Optional[Mapping[str, str]] = None
) -> List[PageGeometry]:
    """Carga geometrías a partir del ``layout_slots.json`` generado por JSX."""

    if not layout_path.exists():
        raise FileNotFoundError(layout_path)

    with layout_path.open(encoding="utf-8") as fh:
        data = json.load(fh)

    pages_raw = data.get("pages", [])

    expected_label: Optional[str] = None
    if slot_selector and isinstance(slot_selector, Mapping):
        raw_label = slot_selector.get("label")
        if isinstance(raw_label, str) and raw_label.strip():
            expected_label = raw_label.strip()

    if not expected_label:
        expected_label = "ESPACIO_NOTAS"

    def _value_mm(entry: Mapping[str, Any], prefix: str) -> Optional[float]:
        key_mm = f"{prefix}_mm"
        key_pt = f"{prefix}_pt"
        raw_mm = entry.get(key_mm) if isinstance(entry, Mapping) else None
        if raw_mm not in (None, ""):
            try:
                return float(raw_mm)
            except (TypeError, ValueError):
                pass
        raw_pt = entry.get(key_pt) if isinstance(entry, Mapping) else None
        if raw_pt not in (None, ""):
            try:
                return float(raw_pt) * POINT_TO_MM
            except (TypeError, ValueError):
                return None
        return None

    def _round_mm(value: float) -> float:
        return round(value, 1)

    geometries: List[PageGeometry] = []

    for page_entry in pages_raw:
        try:
            page_number = int(page_entry.get("page_index") or page_entry.get("page_number") or 0)
        except (TypeError, ValueError):
            page_number = 0
        page_name = str(page_entry.get("page_name", page_number or "")) or str(page_number)

        columns_info = page_entry.get("columns") if isinstance(page_entry, Mapping) else {}
        try:
            columns = int(columns_info.get("count")) if columns_info else 0
        except (TypeError, ValueError):
            columns = 0
        gutter_mm = 0.0
        if isinstance(columns_info, Mapping):
            try:
                gutter_mm = float(columns_info.get("gutter_mm") or 0.0)
            except (TypeError, ValueError):
                gutter_mm = 0.0
        column_width_mm = 0.0
        if isinstance(columns_info, Mapping):
            try:
                column_width_mm = float(columns_info.get("col_width_mm") or 0.0)
            except (TypeError, ValueError):
                column_width_mm = 0.0

        slots_raw = page_entry.get("slots") if isinstance(page_entry, Mapping) else []
        slot_rects: List[Tuple[float, float, float, float]] = []
        for slot in slots_raw or []:
            if not isinstance(slot, Mapping):
                continue
            label = str(slot.get("label", "")).strip()
            if expected_label and label != expected_label:
                continue
            left = _value_mm(slot, "left")
            top = _value_mm(slot, "top")
            width = _value_mm(slot, "width")
            height = _value_mm(slot, "height")
            if None in {left, top, width, height}:
                continue
            slot_rects.append(tuple(_round_mm(v) for v in (left, top, width, height)))

        if slot_rects:
            primary_slot = max(slot_rects, key=lambda rect: rect[2] * rect[3])
        else:
            primary_slot = (0.0, 0.0, 0.0, 0.0)
            logging.warning("Página %s: no se encontraron slots '%s'", page_name, expected_label)

        if columns > 0 and column_width_mm <= 0.0 and primary_slot[2] > 0.0:
            inner_width = max(primary_slot[2] - gutter_mm * max(columns - 1, 0), 0.0)
            column_width_mm = inner_width / columns if columns else 0.0

        margins_entry = page_entry.get("margins") if isinstance(page_entry, Mapping) else {}
        margins: Dict[str, float] = {}
        for key in ("left", "top", "right", "bottom"):
            mm_value = _value_mm(margins_entry or {}, key)
            if mm_value is not None:
                margins[key] = _round_mm(mm_value)

        page_size_entry = page_entry.get("page_size") if isinstance(page_entry, Mapping) else {}
        page_width_mm = _value_mm(page_size_entry or {}, "width") or 0.0
        page_height_mm = _value_mm(page_size_entry or {}, "height") or 0.0

        if "left" not in margins:
            margins["left"] = _round_mm(primary_slot[0])
        if "top" not in margins:
            margins["top"] = _round_mm(primary_slot[1])
        if "right" not in margins and page_width_mm > 0.0:
            margins["right"] = _round_mm(max(page_width_mm - (primary_slot[0] + primary_slot[2]), 0.0))
        if "bottom" not in margins and page_height_mm > 0.0:
            margins["bottom"] = _round_mm(max(page_height_mm - (primary_slot[1] + primary_slot[3]), 0.0))

        geometries.append(
            PageGeometry(
                page_number=page_number,
                name=page_name,
                usable_rect_mm=primary_slot,
                columns=columns,
                gutter_mm=_round_mm(gutter_mm) if gutter_mm else 0.0,
                column_width_mm=_round_mm(column_width_mm) if column_width_mm else column_width_mm,
                margins_mm=margins,
                slots_mm=tuple(slot_rects),
            )
        )

    geometries.sort(key=lambda geom: geom.page_number)
    return geometries


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

        # Intentamos descartar candidatos que evidentemente no pertenecen al área
        # útil de la página.  En particular, ignoramos los que comienzan con
        # coordenadas negativas, ya que suelen corresponder a fondos o elementos
        # sangrados que desbordan el pliego.
        filtered_rects = [
            entry for entry in rects if entry[2][0] >= 0.0 and entry[2][1] >= 0.0
        ]

        rects_to_consider = filtered_rects if filtered_rects else rects

        rects_to_consider.sort(key=lambda entry: (entry[0], entry[1]), reverse=True)
        usable[page_id] = rects_to_consider[0][2]

    return usable


def extract_grid_by_page(pages):
    """Determina la retícula de columnas por página."""

    grid: Dict[str, Dict[str, float]] = {}
    for page in pages:
        width_mm = page.bounds.width * POINT_TO_MM
        columns = 5 if width_mm >= 250.0 else 4

        # La página 3 reserva una columna para staff, por lo que solo quedan
        # cuatro columnas útiles para notas.
        if page.index == 3 or page.name.strip() == "3":
            columns = 4
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
        writer.writerow(
            [
                "page",
                "i",
                "span",
                "x_mm",
                "y_mm",
                "w_mm",
                "h_mm",
                "note_id",
                "title_height_mm",
                "body_height_mm",
                "image_height_mm",
                "img_mode",
                "body_chars_fit",
                "body_chars_overflow",
            ]
        )
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
                    block.note_id or "",
                    f"{block.title_height_mm:.2f}",
                    f"{block.body_height_mm:.2f}",
                    f"{block.image_height_mm:.2f}",
                    block.img_mode,
                    block.body_chars_fit,
                    block.body_chars_overflow,
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

