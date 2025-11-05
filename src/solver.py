"""Beam-search solver responsable de ubicar notas en la retícula."""

from __future__ import annotations

from dataclasses import dataclass
import math
from types import SimpleNamespace
from typing import Dict, Iterable, List, Mapping, Optional, Sequence, Tuple, TYPE_CHECKING

if TYPE_CHECKING:  # pragma: no cover - solo para type checkers
    from .Prediseño_automatizado import CapacityModel, Note, NoteVariant, PageGeometry


# ---------------------------------------------------------------------------
# Modelos públicos


@dataclass(frozen=True)
class PlanSettings:
    """Parámetros que gobiernan la búsqueda."""

    beam_width: int = 8
    fit_bonus: float = 8.0
    gap_penalty_per_mm: float = 0.18  # penaliza el desbalance de huecos entre columnas
    overflow_penalty_per_char: float = 0.12
    overfill_penalty_per_mm: float = 0.65
    drop_penalty: float = 10.0
    final_gap_penalty_per_mm: float = 0.25  # castiga el hueco total restante en la página
    default_title_level: Optional[int] = 1
    default_image_preset: Optional[str] = None
    title_level_attr: str = "title_level"
    image_mode_attr: str = "image_mode"
    image_count_attr: str = "image_count"
    body_chars_attr: str = "chars_body"
    unused_capacity_penalty_per_char: float = 0.02
    mordida_penalty_per_line: float = 0.5
    missing_image_penalty: float = 1.5


@dataclass(frozen=True)
class Assignment:
    """Resultado de ubicar una nota en una columna."""

    note: "Note"
    column_index: int
    span: int
    start_mm: float
    used_height_mm: float
    column_heights_mm: Tuple[float, ...]
    column_body_heights_mm: Tuple[float, ...]
    column_title_heights_mm: Tuple[float, ...]
    column_image_heights_mm: Tuple[float, ...]
    title_height_mm: float
    body_height_mm: float
    image_height_mm: float
    body_chars_fit: int
    body_chars_overflow: int
    img_mode: str
    fit: bool
    remaining_mm: float
    title_lines: float
    body_lines: float
    image_span: int


@dataclass(frozen=True)
class SolverOutcome:
    """Resumen completo luego de ejecutar el solver para una página."""

    assignments: Tuple[Assignment, ...]
    dropped_notes: Tuple["Note", ...]
    column_usage_mm: Tuple[float, ...]
    column_gaps_mm: Tuple[float, ...]
    logs: Tuple[str, ...]
    score: float


# ---------------------------------------------------------------------------
# Estructuras internas


@dataclass(frozen=True)
class _BeamState:
    """Estado parcial dentro del beam search."""

    score: float
    columns_used_mm: Tuple[float, ...]
    assignments: Tuple[Assignment, ...]
    dropped: Tuple["Note", ...]
    logs: Tuple[str, ...]

    def push_assignment(
        self,
        assignment: Assignment,
        score_delta: float,
        column_height: float,
    ) -> "_BeamState":
        used = list(self.columns_used_mm)
        for offset, height in enumerate(assignment.column_heights_mm):
            idx = assignment.column_index + offset
            if idx >= len(used):
                continue
            start_level = max(used[idx], assignment.start_mm)
            used[idx] = min(column_height, start_level + height)
        first_col = assignment.column_index + 1
        last_col = first_col + assignment.span - 1
        label = f"{first_col}-{last_col}" if assignment.span > 1 else str(first_col)
        logs = self.logs + (
            "columna(s) {label}: nota {note} · span={span} · fit={fit} · rem={rem:.1f}mm · overflow={overflow}".format(
                label=label,
                note=getattr(assignment.note, "note_id", "?"),
                span=assignment.span,
                fit="sí" if assignment.fit else "no",
                rem=assignment.remaining_mm,
                overflow=assignment.body_chars_overflow,
            ),
        )
        return _BeamState(
            score=self.score + score_delta,
            columns_used_mm=tuple(used),
            assignments=self.assignments + (assignment,),
            dropped=self.dropped,
            logs=logs,
        )

    def push_drop(self, note: "Note", penalty: float, reason: str) -> "_BeamState":
        logs = self.logs + (
            "descartar nota {note} ({reason})".format(
                note=getattr(note, "note_id", "?"), reason=reason
            ),
        )
        return _BeamState(
            score=self.score - penalty,
            columns_used_mm=self.columns_used_mm,
            assignments=self.assignments,
            dropped=self.dropped + (note,),
            logs=logs,
        )


# ---------------------------------------------------------------------------
# Utilidades


def _chars_for_note(note: "Note", attr: str) -> int:
    value = getattr(note, attr, None)
    try:
        return int(value or 0)
    except (TypeError, ValueError):
        return 0


def _resolve_title_level(note: "Note", settings: PlanSettings) -> Optional[int]:
    raw = getattr(note, settings.title_level_attr, None)
    if raw is None:
        return settings.default_title_level
    try:
        level = int(raw)
    except (TypeError, ValueError):
        return settings.default_title_level
    return level if level > 0 else settings.default_title_level


def _resolve_image_mode(
    note: "Note", capacity_model: "CapacityModel", settings: PlanSettings
) -> Tuple[str, float]:
    def _fallback() -> Tuple[str, float]:
        preset_name = settings.default_image_preset
        if preset_name and preset_name in capacity_model.image_presets:
            preset = capacity_model.image_presets[preset_name]
        else:
            presets = list(capacity_model.image_presets.values())
            if not presets:
                return "none", 0.0
            preset = presets[0]
        return preset.name, preset.cost()

    if not capacity_model.image_presets:
        return "none", 0.0

    raw_mode = getattr(note, settings.image_mode_attr, None)
    if isinstance(raw_mode, str) and raw_mode in capacity_model.image_presets:
        preset = capacity_model.image_presets[raw_mode]
        return preset.name, preset.cost()

    raw_count = getattr(note, settings.image_count_attr, None)
    try:
        count = int(raw_count or 0)
    except (TypeError, ValueError):
        count = 0
    if count <= 0:
        return "none", 0.0

    return _fallback()


def _body_height_from_chars(chars: int, capacity_model: "CapacityModel") -> float:
    density = capacity_model.text_style.chars_per_line * capacity_model.text_style.lines_per_mm
    if density <= 0:
        return 0.0
    return float(chars) / density


@dataclass
class _PlacementEvaluation:
    assignment: Assignment
    score_delta: float


def _variant_span(variant: object) -> int:
    try:
        span = int(getattr(variant, "span", 1) or 1)
    except (TypeError, ValueError):
        span = 1
    return max(span, 1)


def _is_number(value: object) -> bool:
    try:
        float(value)
    except (TypeError, ValueError):
        return False
    return True


def _take_span(values: Iterable[float], span: int) -> Tuple[float, ...]:
    result: List[float] = []
    for value in values:
        result.append(float(value))
        if len(result) >= span:
            break
    if len(result) < span:
        result.extend(0.0 for _ in range(span - len(result)))
    return tuple(result)


def _evaluate_variant_placement(
    note: "Note",
    variant: "NoteVariant",
    column_index: int,
    start_mm: float,
    column_height: float,
    settings: PlanSettings,
) -> Optional[_PlacementEvaluation]:
    span = _variant_span(variant)
    footprint = getattr(variant, "footprint", None)
    if footprint is None:
        return None

    column_heights = _take_span(getattr(footprint, "column_heights_mm", ()), span)
    if not column_heights:
        return None

    tolerance = 1e-6
    end_levels: List[float] = []
    for height in column_heights:
        end = start_mm + height
        if end - column_height > tolerance:
            return None
        end_levels.append(end)

    remaining_mm = min(column_height - end for end in end_levels) if end_levels else column_height

    body_chars_overflow = int(getattr(variant, "body_chars_overflow", 0) or 0)
    fit = body_chars_overflow <= 0 and remaining_mm >= -tolerance

    score_delta = 0.0
    if fit:
        score_delta += settings.fit_bonus
    else:
        if body_chars_overflow > 0:
            score_delta -= settings.overflow_penalty_per_char * body_chars_overflow

    penalties: Dict[str, float] = {}
    raw_penalties = getattr(footprint, "penalties", None)
    if isinstance(raw_penalties, Mapping):
        penalties = {str(key): float(value) for key, value in raw_penalties.items() if _is_number(value)}

    unused_capacity = penalties.get("unused_capacity", 0.0)
    if unused_capacity:
        score_delta -= settings.unused_capacity_penalty_per_char * unused_capacity

    title_mordida = penalties.get("title_mordida_lines", 0.0)
    if title_mordida:
        score_delta -= settings.mordida_penalty_per_line * title_mordida

    body_mordida = penalties.get("body_mordida_lines", 0.0)
    if body_mordida:
        score_delta -= settings.mordida_penalty_per_line * body_mordida

    missing_image_penalty = penalties.get("missing_image", 0.0)
    if missing_image_penalty:
        score_delta -= settings.missing_image_penalty * missing_image_penalty

    for key, value in penalties.items():
        if key in {"overflow_chars", "unused_capacity", "title_mordida_lines", "body_mordida_lines", "missing_image"}:
            continue
        score_delta -= value

    image_priority = getattr(footprint, "image_priority", 0)
    try:
        score_delta += float(image_priority) * 0.05
    except (TypeError, ValueError):
        pass

    column_body = _take_span(getattr(footprint, "column_body_heights_mm", ()), span)
    column_title = _take_span(getattr(footprint, "column_title_heights_mm", ()), span)
    column_image = _take_span(getattr(footprint, "column_image_heights_mm", ()), span)

    total_height = float(getattr(variant, "total_height_mm", max(column_heights) if column_heights else 0.0))
    title_height = float(getattr(variant, "title_height_mm", max(column_title) if column_title else 0.0))
    body_height = float(getattr(variant, "body_height_mm", max(column_body) if column_body else 0.0))
    image_height = float(getattr(variant, "image_height_mm", max(column_image) if column_image else 0.0))

    assignment = Assignment(
        note=note,
        column_index=column_index,
        span=span,
        start_mm=start_mm,
        used_height_mm=total_height,
        column_heights_mm=column_heights,
        column_body_heights_mm=column_body,
        column_title_heights_mm=column_title,
        column_image_heights_mm=column_image,
        title_height_mm=title_height,
        body_height_mm=body_height,
        image_height_mm=image_height,
        body_chars_fit=int(getattr(variant, "body_chars_fit", 0) or 0),
        body_chars_overflow=body_chars_overflow,
        img_mode=str(getattr(variant, "image_preset", None) or "none"),
        fit=fit,
        remaining_mm=max(remaining_mm, 0.0),
        title_lines=float(getattr(variant, "title_lines", 0.0)),
        body_lines=float(getattr(variant, "body_lines", 0.0)),
        image_span=int(getattr(variant, "image_span", 0) or 0),
    )

    return _PlacementEvaluation(assignment=assignment, score_delta=score_delta)


def _lookup_variants(
    note_id: object,
    variants_by_note: Mapping[object, Sequence["NoteVariant"]],
) -> Sequence["NoteVariant"]:
    if note_id in variants_by_note:
        return variants_by_note[note_id]

    candidates: List[object] = []
    try:
        if note_id is not None:
            candidates.append(int(note_id))
    except (TypeError, ValueError):
        pass
    try:
        if note_id is not None:
            candidates.append(str(note_id))
    except Exception:  # pragma: no cover - conversión defensiva
        pass

    for key in candidates:
        if key in variants_by_note:
            return variants_by_note[key]

    return ()


def _build_fallback_variants(
    note: "Note",
    capacity_model: "CapacityModel",
    column_height: float,
    settings: PlanSettings,
) -> Sequence[object]:
    title_level = _resolve_title_level(note, settings)
    title_height = capacity_model.title_cost(title_level) if title_level else 0.0
    img_mode, image_height = _resolve_image_mode(note, capacity_model, settings)

    lines_per_mm = capacity_model.text_style.lines_per_mm or 1.0
    chars_per_line = capacity_model.text_style.chars_per_line or 1
    body_chars = _chars_for_note(note, settings.body_chars_attr)

    available_body_height = max(column_height - title_height - image_height, 0.0)
    if body_chars > 0 and chars_per_line > 0:
        required_lines = math.ceil(body_chars / chars_per_line)
        required_height = required_lines / lines_per_mm
        body_height = min(required_height, available_body_height)
    else:
        body_height = 0.0

    capacity_chars = capacity_model.text_style.capacity_for_height(body_height)
    body_chars_fit = min(capacity_chars, body_chars)
    overflow = max(body_chars - body_chars_fit, 0)

    penalties: Dict[str, float] = {}
    if overflow:
        penalties["overflow_chars"] = float(overflow)
    if img_mode == "none":
        penalties["missing_image"] = 1.0

    footprint = SimpleNamespace(
        span=1,
        column_heights_mm=(title_height + image_height + body_height,),
        column_body_heights_mm=(body_height,),
        column_title_heights_mm=(title_height,),
        column_image_heights_mm=(image_height,),
        penalties=penalties,
        image_priority=0,
        image_preset=None if img_mode == "none" else img_mode,
    )

    variant = SimpleNamespace(
        note_id=getattr(note, "note_id", 0),
        title_span=1 if title_height else 0,
        span=1,
        image_preset=None if img_mode == "none" else img_mode,
        total_height_mm=title_height + image_height + body_height,
        text_height_mm=body_height,
        title_height_mm=title_height,
        body_height_mm=body_height,
        image_height_mm=image_height,
        body_chars_fit=body_chars_fit,
        body_chars_overflow=overflow,
        title_lines=title_height * lines_per_mm,
        body_lines=body_height * lines_per_mm,
        image_span=1 if img_mode != "none" else 0,
        footprint=footprint,
    )

    return (variant,)


def _variants_for_note(
    note: "Note",
    variants_by_note: Optional[Mapping[object, Sequence["NoteVariant"]]],
    capacity_model: "CapacityModel",
    column_height: float,
    settings: PlanSettings,
) -> Sequence[object]:
    if variants_by_note:
        note_id = getattr(note, "note_id", None)
        if note_id is not None:
            candidates = _lookup_variants(note_id, variants_by_note)
            if candidates:
                return candidates
    return _build_fallback_variants(note, capacity_model, column_height, settings)


# ---------------------------------------------------------------------------
# Búsqueda principal


def solve_page_layout(
    page_geom: "PageGeometry",
    notes: Sequence["Note"],
    capacity_model: "CapacityModel",
    settings: Optional[PlanSettings] = None,
    variants_by_note: Optional[Mapping[object, Sequence["NoteVariant"]]] = None,
) -> SolverOutcome:
    """Ejecuta un beam search sobre las notas y columnas de la página."""

    if settings is None:
        settings = PlanSettings()

    columns = int(getattr(page_geom, "columns", 0) or 0)
    column_height = float(getattr(page_geom, "usable_rect_mm", (0, 0, 0, 0))[3])

    if columns <= 0 or column_height <= 0:
        logs = ("sin columnas disponibles; todas las notas quedan pendientes",)
        return SolverOutcome(
            assignments=(),
            dropped_notes=tuple(notes),
            column_usage_mm=tuple(0.0 for _ in range(max(columns, 0))),
            column_gaps_mm=tuple(column_height for _ in range(max(columns, 0))),
            logs=logs,
            score=-settings.drop_penalty * len(notes),
        )

    initial = _BeamState(
        score=0.0,
        columns_used_mm=tuple(0.0 for _ in range(columns)),
        assignments=(),
        dropped=(),
        logs=(),
    )

    beam: List[_BeamState] = [initial]

    for note in notes:
        new_states: List[_BeamState] = []
        note_variants = _variants_for_note(note, variants_by_note, capacity_model, column_height, settings)
        for state in beam:
            placed = False
            for variant in note_variants:
                span = _variant_span(variant)
                if span > columns:
                    continue
                for col_idx in range(0, columns - span + 1):
                    used_slice = state.columns_used_mm[col_idx : col_idx + span]
                    if len(used_slice) < span:
                        continue
                    start_mm = max(used_slice) if used_slice else 0.0
                    evaluation = _evaluate_variant_placement(
                        note=note,
                        variant=variant,
                        column_index=col_idx,
                        start_mm=start_mm,
                        column_height=column_height,
                        settings=settings,
                    )
                    if evaluation is None or evaluation.assignment.used_height_mm <= 0.0:
                        continue
                    new_states.append(
                        state.push_assignment(evaluation.assignment, evaluation.score_delta, column_height)
                    )
                    placed = True
            if not placed:
                reason = "sin espacio"
                new_states.append(state.push_drop(note, settings.drop_penalty, reason))
        if not new_states:
            beam = [state.push_drop(note, settings.drop_penalty, "sin candidatos") for state in beam]
        else:
            new_states.sort(key=lambda entry: entry.score, reverse=True)
            beam = new_states[: settings.beam_width]

    best_state: Optional[_BeamState] = None
    best_score = float("-inf")
    for state in beam:
        gaps = tuple(max(column_height - used, 0.0) for used in state.columns_used_mm)
        total_gap = sum(gaps)
        penalty = settings.final_gap_penalty_per_mm * total_gap
        if gaps and settings.gap_penalty_per_mm > 0.0:
            max_gap = max(gaps)
            min_gap = min(gaps)
            imbalance = max_gap - min_gap
            if imbalance > 0.0:
                penalty += settings.gap_penalty_per_mm * imbalance
        final_score = state.score - penalty
        if final_score > best_score:
            best_score = final_score
            best_state = state

    if best_state is None:
        best_state = initial
        best_score = initial.score

    column_usage = best_state.columns_used_mm
    column_gaps = tuple(max(column_height - used, 0.0) for used in column_usage)

    final_logs = best_state.logs + (
        "huecos finales: "
        + ", ".join(f"col {idx + 1} → {gap:.1f}mm" for idx, gap in enumerate(column_gaps)),
    )

    return SolverOutcome(
        assignments=best_state.assignments,
        dropped_notes=best_state.dropped,
        column_usage_mm=column_usage,
        column_gaps_mm=column_gaps,
        logs=final_logs,
        score=best_score,
    )


def summarize_outcome(outcome: SolverOutcome) -> str:
    """Genera un texto compacto útil para los logs."""

    placed = len(outcome.assignments)
    dropped = len(outcome.dropped_notes)
    gaps = ", ".join(f"{gap:.1f}mm" for gap in outcome.column_gaps_mm)
    return f"{placed} ubicadas · {dropped} pendientes · huecos: [{gaps}]"
