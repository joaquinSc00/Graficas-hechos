"""Beam-search solver responsable de ubicar notas en la retícula."""

from __future__ import annotations

from dataclasses import dataclass
from typing import List, Optional, Sequence, Tuple, TYPE_CHECKING

if TYPE_CHECKING:  # pragma: no cover - solo para type checkers
    from .Prediseño_automatizado import CapacityModel, Note, PageGeometry


# ---------------------------------------------------------------------------
# Modelos públicos


@dataclass(frozen=True)
class PlanSettings:
    """Parámetros que gobiernan la búsqueda."""

    beam_width: int = 8
    fit_bonus: float = 8.0
    gap_penalty_per_mm: float = 0.18
    overflow_penalty_per_char: float = 0.12
    overfill_penalty_per_mm: float = 0.65
    drop_penalty: float = 10.0
    final_gap_penalty_per_mm: float = 0.25
    default_title_level: Optional[int] = 1
    default_image_preset: Optional[str] = None
    title_level_attr: str = "title_level"
    image_mode_attr: str = "image_mode"
    image_count_attr: str = "image_count"
    body_chars_attr: str = "chars_body"


@dataclass(frozen=True)
class Assignment:
    """Resultado de ubicar una nota en una columna."""

    note: "Note"
    column_index: int
    start_mm: float
    used_height_mm: float
    title_height_mm: float
    body_height_mm: float
    image_height_mm: float
    body_chars_fit: int
    body_chars_overflow: int
    img_mode: str
    fit: bool
    remaining_mm: float


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
        idx = assignment.column_index
        used[idx] = min(column_height, used[idx] + assignment.used_height_mm)
        logs = self.logs + (
            "columna {idx}: nota {note} · fit={fit} · rem={rem:.1f}mm · overflow={overflow}".format(
                idx=idx + 1,
                note=getattr(assignment.note, "note_id", "?"),
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


def _evaluate_placement(
    note: "Note",
    column_index: int,
    start_mm: float,
    available_mm: float,
    capacity_model: "CapacityModel",
    settings: PlanSettings,
) -> _PlacementEvaluation:
    title_level = _resolve_title_level(note, settings)
    title_height = capacity_model.title_cost(title_level) if title_level else 0.0
    img_mode, image_height = _resolve_image_mode(note, capacity_model, settings)

    reserved_height = title_height + image_height
    usable_height = max(available_mm - reserved_height, 0.0)

    total_chars = _chars_for_note(note, settings.body_chars_attr)
    if total_chars <= 0:
        body_chars_fit = 0
        body_height = 0.0
    else:
        # Cálculo conservador de capacidad dado el espacio disponible
        cap = capacity_model.text_style.capacity_for_height(usable_height)
        body_chars_fit = min(cap, total_chars)
        body_height = _body_height_from_chars(body_chars_fit, capacity_model)

    body_chars_overflow = max(total_chars - body_chars_fit, 0)
    used_height = reserved_height + body_height
    if used_height > available_mm:
        used_height = available_mm

    ideal_height = reserved_height + _body_height_from_chars(total_chars, capacity_model)
    remaining_mm = max(available_mm - used_height, 0.0)
    fit = body_chars_overflow == 0 and ideal_height <= available_mm + 1e-6

    score_delta = 0.0
    if fit:
        score_delta += settings.fit_bonus
    else:
        score_delta -= settings.overflow_penalty_per_char * body_chars_overflow
        overflow_mm = max(ideal_height - available_mm, 0.0)
        score_delta -= settings.overfill_penalty_per_mm * overflow_mm

    score_delta -= settings.gap_penalty_per_mm * remaining_mm

    assignment = Assignment(
        note=note,
        column_index=column_index,
        start_mm=start_mm,
        used_height_mm=used_height,
        title_height_mm=title_height,
        body_height_mm=body_height,
        image_height_mm=image_height,
        body_chars_fit=body_chars_fit,
        body_chars_overflow=body_chars_overflow,
        img_mode=img_mode,
        fit=fit,
        remaining_mm=remaining_mm,
    )

    return _PlacementEvaluation(assignment=assignment, score_delta=score_delta)


# ---------------------------------------------------------------------------
# Búsqueda principal


def solve_page_layout(
    page_geom: "PageGeometry",
    notes: Sequence["Note"],
    capacity_model: "CapacityModel",
    settings: Optional[PlanSettings] = None,
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
        for state in beam:
            placed = False
            for col_idx in range(columns):
                used = state.columns_used_mm[col_idx]
                available = max(column_height - used, 0.0)
                if available <= 0.0:
                    continue
                evaluation = _evaluate_placement(
                    note=note,
                    column_index=col_idx,
                    start_mm=used,
                    available_mm=available,
                    capacity_model=capacity_model,
                    settings=settings,
                )
                if evaluation.assignment.used_height_mm <= 0.0:
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
        penalty = settings.final_gap_penalty_per_mm * sum(gaps)
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
