"""Beam-search solver that plans note placement column by column."""

from __future__ import annotations

from dataclasses import dataclass
from typing import List, Optional, Sequence, Tuple, TYPE_CHECKING

if TYPE_CHECKING:  # pragma: no cover - hints only, avoid circular imports
    from Prediseño_automatizado import CapacityModel, Note, PageGeometry


@dataclass(frozen=True)
class PlanSettings:
    """Hyperparameters that shape the beam-search exploration."""

    beam_width: int = 6
    fit_bonus: float = 6.0
    gap_penalty_per_mm: float = 0.25
    overflow_penalty_per_char: float = 0.08
    overfill_penalty_per_mm: float = 1.0
    drop_penalty: float = 12.0
    final_gap_penalty_per_mm: float = 0.35
    default_title_level: Optional[int] = 1
    default_image_preset: Optional[str] = None


@dataclass(frozen=True)
class Assignment:
    """Decision taken by the solver for one note."""

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
    """Result of running the beam-search solver over a page."""

    assignments: Tuple[Assignment, ...]
    dropped_notes: Tuple["Note", ...]
    column_usage_mm: Tuple[float, ...]
    column_gaps_mm: Tuple[float, ...]
    logs: Tuple[str, ...]
    score: float


@dataclass(frozen=True)
class _BeamState:
    """Internal representation of a partial exploration state."""

    score: float
    columns_used_mm: Tuple[float, ...]
    assignments: Tuple[Assignment, ...]
    dropped: Tuple["Note", ...]
    logs: Tuple[str, ...]

    def with_assignment(self, assignment: Assignment, score_delta: float, column_height: float) -> "_BeamState":
        used = list(self.columns_used_mm)
        idx = assignment.column_index
        used[idx] = min(column_height, used[idx] + assignment.used_height_mm)
        logs = self.logs + (
            f"columna {idx + 1}: nota {getattr(assignment.note, 'note_id', '?')}"
            f" → fit={'sí' if assignment.fit else 'no'}"
            f" rem={assignment.remaining_mm:.1f}mm overflow={assignment.body_chars_overflow}",
        )
        return _BeamState(
            score=self.score + score_delta,
            columns_used_mm=tuple(used),
            assignments=self.assignments + (assignment,),
            dropped=self.dropped,
            logs=logs,
        )

    def with_drop(self, note: "Note", penalty: float) -> "_BeamState":
        logs = self.logs + (
            f"descartar nota {getattr(note, 'note_id', '?')} (sin espacio viable)",
        )
        return _BeamState(
            score=self.score - penalty,
            columns_used_mm=self.columns_used_mm,
            assignments=self.assignments,
            dropped=self.dropped + (note,),
            logs=logs,
        )


def _body_height_from_chars(chars: int, capacity_model: "CapacityModel") -> float:
    density = capacity_model.text_style.chars_per_line * capacity_model.text_style.lines_per_mm
    if density <= 0:
        return 0.0
    return float(chars) / density


def _select_image_preset(note: "Note", capacity_model: "CapacityModel", settings: PlanSettings) -> Tuple[str, float]:
    image_count = int(getattr(note, "image_count", 0) or 0)
    if image_count <= 0:
        return "none", 0.0

    preset_name = settings.default_image_preset
    if preset_name and preset_name in capacity_model.image_presets:
        preset = capacity_model.image_presets[preset_name]
    else:
        # Tomamos el primer preset disponible como fallback razonable.
        presets = list(capacity_model.image_presets.values())
        if not presets:
            return "none", 0.0
        preset = presets[0]
    return preset.name, preset.cost()


def _evaluate_placement(
    note: "Note",
    column_index: int,
    start_mm: float,
    available_mm: float,
    capacity_model: "CapacityModel",
    settings: PlanSettings,
) -> Tuple[Assignment, float]:
    title_level = settings.default_title_level
    title_height_mm = capacity_model.title_cost(title_level, span=1) if title_level else 0.0
    img_mode, image_height_mm = _select_image_preset(note, capacity_model, settings)

    available_for_body = max(available_mm - title_height_mm - image_height_mm, 0.0)

    capacity_chars = capacity_model.capacity_per_column(
        available_mm, span=1, title_level=title_level, image_preset=None if img_mode == "none" else img_mode
    )
    body_chars_fit = min(int(capacity_chars), int(getattr(note, "chars_body", 0)))
    body_chars_overflow = max(int(getattr(note, "chars_body", 0)) - body_chars_fit, 0)

    body_height_mm = min(_body_height_from_chars(body_chars_fit, capacity_model), available_for_body)

    used_height_mm = max(title_height_mm + image_height_mm + body_height_mm, 0.0)
    if used_height_mm > available_mm:
        used_height_mm = available_mm

    target_total_mm = title_height_mm + image_height_mm + _body_height_from_chars(int(getattr(note, "chars_body", 0)), capacity_model)
    remaining_mm = max(available_mm - used_height_mm, 0.0)
    fit = body_chars_overflow == 0 and target_total_mm <= available_mm + 1e-6

    score_delta = settings.fit_bonus if fit else 0.0
    if not fit:
        score_delta -= settings.overflow_penalty_per_char * body_chars_overflow
        overflow_mm = max(target_total_mm - available_mm, 0.0)
        score_delta -= settings.overfill_penalty_per_mm * overflow_mm

    score_delta -= settings.gap_penalty_per_mm * remaining_mm

    assignment = Assignment(
        note=note,
        column_index=column_index,
        start_mm=start_mm,
        used_height_mm=used_height_mm,
        title_height_mm=title_height_mm,
        body_height_mm=body_height_mm,
        image_height_mm=image_height_mm,
        body_chars_fit=body_chars_fit,
        body_chars_overflow=body_chars_overflow,
        img_mode=img_mode,
        fit=fit,
        remaining_mm=remaining_mm,
    )

    return assignment, score_delta


def solve_page_layout(
    page_geom: "PageGeometry",
    notes: Sequence["Note"],
    capacity_model: "CapacityModel",
    settings: Optional[PlanSettings] = None,
) -> SolverOutcome:
    """Plan the layout of a page using a beam search."""

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
        next_states: List[_BeamState] = []
        for state in beam:
            for col_idx in range(columns):
                used_mm = state.columns_used_mm[col_idx]
                available_mm = max(column_height - used_mm, 0.0)
                if available_mm <= 0:
                    continue
                assignment, delta = _evaluate_placement(
                    note=note,
                    column_index=col_idx,
                    start_mm=used_mm,
                    available_mm=available_mm,
                    capacity_model=capacity_model,
                    settings=settings,
                )
                next_states.append(state.with_assignment(assignment, delta, column_height))

            next_states.append(state.with_drop(note, settings.drop_penalty))

        if not next_states:
            beam = [state.with_drop(note, settings.drop_penalty) for state in beam]
        else:
            next_states.sort(key=lambda entry: entry.score, reverse=True)
            beam = next_states[: settings.beam_width]

    best_state: Optional[_BeamState] = None
    best_score = float("-inf")
    for state in beam:
        gaps = tuple(max(column_height - used, 0.0) for used in state.columns_used_mm)
        gap_penalty = settings.final_gap_penalty_per_mm * sum(gaps)
        score = state.score - gap_penalty
        if score > best_score:
            best_state = state
            best_score = score
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
    """Generate a human-readable summary string for logging purposes."""

    placed = len(outcome.assignments)
    dropped = len(outcome.dropped_notes)
    gaps = ", ".join(f"{gap:.1f}mm" for gap in outcome.column_gaps_mm)
    return f"{placed} ubicadas · {dropped} pendientes · huecos: [{gaps}]"

