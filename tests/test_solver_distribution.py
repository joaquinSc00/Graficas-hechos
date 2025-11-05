from dataclasses import dataclass, field
from pathlib import Path
from typing import Sequence
import sys

sys.path.append(str(Path(__file__).resolve().parents[1]))

from src.solver import PlanSettings, solve_page_layout


@dataclass
class DummyTextStyle:
    chars_per_line: float = 20
    lines_per_mm: float = 0.5

    def capacity_for_height(self, height: float) -> int:
        return int(self.chars_per_line * self.lines_per_mm * height)

    def title_cost(self, level: int | None = None, span: int = 1) -> float:
        return 0.0


@dataclass
class DummyCapacityModel:
    text_style: DummyTextStyle = field(default_factory=DummyTextStyle)
    image_presets: dict[str, object] = field(default_factory=dict)

    def title_cost(self, level: int | None) -> float:
        return 0.0


@dataclass
class DummyNote:
    note_id: str
    chars_body: int


@dataclass
class DummyPageGeometry:
    columns: int
    usable_rect_mm: Sequence[float]


def test_solver_uses_multiple_columns_for_short_notes() -> None:
    notes = [DummyNote(f"n{idx}", 50) for idx in range(1, 5)]
    page_geom = DummyPageGeometry(columns=2, usable_rect_mm=(0.0, 0.0, 0.0, 120.0))
    capacity_model = DummyCapacityModel()

    outcome = solve_page_layout(page_geom, notes, capacity_model, PlanSettings())

    used_columns = {assignment.column_index for assignment in outcome.assignments}
    assert used_columns == {0, 1}
    assert not outcome.dropped_notes
