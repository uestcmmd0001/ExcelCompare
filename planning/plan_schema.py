from __future__ import annotations

"""Schemas for automatic Excel comparison plans."""

from dataclasses import asdict, dataclass, field
from typing import Any


@dataclass
class SheetPlan:
    """One planned sheet-to-sheet comparison."""

    baseline_sheet: str
    revised_sheet: str
    header_row_baseline: int
    header_row_revised: int
    key_columns: list[str] = field(default_factory=list)
    column_mapping: dict[str, str] = field(default_factory=dict)
    ignore_columns: list[str] = field(default_factory=list)
    numeric_columns: dict[str, int] = field(default_factory=dict)
    date_columns: list[str] = field(default_factory=list)
    compare_formulas: bool = True
    confidence: float = 0.0
    reasons: list[str] = field(default_factory=list)
    uncertainties: list[str] = field(default_factory=list)


@dataclass
class PlanCandidate:
    """A candidate comparison strategy."""

    candidate_id: str
    strategy: str
    confidence: float
    sheet_pairs: list[SheetPlan] = field(default_factory=list)
    reasons: list[str] = field(default_factory=list)
    uncertainties: list[str] = field(default_factory=list)


@dataclass
class ComparisonPlan:
    """Top-level automatic plan bundle."""

    version: str
    planner: str
    primary_plan: PlanCandidate
    alternative_plans: list[PlanCandidate] = field(default_factory=list)
    uncertainties: list[str] = field(default_factory=list)
    source: str = "auto"


@dataclass
class SheetValidation:
    """Program validation result for one sheet plan."""

    baseline_sheet: str
    revised_sheet: str
    sheet_match_score: float
    header_score: float
    key_coverage_score: float
    key_uniqueness_score: float
    column_mapping_score: float
    type_consistency_score: float
    overall_confidence: float
    selected_key_columns: list[str] = field(default_factory=list)
    risks: list[str] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)


@dataclass
class PlanValidationResult:
    """Validated plan plus executable config."""

    overall_confidence: float
    sheet_validations: list[SheetValidation]
    risks: list[str]
    effective_config: dict[str, Any]


def to_dict(value: Any) -> Any:
    """Convert nested dataclasses to plain JSON/YAML serializable structures."""

    if hasattr(value, "__dataclass_fields__"):
        return asdict(value)
    if isinstance(value, list):
        return [to_dict(item) for item in value]
    if isinstance(value, dict):
        return {key: to_dict(item) for key, item in value.items()}
    return value

