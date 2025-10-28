"""
Rules schema + loader for docx-formatter.

This module keeps the structure simple to match ops.py:
- Step.select is a plain dict[str, Any]
- Step.actions is a list[dict[str, Any]] and each action dict must have exactly one key
"""

from __future__ import annotations

import json
import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Literal

import yaml

logger = logging.getLogger(__name__)


@dataclass
class Safety:
    require_same_paragraph_count: bool = True
    require_same_bookmark_count: bool = True
    require_same_inline_shape_count: bool = True
    allow_text_changes: bool = False


@dataclass
class Step:
    name: str
    select: dict[str, Any]
    actions: list[dict[str, Any]] = field(default_factory=list)


@dataclass
class Rules:
    engine: Literal["auto", "word", "libre"] = "auto"
    safety: Safety = field(default_factory=Safety)
    steps: list[Step] = field(default_factory=list)

    @classmethod
    def from_dict(cls, raw: dict[str, Any]) -> Rules:
        """Create Rules instance from dictionary."""
        if not isinstance(raw, dict):
            raise ValueError("Top-level rules must be an object.")

        engine = raw.get("engine", "auto")
        if engine not in ("auto", "word", "libre"):
            raise ValueError("rules.engine must be one of: auto, word, libre")

        safety_raw = raw.get("safety", {}) or {}
        safety = Safety(
            require_same_paragraph_count=bool(safety_raw.get("require_same_paragraph_count", True)),
            require_same_bookmark_count=bool(safety_raw.get("require_same_bookmark_count", True)),
            require_same_inline_shape_count=bool(safety_raw.get("require_same_inline_shape_count", True)),
            allow_text_changes=bool(safety_raw.get("allow_text_changes", False)),
        )

        steps_raw = raw.get("steps")
        if not isinstance(steps_raw, list) or not steps_raw:
            raise ValueError("rules.steps must be a non-empty array.")

        steps: list[Step] = []
        for i, s in enumerate(steps_raw, start=1):
            if not isinstance(s, dict):
                raise ValueError(f"Step #{i} must be an object.")
            step = Step(
                name=s.get("name", f"Step #{i}"),
                select=s.get("select") or {},
                actions=s.get("actions") or [],
            )
            _validate_step(step)
            steps.append(step)

        rules = cls(engine=engine, safety=safety, steps=steps)
        logger.info("Loaded rules: engine=%s, steps=%d", rules.engine, len(rules.steps))
        return rules


def _validate_step(step: Step) -> None:
    """Validate a single step definition."""
    if not isinstance(step.name, str) or not step.name.strip():
        raise ValueError("Each step requires a non-empty 'name' string.")

    if not isinstance(step.select, dict) or not step.select:
        raise ValueError(f"Step '{step.name}': 'select' must be a non-empty object.")

    if not isinstance(step.actions, list) or not step.actions:
        raise ValueError(f"Step '{step.name}': 'actions' must be a non-empty array.")

    for idx, action in enumerate(step.actions, start=1):
        if not isinstance(action, dict) or not action:
            raise ValueError(f"Step '{step.name}': action #{idx} must be an object.")
        if len(action.keys()) != 1:
            raise ValueError(
                f"Step '{step.name}': action #{idx} must contain exactly one key (got {list(action.keys())})."
            )


def load_rules(path: str | Path) -> Rules:
    """Load YAML/JSON rules and return a validated Rules object (dict-friendly)."""
    p = Path(path)
    if not p.exists():
        raise ValueError(f"Rules file not found: {p}")

    try:
        if p.suffix.lower() in {".yml", ".yaml"}:
            raw = yaml.safe_load(p.read_text(encoding="utf-8"))
        elif p.suffix.lower() == ".json":
            raw = json.loads(p.read_text(encoding="utf-8"))
        else:
            raise ValueError("Rules file must be .yaml/.yml or .json")
    except Exception as e:
        raise ValueError(f"Failed to read rules file '{p}': {e}") from e

    return Rules.from_dict(raw)