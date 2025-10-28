"""
Operations module: maps selectors to actions using the engine.
Now supports dynamic recipe discovery for Word-based formatters.
"""

import logging
import importlib
import pkgutil
from typing import Any

from .engines import Engine, WordComEngine  # type: ignore
from .rules import Safety, Step

logger = logging.getLogger(__name__)


# --- Dynamic Recipe Discovery ------------------------------------------------

def discover_word_recipes() -> dict[str, Any]:
    """
    Dynamically import all *_py functions from formatter.recipes_word.* modules.
    """
    import os
    import os
    from . import recipes_word as recipes_pkg

    registry: dict[str, Any] = {}
    package_path = os.path.dirname(os.path.abspath(recipes_pkg.__file__))

    for _, mod_name, _ in pkgutil.iter_modules([package_path]):
        print("[DEBUG] Found module candidate:", mod_name)
        if mod_name.startswith("_"):
            continue
        try:
            module = importlib.import_module(f".recipes_word.{mod_name}", package="myapp.processor")
            for attr_name in dir(module):
                if attr_name.endswith("_py"):
                    fn = getattr(module, attr_name)
                    if callable(fn):
                        recipe_name = mod_name  # use file name, not attr_name
                        registry[recipe_name] = fn
                        logger.debug(f"Registered Word recipe: {recipe_name}")
        except Exception as e:
            logger.warning(f"Failed to import recipe {mod_name}: {e}")

    return registry


# Build registry once at import
RECIPE_REGISTRY = discover_word_recipes()
logger.info(f"Discovered Word recipes: {list(RECIPE_REGISTRY.keys())}")


# --- Core Engine Operations --------------------------------------------------

def apply_steps(
    engine: Engine, doc: Any, steps: list[Step], safety: Safety, log: logging.Logger
) -> dict[str, Any]:
    summary: dict[str, Any] = {"steps": [], "total_modifications": 0}

    for step_idx, step in enumerate(steps, start=1):
        log.info(f"Step {step_idx}/{len(steps)}: {step.name}")

        try:
            step_result = _apply_single_step(engine, doc, step, safety, log)
            summary["steps"].append(
                {"name": step.name, "modifications": step_result["modifications"]}
            )
            summary["total_modifications"] += step_result["modifications"]
        except Exception as e:
            log.error(f"Step '{step.name}' failed: {e}")
            raise RuntimeError(f"Step '{step.name}' failed: {e}") from e

    log.info(f"Completed {len(steps)} steps, {summary['total_modifications']} modifications")
    return summary


def _apply_single_step(
    engine: Engine, doc: Any, step: Step, safety: Safety, log: logging.Logger
) -> dict[str, Any]:
    modifications = 0
    ranges = _resolve_selector(engine, doc, step.select, log)
    log.debug(f"  Selector matched {len(ranges)} ranges")

    for action_idx, action_dict in enumerate(step.actions, start=1):
        action_type = list(action_dict.keys())[0]
        action_config = action_dict[action_type]
        log.debug(f"  Action {action_idx}: {action_type}")

        try:
            count = _apply_action(engine, doc, ranges, action_type, action_config, safety, log)
            modifications += count
        except Exception as e:
            log.error(f"  Action '{action_type}' failed: {e}")
            raise ValueError(f"Action '{action_type}' failed: {e}") from e

    return {"modifications": modifications}


def _resolve_selector(
    engine: Engine, doc: Any, selector: dict[str, Any], log: logging.Logger
) -> list[Any]:
    """Resolve selector to list of ranges."""
    if selector.get("document") is True:
        return [doc.Content]

    if "by_style" in selector:
        return engine.select_by_style(doc, selector["by_style"])
    if "by_regex" in selector:
        cfg = selector["by_regex"]
        return engine.select_by_regex(
            doc,
            pattern=cfg["pattern"],
            scope=cfg.get("scope", "paragraphs"),
            flags=cfg.get("flags", []),
            page_range=cfg.get("page_range"),
        )
    if "by_bookmark" in selector:
        return engine.select_by_bookmark(doc, selector["by_bookmark"])
    if "by_content_control" in selector:
        return engine.select_by_content_control(doc, selector["by_content_control"])
    if "by_table" in selector:
        return engine.select_by_table(doc, **selector["by_table"])
    if "by_range" in selector:
        return engine.select_by_range(doc, **selector["by_range"])

    raise ValueError(f"Unknown selector type: {selector}")


def _apply_action(
    engine: Engine,
    doc: Any,
    ranges: list[Any],
    action_type: str,
    config: dict[str, Any],
    safety: Safety,
    log: logging.Logger,
) -> int:
    """Apply a single action to one or more ranges."""
    text_changing_actions = {"find_replace", "bookmark_text", "content_control_text"}
    if action_type in text_changing_actions:
        allow_change = config.get("allow_text_change", safety.allow_text_changes)
        if not allow_change:
            raise RuntimeError(
                f"Action '{action_type}' requires text changes but allow_text_changes=False."
            )

    # === Standard Actions ===
    if action_type == "paragraph_format":
        return engine.apply_paragraph_format(ranges, config)
    if action_type == "style_apply":
        return engine.apply_style(ranges, config["name"])
    if action_type == "numbering":
        return engine.apply_numbering(ranges, config)
    if action_type == "headers_footers":
        engine.set_headers_footers(doc, config)
        return 1
    if action_type == "field_update":
        engine.update_fields_and_toc(
            doc,
            update_all=config.get("update_all_fields", False),
            update_toc=config.get("update_toc", False),
        )
        return 1
    if action_type == "find_replace":
        return engine.find_replace(
            doc,
            find=config["find"],
            replace=config["replace"],
            regex=config.get("regex", False),
            wildcards=config.get("wildcards", False),
            whole_word=config.get("whole_word", False),
            match_case=config.get("match_case", False),
        )
    if action_type == "page_setup":
        engine.apply_page_setup(doc, config)
        return 1
    if action_type == "section_breaks":
        engine.insert_section_break(
            doc,
            before_selector=config.get("insert_before_selector", False),
            break_type=config.get("type", "next_page"),
        )
        return 1
    if action_type == "bookmark_text":
        engine.replace_bookmark_text(doc, config["name"], config["replace_text"])
        return 1
    if action_type == "content_control_text":
        engine.replace_content_control_text(doc, config["title_or_tag"], config["replace_text"])
        return 1
    if action_type == "table_format":
        engine.format_table(doc, config)
        return 1
    if action_type == "insert_image":
        engine.insert_image(doc, config)
        return 1
    if action_type == "raw_word_com":
        engine.raw_word_com(doc, config["commands"])
        return len(config["commands"])

    # === Dynamic Word Recipe ===
    if action_type == "word_recipe":
        name = (config or {}).get("name")
        enabled = bool((config or {}).get("enabled", True))
        params = (config or {}).get("params", {}) or {}

        if not enabled:
            log.info("word_recipe '%s' disabled; skipping", name)
            return 0
        if not isinstance(engine, WordComEngine):
            raise RuntimeError(f"word_recipe '{name}' requires the WordComEngine")

        fn = RECIPE_REGISTRY.get(name)
        if not callable(fn):
            raise ValueError(f"Unknown word_recipe: {name}. Available: {list(RECIPE_REGISTRY)}")

        log.info("Running word_recipe: %s (params=%s)", name, params)
        result = fn(doc, **params)
        log.info("word_recipe '%s' result: %s", name, result)

        mods = 1 if (isinstance(result, dict) and result.get("ok")) else 0
        try:
            if isinstance(result, dict) and "items_touched" in result:
                mods = int(result["items_touched"])
        except Exception:
            pass
        return mods

    raise ValueError(f"Unknown action type: {action_type}")
