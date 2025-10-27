"""
Pipeline orchestrator: load rules → open doc → snapshot → apply → snapshot → compare → save → (optional) audit.
"""

from __future__ import annotations

import glob
import logging
from pathlib import Path
from typing import List

from . import __version__
from .audit_export import Snapshot, compare, save_artifacts, write_audit_file
from .engines import pick_engine, Engine
from .rules import load_rules
from . import ops


def run(
    input_path: Path,
    rules_path: Path,
    engine_hint: str,
    base_out_dir: Path,
    write_audit: bool,
    dry_run: bool,
    verbose: bool,
) -> int:
    """
    Run the formatting pipeline for a single file.
    Returns 0 on success, 1 on error.
    """
    logger = logging.getLogger("formatter")
    logger.setLevel(logging.DEBUG if verbose else logging.INFO)

    try:
        rules = load_rules(rules_path)
    except Exception as e:
        logger.error("Failed to load rules: %s", e)
        return 1

    # Engine choice: allow CLI override unless it's "auto"
    try:
        chosen = rules.engine if engine_hint == "auto" else engine_hint
        engine: Engine = pick_engine(chosen)
    except Exception as e:
        logger.error("Failed to initialize engine: %s", e)
        return 1

    logger.info("docx-formatter %s | engine=%s | input=%s", __version__, chosen, input_path)

    doc = None
    try:
        doc = engine.open_document(input_path)
        pre = Snapshot.take(engine, doc)

        step_summary: dict = {}
        if dry_run:
            logger.info("Dry-run mode: selectors/actions will NOT be applied; no DOCX/PDF saved.")
        else:
            # Apply steps (uses your ops.py behavior)
            step_summary = ops.apply_steps(engine, doc, rules.steps, rules.safety, logger)

        post = Snapshot.take(engine, doc)
        warnings = compare(pre, post, rules.safety, logger)
        for w in warnings:
            logger.warning("Warning: %s", w)

        # Create output folder and save artifacts unless dry-run
        file_out_dir = (base_out_dir / input_path.stem)
        file_out_dir.mkdir(parents=True, exist_ok=True)

        if not dry_run:
            # Always save PDF (Word exporter if available; engine handles it)
            save_artifacts(engine, doc, input_path, base_out_dir, write_pdf=True, logger=logger)

        if write_audit:
            write_audit_file(
                file_out_dir=file_out_dir,
                stem=input_path.stem,
                pre=pre,
                post=post,
                step_summary=step_summary,
                engine_name=type(engine).__name__,
                dry_run=dry_run,
            )

        return 0

    except Exception as e:
        logger.exception("Run failed: %s", e)
        return 1
    finally:
        try:
            if doc is not None:
                engine.close_document(doc)
        finally:
            try:
                engine.shutdown()
            except Exception:
                pass


def run_batch(
    inputs: List[str | Path],
    rules_path: Path,
    engine_hint: str,
    base_out_dir: Path,
    write_audit: bool,
    dry_run: bool,
    verbose: bool,
) -> int:
    """
    Expand globs and run pipeline for each item. Aggregate exit codes.
    """
    paths: list[Path] = []
    for item in inputs:
        s = str(item)
        if any(ch in s for ch in ["*", "?", "["]):
            paths.extend(Path(p) for p in glob.glob(s))
        else:
            paths.append(Path(item))

    if not paths:
        logging.getLogger("formatter").error("No inputs matched.")
        return 1

    overall = 0
    for p in paths:
        code = run(
            input_path=p,
            rules_path=rules_path,
            engine_hint=engine_hint,
            base_out_dir=base_out_dir,
            write_audit=write_audit,
            dry_run=dry_run,
            verbose=verbose,
        )
        if code != 0:
            overall = code
    return overall
