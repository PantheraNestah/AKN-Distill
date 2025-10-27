"""
CLI entrypoint.
"""

from __future__ import annotations

import argparse
import logging
from pathlib import Path

from .pipeline import run, run_batch


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="format-docx",
        description="Precise DOCX formatter with Word COM PDF export.",
    )
    parser.add_argument("inputs", nargs="+", help="Input .docx file(s) or glob(s).")
    parser.add_argument("--rules", required=True, help="Path to rules YAML/JSON.")
    parser.add_argument("--engine", choices=["auto", "word", "libre"], default="auto", help="Engine selection (default: auto).")
    parser.add_argument("--out", default="output", help="Base output directory (default: output).")
    parser.add_argument("--audit", action="store_true", help="Write audit JSON per file.")
    parser.add_argument("--dry-run", action="store_true", help="No DOCX/PDF saved; only audit if --audit is set.")
    parser.add_argument("--verbose", action="store_true", help="Verbose logging.")

    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )

    base_out_dir = Path(args.out)
    base_out_dir.mkdir(parents=True, exist_ok=True)

    # If multiple inputs/globs, run batch
    is_glob = any(any(ch in s for ch in "*?[") for s in args.inputs)
    if len(args.inputs) > 1 or is_glob:
        code = run_batch(
            inputs=args.inputs,
            rules_path=Path(args.rules),
            engine_hint=args.engine,
            base_out_dir=base_out_dir,
            write_audit=bool(args.audit),
            dry_run=bool(args.dry_run),
            verbose=bool(args.verbose),
        )
    else:
        code = run(
            input_path=Path(args.inputs[0]),
            rules_path=Path(args.rules),
            engine_hint=args.engine,
            base_out_dir=base_out_dir,
            write_audit=bool(args.audit),
            dry_run=bool(args.dry_run),
            verbose=bool(args.verbose),
        )

    raise SystemExit(code)


if __name__ == "__main__":
    main()
