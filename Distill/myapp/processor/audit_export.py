"""
Audit functionality for document processing.
"""
from dataclasses import dataclass
from pathlib import Path
import json
import logging
from typing import Any

@dataclass
class Snapshot:
    paragraph_count: int
    bookmark_count: int
    inline_shape_count: int
    content_control_count: int
    tables_count: int
    headings_by_level: dict

    @classmethod
    def take(cls, engine: Any, doc: Any) -> 'Snapshot':
        """Take a snapshot of document metrics"""
        snap = engine.snapshot(doc)
        return cls(
            paragraph_count=snap['paragraph_count'],
            bookmark_count=snap['bookmark_count'],
            inline_shape_count=snap['inline_shape_count'],
            content_control_count=snap['content_control_count'],
            tables_count=snap['tables_count'],
            headings_by_level=snap['headings_by_level']
        )

def compare(pre: Snapshot, post: Snapshot, safety: Any, logger: logging.Logger) -> list[str]:
    """Compare pre/post snapshots and return warnings"""
    warnings = []
    
    if safety.require_same_paragraph_count and pre.paragraph_count != post.paragraph_count:
        warnings.append(
            f"Paragraph count changed: {pre.paragraph_count} -> {post.paragraph_count}"
        )
    
    if safety.require_same_bookmark_count and pre.bookmark_count != post.bookmark_count:
        warnings.append(
            f"Bookmark count changed: {pre.bookmark_count} -> {post.bookmark_count}"
        )
        
    if (safety.require_same_inline_shape_count and 
        pre.inline_shape_count != post.inline_shape_count):
        warnings.append(
            f"Inline shape count changed: {pre.inline_shape_count} -> {post.inline_shape_count}"
        )
        
    return warnings

def save_artifacts(
    engine: Any,
    doc: Any,
    input_path: Path,
    output_dir: Path,
    write_pdf: bool,
    logger: logging.Logger
) -> None:
    """Save DOCX and optionally PDF"""
    stem = input_path.stem
    if write_pdf:
        pdf_path = output_dir / f"{stem}.pdf"
        logger.info(f"Exporting PDF: {pdf_path}")
        engine.export_pdf(doc, pdf_path)

def write_audit_file(
    file_out_dir: Path,
    stem: str,
    pre: Snapshot,
    post: Snapshot,
    step_summary: dict,
    engine_name: str,
    dry_run: bool,
) -> None:
    """Write audit JSON file"""
    audit = {
        "engine": engine_name,
        "dry_run": dry_run,
        "pre": pre.__dict__,
        "post": post.__dict__,
        "summary": step_summary,
    }
    
    audit_path = file_out_dir / f"{stem}.audit.json"
    with open(audit_path, "w", encoding="utf-8") as f:
        json.dump(audit, f, indent=2)