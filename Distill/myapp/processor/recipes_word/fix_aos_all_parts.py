from __future__ import annotations
from typing import Any
from win32com.client import constants as C
from ._utils import (
    PT_PER_CM,
    looks_like_manual_number,
    replace_in_para,
    remove_all_spaces_after_dash,
)


def fix_aos_all_parts_py(doc: Any, block_width_cm: float = 12.0, pages_span: int = 3) -> dict:
    """Centers AoS section and normalizes numbering (Word-based version)."""
    r = doc.Content.Duplicate
    f = r.Find
    f.ClearFormatting()
    f.Replacement.ClearFormatting()
    f.Text = "ARRANGEMENT OF SECTIONS"
    f.MatchWildcards = False
    f.MatchCase = False
    f.Wrap = C.wdFindStop
    if not f.Execute():
        return {"ok": False, "reason": "AOS heading not found"}

    aos_first_pg = r.Information(C.wdActiveEndPageNumber)
    aos_last_pg = aos_first_pg + int(pages_span)

    page_w = doc.PageSetup.PageWidth
    lm = doc.PageSetup.LeftMargin
    rm = doc.PageSetup.RightMargin
    text_w = page_w - lm - rm
    block_w = min(max(block_width_cm * PT_PER_CM, 1.0), text_w * 0.9)
    side = (text_w - block_w) / 2.0

    first_item = None
    for p in doc.Paragraphs:
        pg = p.Range.Information(C.wdActiveEndPageNumber)
        if aos_first_pg <= pg <= aos_last_pg:
            if p.Range.ListFormat.ListType != C.wdListNoNumbering or looks_like_manual_number(p):
                first_item = p
                break
    if first_item is None:
        return {"ok": False, "reason": "No list items on AOS pages"}

    if first_item.Range.ListFormat.ListType != C.wdListNoNumbering:
        lvl = first_item.Range.ListFormat.ListTemplate.ListLevels(1)
        lvl.NumberStyle = C.wdListNumberStyleArabic
        lvl.NumberFormat = "%1—"
        lvl.TrailingCharacter = C.wdTrailingNone
        lvl.Alignment = C.wdListLevelAlignLeft
        lvl.NumberPosition = 0
        lvl.TextPosition = 0
        lvl.TabPosition = C.wdUndefined

    touched = 0
    for p in doc.Paragraphs:
        pg = p.Range.Information(C.wdActiveEndPageNumber)
        if not (aos_first_pg <= pg <= aos_last_pg):
            continue

        if p.Range.ListFormat.ListType != C.wdListNoNumbering:
            try:
                if p.Range.ListFormat.ListLevelNumber != 1:
                    p.Range.ListFormat.ListLevelNumber = 1
            except Exception:
                pass
        elif looks_like_manual_number(p):
            replace_in_para(p, "([0-9]{1,}[A-Z]{1,3})[.)][ ^t]@", "\\1—")
            replace_in_para(p, "([0-9]{1,}[A-Z]{1,3})[ ^t]@", "\\1—")
            replace_in_para(p, "([0-9]{1,})[.)][ ^t]@", "\\1—")
            replace_in_para(p, "([0-9]{1,})[ ^t]@", "\\1—")
            remove_all_spaces_after_dash(p)

        pf = p.Range.ParagraphFormat
        pf.LeftIndent = side
        pf.RightIndent = side
        pf.FirstLineIndent = 0
        pf.Alignment = C.wdAlignParagraphLeft
        p.Range.ParagraphFormat.TabStops.ClearAll()
        pf.SpaceBefore = 0
        pf.SpaceAfter = 0
        touched += 1

    return {
        "ok": True,
        "aos_pages": f"{aos_first_pg}-{aos_last_pg}",
        "items_touched": touched,
        "center_width_pts": block_w,
        "side_indent_pts": side,
    }
