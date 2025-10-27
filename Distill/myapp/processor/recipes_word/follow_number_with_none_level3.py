# formatter/recipes_word/word_follow_number_none3.py
from __future__ import annotations
from typing import Any
from win32com.client import constants as C


def follow_number_with_none_level3_py(
    doc: any,
    page_start: int = 1,
    page_end: int = 9999,
    **params
) -> dict:
    """
    Bulletproof Level 3 list fixer.
    Accepts any extra params gracefully (for YAML flexibility).
    Only affects *true* level 3 numbered lists within given page range.
    """

    import re
    from win32com.client import constants as C

    if page_end < page_start:
        return {"ok": False, "reason": "End page must be >= start page"}

    app = doc.Application
    app.ScreenUpdating = False
    count = 0

    try:
        start_rng = doc.GoTo(What=C.wdGoToPage, Which=C.wdGoToAbsolute, Count=page_start)
        end_rng = doc.GoTo(What=C.wdGoToPage, Which=C.wdGoToAbsolute, Count=page_end + 1)
        rng = doc.Range(start_rng.Start, end_rng.Start)

        exclude_patterns = [
            re.compile(pat, re.I)
            for pat in params.get("exclude_patterns", ["AOS", "Appendix", "Schedule"])
        ]

        for p in rng.Paragraphs:
            try:
                lf = p.Range.ListFormat

                # --- Skip if no numbering or not in a list ---
                if not lf or lf.ListType in (C.wdListNoNumbering, 0):
                    continue

                # --- Skip if not level 3 ---
                if lf.ListLevelNumber != 3:
                    continue

                # --- Ensure valid list template ---
                lt = lf.ListTemplate
                if not lt or lt.ListLevels.Count < 3:
                    continue

                # --- Exclude AOS, Appendix, etc. ---
                text = p.Range.Text.strip()
                if any(pat.search(text) for pat in exclude_patterns):
                    continue

                # --- Apply spacing tweaks if requested ---
                fmt = p.Format
                fmt.SpaceBefore = params.get("spacing_before_pt", 0)
                fmt.SpaceAfter = params.get("spacing_after_pt", 0)
                fmt.LineSpacingRule = 4  # wdLineSpaceMultiple
                fmt.LineSpacing = params.get("line_spacing_multiple", 1.0) * 12
                fmt.KeepWithNext = params.get("keep_with_next", False)
                fmt.Alignment = {
                    "left": 0, "center": 1, "right": 2
                }.get(params.get("paragraph_alignment", "left").lower(), 0)

                # --- Perform numbering fix ---
                lvl = lt.ListLevels(3)
                lvl.TrailingCharacter = C.wdTrailingNone
                lvl.TabPosition = 0
                lvl.TextPosition = lvl.NumberPosition
                p.Range.ListFormat.ApplyListTemplateWithLevel(
                    ListTemplate=lt,
                    ContinuePreviousList=True,
                    ApplyLevel=3,
                )

                count += 1

            except Exception:
                continue

        return {
            "ok": True,
            "page_range": f"{page_start}-{page_end}",
            "count_updated": count,
        }

    finally:
        app.ScreenUpdating = True
