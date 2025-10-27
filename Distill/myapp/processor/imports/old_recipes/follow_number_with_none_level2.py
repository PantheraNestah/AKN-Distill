# formatter/recipes/word_follow_number_none.py
from __future__ import annotations
from typing import Any
from win32com.client import constants as C


def follow_number_with_none_level2_py(
    doc: Any,
    page_start: int,
    page_end: int,
) -> dict:
    """
    Sets "Follow number with" = Nothing (no tab/space) for numbered lists
    at level 2 within a given page range.

    Equivalent to the VBA macro:
      FollowNumberWithNothing_Level2_PageRange
    """

    if page_end < page_start:
        return {"ok": False, "reason": "End page must be >= start page"}

    app = doc.Application
    app.ScreenUpdating = False
    count = 0
    errors = []

    try:
        # --- Define Range for page span ---
        try:
            start_rng = doc.GoTo(What=C.wdGoToPage, Which=C.wdGoToAbsolute, Count=page_start)
            end_rng = doc.GoTo(What=C.wdGoToPage, Which=C.wdGoToAbsolute, Count=page_end + 1)
            rng = doc.Range(start_rng.Start, end_rng.Start)
        except Exception as e:
            return {"ok": False, "reason": f"Failed to set page range: {str(e)}"}

        # --- Loop through all paragraphs in this range ---
        for p in rng.Paragraphs:
            try:
                lf = p.Range.ListFormat
                # Check if paragraph has any list formatting
                if lf.ListType == C.wdListNoNumbering:
                    continue
                    
                # Check if it's level 2
                if lf.ListLevelNumber != 2:
                    continue

                lvl = lf.ListTemplate.ListLevels(2)
                
                # Store original values for verification
                old_trailing = lvl.TrailingCharacter
                old_tab_pos = lvl.TabPosition
                
                # Apply changes
                lvl.TrailingCharacter = C.wdTrailingNone
                lvl.TabPosition = 0
                lvl.TextPosition = lvl.NumberPosition

                # Reapply level 2 so change reflects immediately
                p.Range.ListFormat.ApplyListTemplateWithLevel(
                    ListTemplate=lf.ListTemplate,
                    ContinuePreviousList=True,
                    ApplyLevel=2,
                )
                
                # Verify changes were applied
                if (lvl.TrailingCharacter == C.wdTrailingNone and 
                    lvl.TabPosition == 0):
                    count += 1
                else:
                    errors.append(f"Failed to apply changes to paragraph at position {p.Range.Start}")
                    
            except Exception as e:
                errors.append(f"Error processing paragraph: {str(e)}")

        result = {
            "ok": True,
            "page_range": f"{page_start}-{page_end}",
            "count_updated": count,
        }
        
        if errors:
            result["warnings"] = errors
            
        return result

    finally:
        app.ScreenUpdating = True
