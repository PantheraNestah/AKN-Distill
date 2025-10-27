from __future__ import annotations
from typing import Any
from win32com.client import constants as C

def no_space_after_number_all_lists_fix_py(doc: Any) -> dict:
    """
    Removes space after number for all list levels in the document.
    Applies consistent formatting by setting TrailingCharacter to None,
    aligning TextPosition with NumberPosition, and removing tab stops.
    """
    app = doc.Application
    app.ScreenUpdating = False
    changed = 0
    errors = []

    try:
        for para in doc.Paragraphs:
            try:
                lf = para.Range.ListFormat
                if lf.ListType != C.wdListNoNumbering:
                    try:
                        lvl_number = lf.ListLevelNumber
                        lvl = lf.ListTemplate.ListLevels(lvl_number)

                        # Store original values for verification
                        old_trailing = lvl.TrailingCharacter
                        old_text_pos = lvl.TextPosition
                        
                        # Apply formatting changes
                        lvl.TrailingCharacter = C.wdTrailingNone
                        lvl.TextPosition = lvl.NumberPosition
                        lvl.TabPosition = C.wdUndefined

                        # Verify changes were applied
                        if (lvl.TrailingCharacter == C.wdTrailingNone and
                            lvl.TextPosition == lvl.NumberPosition):
                            changed += 1
                        else:
                            errors.append(f"Failed to apply changes to paragraph at position {para.Range.Start}")

                    except Exception as e:
                        errors.append(f"Error processing list level: {str(e)}")
            except Exception as e:
                errors.append(f"Error accessing paragraph: {str(e)}")

        result = {
            "ok": True,
            "count_updated": changed,
        }
        
        if errors:
            result["warnings"] = errors
            
        return result

    finally:
        app.ScreenUpdating = True