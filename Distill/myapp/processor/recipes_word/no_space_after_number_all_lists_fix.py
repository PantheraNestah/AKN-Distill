from __future__ import annotations
from typing import Any
from win32com.client import constants as C

def no_space_after_number_all_lists_fix_py(doc: Any) -> dict:
    """
    Removes space after number for all list levels in the document.
    Applies consistent formatting by setting TrailingCharacter to None,
    aligning TextPosition with NumberPosition, and removing tab stops.
    
    CRITICAL: Must reapply list template for changes to take effect!
    """
    app = doc.Application
    app.ScreenUpdating = False
    changed = 0
    errors = []

    try:
        for para in doc.Paragraphs:
            try:
                # Get list format for the paragraph
                lf = para.Range.ListFormat
                if lf.ListType != C.wdListNoNumbering:  # Skip if not a list
                    try:
                        # Get the list level details
                        lvl_number = lf.ListLevelNumber
                        list_template = lf.ListTemplate
                        lvl = list_template.ListLevels(lvl_number)

                        # Store original values to detect if changes needed
                        needs_update = (
                            lvl.TrailingCharacter != C.wdTrailingNone or
                            lvl.TextPosition != lvl.NumberPosition or
                            lvl.TabPosition != C.wdUndefined
                        )
                        
                        if needs_update:
                            # Apply formatting changes
                            lvl.TrailingCharacter = C.wdTrailingNone
                            lvl.TextPosition = lvl.NumberPosition
                            lvl.TabPosition = C.wdUndefined

                            # CRITICAL: Reapply list template to make changes stick
                            para.Range.ListFormat.ApplyListTemplateWithLevel(
                                ListTemplate=list_template,
                                ContinuePreviousList=True,
                                ApplyTo=C.wdListApplyToWholeList,
                                ApplyLevel=lvl_number
                            )
                            changed += 1

                    except Exception as e:
                        errors.append(f"Error processing list level at position {para.Range.Start}: {str(e)}")
            except Exception as e:
                # Skip non-list paragraphs silently
                pass

        # Prepare result with essential information
        result = {
            "ok": True,
            "count_updated": changed,
            "description": f"Updated {changed} list paragraphs (removed space after number)"
        }
        
        # Add warnings if any errors occurred (but only first 10 to avoid spam)
        if errors:
            result["warnings"] = errors[:10]
            if len(errors) > 10:
                result["warnings"].append(f"... and {len(errors) - 10} more errors")
            
        return result

    finally:
        # Always restore screen updating
        app.ScreenUpdating = True