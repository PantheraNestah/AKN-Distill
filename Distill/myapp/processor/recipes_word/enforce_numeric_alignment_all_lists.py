"""
Word automation recipe to enforce numeric alignment in lists.
Ensures consistent spacing and alignment of numbers in multilevel lists.
"""

from .engines import C

def enforce_numeric_alignment_all_lists_py(app, doc, page_start=1, page_end=999, **_):
    """
    Process a Word document to enforce numeric alignment in lists.
    
    Args:
        app: Word application instance
        doc: Word document instance
        page_start: First page to process (default: 1)
        page_end: Last page to process (default: 999)
        
    Returns:
        dict: Results with ok/error status and count of changes
    """
    try:
        # Store initial selection
        initial_start = app.Selection.Start
        initial_end = app.Selection.End
        
        app.ScreenUpdating = False
        errors = []
        changed = 0

        # Define numeric alignment columns
        inches_to_points = app.InchesToPoints
        numPos = {
            1: inches_to_points(0.5),  # Level 1 number column
            2: inches_to_points(0.8),  # Level 2 number column
            3: inches_to_points(1.1),  # Level 3 number column
        }
        gap = inches_to_points(0.1)    # small space between number and text

        # Loop through all paragraphs
        for para in doc.Paragraphs:
            try:
                lf = para.Range.ListFormat
                if lf.ListType != C.wdListNoNumbering:
                    level_num = lf.ListLevelNumber

                    if 1 <= level_num <= 3:
                        try:
                            lvl = lf.ListTemplate.ListLevels(level_num)

                            # --- Key alignment rule ---
                            lvl.Alignment = C.wdListLevelAlignRight
                            lvl.NumberPosition = numPos[level_num]
                            lvl.TextPosition = numPos[level_num] + gap
                            lvl.TrailingCharacter = (
                                C.wdTrailingNone if gap == 0 else C.wdTrailingTab
                            )
                            lvl.TabPosition = C.wdUndefined

                            changed += 1
                        except Exception as e:
                            errors.append(f"Error processing list level {level_num}: {str(e)}")
            except Exception as e:
                # Skip non-list paragraphs silently
                pass

        # Restore original selection
        app.Selection.SetRange(initial_start, initial_end)

        result = {
            "ok": True,
            "count_updated": changed,
            "description": f"Adjusted numeric alignment in {changed} list paragraphs"
        }

        if errors:
            result["warnings"] = errors

        return result

    except Exception as e:
        return {
            "ok": False,
            "error": str(e)
        }

    finally:
        app.ScreenUpdating = True

