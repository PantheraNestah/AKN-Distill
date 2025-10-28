"""
Word automation recipe to enforce numeric alignment in lists.
Ensures consistent spacing and alignment of numbers in multilevel lists.
"""

from win32com.client import constants as C

def enforce_numeric_alignment_all_lists_py(doc, page_start=1, page_end=999, **_):
    """
    Process a Word document to enforce numeric alignment in lists.
    
    Args:
        doc: Word document instance
        page_start: First page to process (default: 1)
        page_end: Last page to process (default: 999)
        
    Returns:
        dict: Results with ok/error status and count of changes
    """
    try:
        # Get Word application instance
        app = doc.Application
        
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
                            list_template = lf.ListTemplate
                            lvl = list_template.ListLevels(level_num)

                            # Check if update needed
                            needs_update = (
                                lvl.Alignment != C.wdListLevelAlignRight or
                                abs(lvl.NumberPosition - numPos[level_num]) > 0.1 or
                                abs(lvl.TextPosition - (numPos[level_num] + gap)) > 0.1
                            )
                            
                            if needs_update:
                                # --- Key alignment rule ---
                                lvl.Alignment = C.wdListLevelAlignRight
                                lvl.NumberPosition = numPos[level_num]
                                lvl.TextPosition = numPos[level_num] + gap
                                lvl.TrailingCharacter = (
                                    C.wdTrailingNone if gap == 0 else C.wdTrailingTab
                                )
                                lvl.TabPosition = C.wdUndefined

                                # Reapply list template to ensure changes take effect
                                para.Range.ListFormat.ApplyListTemplateWithLevel(
                                    ListTemplate=list_template,
                                    ContinuePreviousList=True,
                                    ApplyTo=C.wdListApplyToWholeList,
                                    ApplyLevel=level_num
                                )
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

