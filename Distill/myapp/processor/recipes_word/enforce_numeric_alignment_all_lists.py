"""
Word automation recipe to enforce numeric alignment in lists.
Provides a python-docx fallback that uses paragraph-level adjustments where possible.
"""

import re

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
    # python-docx fallback: do text-level adjustments and simple paragraph formatting
    try:
        if hasattr(doc, "paragraphs"):
            changed = 0
            num_re = re.compile(r"^(\s*)(\d+)\.\s+")
            for para in doc.paragraphs:
                text = para.text or ""
                m = num_re.match(text)
                if m:
                    # Ensure there's no extra space after the dot (collapse to single space)
                    new = m.group(1) + m.group(2) + ". " + text[m.end():]
                    if new != text:
                        para.text = new
                        changed += 1
            return {"ok": True, "count_updated": changed, "description": f"Adjusted numeric alignment in {changed} list paragraphs (docx)"}
    except Exception:
        pass

    # Fallback to COM behavior
    try:
        from win32com.client import constants as C

        app = doc.Application
        initial_start = app.Selection.Start
        initial_end = app.Selection.End
        app.ScreenUpdating = False
        errors = []
        changed = 0

        inches_to_points = app.InchesToPoints
        numPos = {
            1: inches_to_points(0.5),
            2: inches_to_points(0.8),
            3: inches_to_points(1.1),
        }
        gap = inches_to_points(0.1)

        for para in doc.Paragraphs:
            try:
                lf = para.Range.ListFormat
                if lf.ListType != C.wdListNoNumbering:
                    level_num = lf.ListLevelNumber
                    if 1 <= level_num <= 3:
                        try:
                            list_template = lf.ListTemplate
                            lvl = list_template.ListLevels(level_num)
                            needs_update = (
                                lvl.Alignment != C.wdListLevelAlignRight or
                                abs(lvl.NumberPosition - numPos[level_num]) > 0.1 or
                                abs(lvl.TextPosition - (numPos[level_num] + gap)) > 0.1
                            )
                            if needs_update:
                                lvl.Alignment = C.wdListLevelAlignRight
                                lvl.NumberPosition = numPos[level_num]
                                lvl.TextPosition = numPos[level_num] + gap
                                lvl.TrailingCharacter = C.wdTrailingTab if gap != 0 else C.wdTrailingNone
                                lvl.TabPosition = C.wdUndefined

                                para.Range.ListFormat.ApplyListTemplateWithLevel(
                                    ListTemplate=list_template,
                                    ContinuePreviousList=True,
                                    ApplyTo=C.wdListApplyToWholeList,
                                    ApplyLevel=level_num,
                                )
                                changed += 1
                        except Exception as e:
                            errors.append(f"Error processing list level {level_num}: {str(e)}")
            except Exception:
                pass

        app.Selection.SetRange(initial_start, initial_end)

        result = {"ok": True, "count_updated": changed, "description": f"Adjusted numeric alignment in {changed} list paragraphs"}
        if errors:
            result["warnings"] = errors
        return result

    except Exception as e:
        return {"ok": False, "error": str(e)}

    finally:
        try:
            app.ScreenUpdating = True
        except Exception:
            pass

