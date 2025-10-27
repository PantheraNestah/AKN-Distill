from __future__ import annotations
from typing import Any
from win32com.client import constants as C

def enforce_list_left_indents_level1to3_py(doc: Any) -> dict:
    """
    Sets consistent left indents for list levels 1-3:
    - Level 1: 0.3 inches
    - Level 2: 0.6 inches
    - Level 3: 0.9 inches
    Also removes spaces after numbers and aligns text with numbers.
    """
    app = doc.Application
    app.ScreenUpdating = False
    changed = 0
    errors = []

    try:
        inches_to_points = app.InchesToPoints

        for para in doc.Paragraphs:
            try:
                lf = para.Range.ListFormat
                if lf.ListType != C.wdListNoNumbering:
                    level_num = lf.ListLevelNumber

                    if 1 <= level_num <= 3:
                        try:
                            lvl = lf.ListTemplate.ListLevels(level_num)

                            # Store original values for verification
                            old_number_pos = lvl.NumberPosition
                            old_text_pos = lvl.TextPosition
                            old_left_indent = para.Format.LeftIndent

                            # Set NumberPosition and TextPosition depending on list level
                            target_indent = inches_to_points(level_num * 0.3)
                            lvl.NumberPosition = target_indent
                            lvl.TextPosition = target_indent

                            # Remove tab/space after number
                            lvl.TrailingCharacter = C.wdTrailingNone
                            # Keep text and number aligned
                            lvl.TabPosition = C.wdUndefined

                            # Match paragraph indent to level depth
                            para.Format.LeftIndent = target_indent
                            para.Format.FirstLineIndent = 0

                            # Reapply list formatting to ensure changes take effect
                            para.Range.ListFormat.ApplyListTemplateWithLevel(
                                ListTemplate=lf.ListTemplate,
                                ContinuePreviousList=True,
                                ApplyLevel=level_num
                            )

                            # Verify changes were applied
                            if (lvl.NumberPosition == target_indent and 
                                lvl.TextPosition == target_indent and 
                                para.Format.LeftIndent == target_indent):
                                changed += 1
                            else:
                                errors.append(f"Failed to apply indent changes to paragraph at position {para.Range.Start}")

                        except Exception as e:
                            errors.append(f"Error processing list level: {str(e)}")

            except Exception as e:
                errors.append(f"Error processing paragraph: {str(e)}")

        result = {
            "ok": True,
            "count_updated": changed,
            "description": f"Adjusted {changed} list paragraph(s) for Level 1-3 indents"
        }

        if errors:
            result["warnings"] = errors

        return result

    finally:
        app.ScreenUpdating = True
