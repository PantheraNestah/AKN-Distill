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
                            list_template = lf.ListTemplate
                            lvl = list_template.ListLevels(level_num)

                            # Set NumberPosition and TextPosition depending on list level
                            target_indent = inches_to_points(level_num * 0.3)
                            
                            # Check if update needed
                            needs_update = (
                                abs(lvl.NumberPosition - target_indent) > 0.1 or
                                abs(lvl.TextPosition - target_indent) > 0.1 or
                                lvl.TrailingCharacter != C.wdTrailingNone or
                                abs(para.Format.LeftIndent - target_indent) > 0.1
                            )
                            
                            if needs_update:
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
                                    ListTemplate=list_template,
                                    ContinuePreviousList=True,
                                    ApplyTo=C.wdListApplyToWholeList,
                                    ApplyLevel=level_num
                                )
                                changed += 1

                        except Exception as e:
                            errors.append(f"Error processing list level at position {para.Range.Start}: {str(e)}")

            except Exception as e:
                errors.append(f"Error processing paragraph: {str(e)}")

        result = {
            "ok": True,
            "count_updated": changed,
            "description": f"Adjusted {changed} list paragraph(s) for Level 1-3 indents"
        }

        if errors:
            result["warnings"] = errors[:10]
            if len(errors) > 10:
                result["warnings"].append(f"... and {len(errors) - 10} more errors")

        return result

    finally:
        app.ScreenUpdating = True
