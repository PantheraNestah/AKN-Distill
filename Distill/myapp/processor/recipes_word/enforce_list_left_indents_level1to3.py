from __future__ import annotations
from typing import Any
import re

def enforce_list_left_indents_level1to3_py(doc: Any) -> dict:
    """
    Attempts to set consistent left indents for list levels 1-3.
    For python-docx we perform best-effort paragraph-level indent adjustments.
    The COM implementation is kept as a fallback for template-level changes.
    """
    # python-docx path
    try:
        if hasattr(doc, "paragraphs"):
            changed = 0
            # Mapping level -> indent in inches
            indents = {1: 0.3, 2: 0.6, 3: 0.9}
            for para in doc.paragraphs:
                text = para.text or ""
                # crude detection: starts with digit + dot or digit+closeparen etc.
                m = re.match(r"^\s*(\d+\.|\(\d+\)|\([a-zA-Z]\))", text)
                if m:
                    # choose level by simple heuristics: length of marker
                    marker = m.group(1)
                    if re.match(r"^\d+\.", marker):
                        level = 1
                    elif re.match(r"^\(\d+\)", marker):
                        level = 2
                    else:
                        level = 3
                    try:
                        target_inch = indents.get(level, 0.3)
                        # python-docx expects length objects; use Pt via engines when applied
                        # Here we set numeric values; DocxEngine._parse_unit will convert if re-applied
                        # But paragraph.left_indent expects a Length object; set as a number of inches in points
                        from docx.shared import Inches
                        para.paragraph_format.left_indent = Inches(target_inch)
                        para.paragraph_format.first_line_indent = Inches(0)
                        changed += 1
                    except Exception:
                        pass

            return {"ok": True, "count_updated": changed, "description": f"Adjusted {changed} list paragraph(s) for Level 1-3 indents (docx)"}
    except Exception:
        pass

    # Fallback: original COM behavior
    try:
        from win32com.client import constants as C

        app = doc.Application
        app.ScreenUpdating = False
        changed = 0
        errors = []

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

                            target_indent = inches_to_points(level_num * 0.3)

                            needs_update = (
                                abs(lvl.NumberPosition - target_indent) > 0.1
                                or abs(lvl.TextPosition - target_indent) > 0.1
                                or lvl.TrailingCharacter != C.wdTrailingNone
                                or abs(para.Format.LeftIndent - target_indent) > 0.1
                            )

                            if needs_update:
                                lvl.NumberPosition = target_indent
                                lvl.TextPosition = target_indent
                                lvl.TrailingCharacter = C.wdTrailingNone
                                lvl.TabPosition = C.wdUndefined
                                para.Format.LeftIndent = target_indent
                                para.Format.FirstLineIndent = 0
                                para.Range.ListFormat.ApplyListTemplateWithLevel(
                                    ListTemplate=list_template,
                                    ContinuePreviousList=True,
                                    ApplyTo=C.wdListApplyToWholeList,
                                    ApplyLevel=level_num,
                                )
                                changed += 1

                        except Exception as e:
                            errors.append(f"Error processing list level at position {para.Range.Start}: {str(e)}")

            except Exception as e:
                errors.append(f"Error processing paragraph: {str(e)}")

        result = {"ok": True, "count_updated": changed, "description": f"Adjusted {changed} list paragraph(s) for Level 1-3 indents"}
        if errors:
            result["warnings"] = errors[:10]
            if len(errors) > 10:
                result["warnings"].append(f"... and {len(errors) - 10} more errors")
        return result

    finally:
        try:
            app.ScreenUpdating = True
        except Exception:
            pass
