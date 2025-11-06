from __future__ import annotations
from typing import Any
import re

def no_space_after_number_all_lists_fix_py(doc: Any) -> dict:
    """
    Removes space after number for all list levels. For python-docx this is
    implemented as text-level adjustments (replacing "1. " with "1." or
    converting to em dash depending on original recipe semantics).
    The COM implementation attempts to change list templates; kept as fallback.
    """
    # python-docx path: simple regex replace in paragraphs
    try:
        if hasattr(doc, "paragraphs"):
            changed = 0
            # Replace digit + dot + space (e.g., "1. ") with digit + dot (remove space)
            pat = re.compile(r"(\d+\.)[ \t]+")
            for para in doc.paragraphs:
                text = para.text or ""
                new = pat.sub(r"\1", text)
                if new != text:
                    para.text = new
                    changed += 1

            return {"ok": True, "count_updated": changed, "description": f"Updated {changed} paragraphs (removed space after number) (docx)"}
    except Exception:
        pass

    # Fallback: COM implementation
    try:
        from win32com.client import constants as C

        app = doc.Application
        app.ScreenUpdating = False
        changed = 0
        errors = []

        for para in doc.Paragraphs:
            try:
                lf = para.Range.ListFormat
                if lf.ListType != C.wdListNoNumbering:
                    try:
                        lvl_number = lf.ListLevelNumber
                        list_template = lf.ListTemplate
                        lvl = list_template.ListLevels(lvl_number)

                        needs_update = (
                            lvl.TrailingCharacter != C.wdTrailingNone or
                            lvl.TextPosition != lvl.NumberPosition or
                            lvl.TabPosition != C.wdUndefined
                        )

                        if needs_update:
                            lvl.TrailingCharacter = C.wdTrailingNone
                            lvl.TextPosition = lvl.NumberPosition
                            lvl.TabPosition = C.wdUndefined

                            para.Range.ListFormat.ApplyListTemplateWithLevel(
                                ListTemplate=list_template,
                                ContinuePreviousList=True,
                                ApplyTo=C.wdListApplyToWholeList,
                                ApplyLevel=lvl_number,
                            )
                            changed += 1

                    except Exception as e:
                        errors.append(f"Error processing list level at position {para.Range.Start}: {str(e)}")
            except Exception:
                pass

        result = {"ok": True, "count_updated": changed, "description": f"Updated {changed} list paragraphs (removed space after number)"}
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