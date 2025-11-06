from __future__ import annotations
from typing import Any
import re


def enforce_structured_list_indents_with_styles_py(doc: Any) -> dict:
    """
    Detects hierarchical legal list levels (1., (1), (a)) and applies:
    - Consistent left indents per level
    - Corresponding paragraph styles ("List Level 1", "List Level 2", "List Level 3")
    Keeps text and formatting intact.
    """
    # First attempt python-docx style operation (best-effort)
    try:
        if hasattr(doc, "paragraphs"):
            changed = 0
            from docx.shared import Inches
            from docx.enum.style import WD_STYLE_TYPE

            def ensure_style(docx_doc, name, indent_inch):
                try:
                    style = docx_doc.styles[name]
                except Exception:
                    try:
                        style = docx_doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
                        style.paragraph_format.left_indent = Inches(indent_inch)
                        style.paragraph_format.first_line_indent = Inches(0)
                    except Exception:
                        return None
                return style

            style_level1 = ensure_style(doc, "List Level 1", 0.0)
            style_level2 = ensure_style(doc, "List Level 2", 0.3)
            style_level3 = ensure_style(doc, "List Level 3", 0.6)

            for para in doc.paragraphs:
                text = (para.text or "").strip()
                try:
                    if re.match(r"^\d+\.", text):
                        target_style = style_level1
                    elif re.match(r"^\(\d+\)", text):
                        target_style = style_level2
                    elif re.match(r"^\([a-z]\)", text, re.I):
                        target_style = style_level3
                    else:
                        continue

                    if target_style is not None:
                        try:
                            para.style = target_style
                            para.paragraph_format.left_indent = Inches(0 if target_style == style_level1 else (0.3 if target_style == style_level2 else 0.6))
                            para.paragraph_format.first_line_indent = Inches(0)
                            changed += 1
                        except Exception:
                            pass

                except Exception:
                    # ignore paragraph-level errors
                    pass

            return {"ok": True, "count_updated": changed, "description": f"Applied indentation and styles to {changed} paragraph(s) (docx)"}
    except Exception:
        pass

    # Fallback to Word COM implementation
    try:
        import pythoncom
        import win32com.client
        from win32com.client import constants as C

        pythoncom.CoInitialize()
        app = None
        changed = 0
        errors = []

        app = doc.Application
        app.ScreenUpdating = False
        inches_to_points = app.InchesToPoints

        def get_or_create_style(style_name: str, base_indent_inch: float):
            try:
                style = doc.Styles(style_name)
            except Exception:
                style = doc.Styles.Add(Name=style_name, Type=C.wdStyleTypeParagraph)
                style.ParagraphFormat.LeftIndent = inches_to_points(base_indent_inch)
                style.ParagraphFormat.FirstLineIndent = 0
            return style

        style_level1 = get_or_create_style("List Level 1", 0.0)
        style_level2 = get_or_create_style("List Level 2", 0.3)
        style_level3 = get_or_create_style("List Level 3", 0.6)

        for para in doc.Paragraphs:
            text = para.Range.Text.strip()
            try:
                if re.match(r"^\d+\.", text):
                    target_indent = inches_to_points(0.0)
                    target_style = style_level1
                elif re.match(r"^\(\d+\)", text):
                    target_indent = inches_to_points(0.3)
                    target_style = style_level2
                elif re.match(r"^\([a-z]\)", text, re.I):
                    target_indent = inches_to_points(0.6)
                    target_style = style_level3
                else:
                    continue

                fmt = para.Format
                needs_indent_update = (abs(fmt.LeftIndent - target_indent) > 0.1 or abs(fmt.FirstLineIndent) > 0.1)
                needs_style_update = para.Style.NameLocal != target_style.NameLocal

                if needs_indent_update or needs_style_update:
                    para.Style = target_style
                    fmt.LeftIndent = target_indent
                    fmt.FirstLineIndent = 0
                    changed += 1

            except Exception as e:
                errors.append(f"Error at paragraph starting '{text[:20]}...': {str(e)}")

        result = {"ok": True, "count_updated": changed, "description": (f"Applied indentation and styles to {changed} paragraph(s). Levels auto-detected: 1→'List Level 1', 2→'List Level 2', 3→'List Level 3'.")}
        if errors:
            result["warnings"] = errors[:10]
            if len(errors) > 10:
                result["warnings"].append(f"... and {len(errors) - 10} more errors")

        return result

    except Exception as e:
        return {"ok": False, "error": str(e)}

    finally:
        try:
            if app is not None:
                app.ScreenUpdating = True
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
