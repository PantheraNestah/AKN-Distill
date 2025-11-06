"""
Word automation recipe to remove spaces around em dashes.
Provides a python-docx fallback that performs regex replaces on paragraph text.
"""

import re


def remove_spaces_around_em_dash_py(doc, page_start=1, page_end=999, **_):
    """
    Process a Word document to remove spaces around em dashes.
    
    Args:
        doc: Word document instance
        page_start: First page to process (default: 1)
        page_end: Last page to process (default: 999)
        
    Returns:
        dict: Results with ok/error status and count of changes
    """
    # python-docx path: do paragraph-level regex replacements
    try:
        if hasattr(doc, "paragraphs"):
            changed = 0
            pat_both = re.compile(r"\s+—\s+")
            pat_before = re.compile(r"\s+—")
            pat_after = re.compile(r"—\s+")

            for para in doc.paragraphs:
                text = para.text or ""
                new = pat_both.sub("—", text)
                new = pat_before.sub("—", new)
                new = pat_after.sub("—", new)
                if new != text:
                    para.text = new
                    changed += 1

            return {"ok": True, "count_updated": changed, "description": f"Removed spaces around {changed} em dashes (docx)"}
    except Exception:
        pass

    # Fallback: Word COM implementation
    try:
        from win32com.client import constants as C
        app = doc.Application

        # Store initial selection
        initial_start = app.Selection.Start
        initial_end = app.Selection.End

        app.ScreenUpdating = False
        errors = []
        changed = 0

        # Get document content
        rng = doc.Content
        rng.Find.ClearFormatting()
        rng.Find.Replacement.ClearFormatting()

        # --- 1️⃣ Remove spaces around em dash (both sides) ---
        find = rng.Find
        find.Text = "([  ]{1,})—([  ]{1,})"
        find.Replacement.Text = "—"
        find.MatchWildcards = True
        find.Forward = True
        find.Wrap = C.wdFindContinue
        find.Format = False

        if find.Execute(Replace=C.wdReplaceAll):
            changed += 1

        # --- 2️⃣ Remove spaces before only ---
        find.Text = "([  ]{1,})—"
        find.Replacement.Text = "—"
        find.MatchWildcards = True

        if find.Execute(Replace=C.wdReplaceAll):
            changed += 1

        # --- 3️⃣ Remove spaces after only ---
        find.Text = "—([  ]{1,})"
        find.Replacement.Text = "—"
        find.MatchWildcards = True

        if find.Execute(Replace=C.wdReplaceAll):
            changed += 1

        # Restore original selection
        app.Selection.SetRange(initial_start, initial_end)

        result = {
            "ok": True,
            "count_updated": changed,
            "description": f"Removed spaces around {changed} em dashes"
        }

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
