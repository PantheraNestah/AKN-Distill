"""
Word automation recipe to remove spaces around em dashes.
"""

from .engines import C


def remove_spaces_around_em_dash_py(app, doc, page_start=1, page_end=999, **_):
    """
    Process a Word document to remove spaces around em dashes.
    
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

        # Get document content
        rng = doc.Content
        rng.Find.ClearFormatting()
        rng.Find.Replacement.ClearFormatting()

        # --- 1️⃣ Remove spaces around em dash (both sides) ---
        find = rng.Find
        find.Text = "([  ]{1,})—([  ]{1,})"  # Any spaces around em dash
        find.Replacement.Text = "—"
        find.MatchWildcards = True
        find.Forward = True
        find.Wrap = C.wdFindContinue
        find.Format = False

        if find.Execute(Replace=C.wdReplaceAll):
            changed += 1

        # --- 2️⃣ Remove spaces before only (e.g. " —word" → "—word") ---
        find.Text = "([  ]{1,})—"
        find.Replacement.Text = "—"
        find.MatchWildcards = True

        if find.Execute(Replace=C.wdReplaceAll):
            changed += 1

        # --- 3️⃣ Remove spaces after only (e.g. "word— " → "word—") ---
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
        return {
            "ok": False,
            "error": str(e)
        }

    finally:
        app.ScreenUpdating = True
