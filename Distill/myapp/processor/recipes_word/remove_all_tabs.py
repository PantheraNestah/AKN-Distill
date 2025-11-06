from __future__ import annotations
from typing import Any
import re

def remove_all_tabs_py(doc: Any) -> dict:
    """
    Removes all tab characters from the document. Supports python-docx Documents
    (edits paragraph text) and falls back to Word COM Find/Replace where available.
    """
    # python-docx path
    try:
        if hasattr(doc, "paragraphs"):
            changed = 0
            for para in doc.paragraphs:
                if "\t" in (para.text or ""):
                    para.text = (para.text or "").replace("\t", "")
                    changed += 1

            return {
                "ok": True,
                "count_updated": changed,
                "description": "All tab characters have been removed from the document (docx)"
            }
    except Exception:
        pass

    # Fallback to Word COM implementation
    try:
        app = doc.Application
        app.ScreenUpdating = False
        changed = 0
        errors = []

        # Create a Range for the whole document
        rng = doc.Content
        find = rng.Find

        # Store initial selection to restore later
        initial_start = app.Selection.Start
        initial_end = app.Selection.End

        try:
            # Clear any existing find/replace formatting
            find.ClearFormatting()
            find.Replacement.ClearFormatting()

            # Configure find/replace parameters
            find.Text = "^t"               # tab character
            find.Replacement.Text = ""     # replace with nothing
            find.Forward = True
            find.Wrap = C.wdFindContinue
            find.Format = False
            find.MatchCase = False
            find.MatchWholeWord = False
            find.MatchWildcards = False

            # Execute find/replace
            replaced = find.Execute(
                Replace=C.wdReplaceAll
            )

            if replaced:
                changed = 1

        except Exception as e:
            errors.append(f"Error during find/replace operation: {str(e)}")

        # Restore original selection
        try:
            app.Selection.SetRange(initial_start, initial_end)
        except Exception:
            pass

        result = {
            "ok": True,
            "count_updated": changed,
            "description": "All tab characters have been removed from the document"
        }

        if errors:
            result["warnings"] = errors

        return result

    finally:
        try:
            app.ScreenUpdating = True
        except Exception:
            pass
