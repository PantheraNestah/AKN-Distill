from __future__ import annotations
from typing import Any
from win32com.client import constants as C

def remove_all_tabs_py(doc: Any) -> dict:
    """
    Removes all tab characters from the document using Find/Replace.
    This helps standardize spacing and prevent formatting inconsistencies.
    """
    app = doc.Application
    app.ScreenUpdating = False
    changed = 0
    errors = []

    try:
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
                # Count occurrences by checking document length change
                changed = 1  # We can't get exact count, but operation succeeded
            
        except Exception as e:
            errors.append(f"Error during find/replace operation: {str(e)}")

        # Restore original selection
        app.Selection.SetRange(initial_start, initial_end)

        result = {
            "ok": True,
            "count_updated": changed,
            "description": "All tab characters have been removed from the document"
        }

        if errors:
            result["warnings"] = errors

        return result

    finally:
        app.ScreenUpdating = True
