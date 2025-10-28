from __future__ import annotations
from typing import Any
import win32com.client
from win32com.client import constants as C
import pythoncom

# Ensure constants are properly initialized
_ = win32com.client.gencache.EnsureDispatch("Word.Application")

def lists_dot_to_emdash_py(doc: Any, page_start: int = 1, page_end: int = 4) -> dict:
    """
    Converts numbered list dots to em dashes and removes any following spaces.
    Example: "1." becomes "1—" (without space after dash)
    
    Args:
        doc: Word document object
        page_start: First page to process (default: 1)
        page_end: Last page to process (default: 4)
    """
    # Initialize COM for this thread
    pythoncom.CoInitialize()
    app = None

    try:
        # Initialize Word application and prepare
        app = doc.Application
        app.ScreenUpdating = False
        changed = 0
        errors = []
        
        # Ensure we only process up to page 4 or document end, whichever is less
        total_pages = doc.ComputeStatistics(C.wdStatisticPages)
        page_end = min(page_end, total_pages)
        # Validate page range
        if page_end < page_start:
            return {
                "ok": False,
                "error": "End page must be >= start page"
            }

        # Store initial selection
        initial_start = app.Selection.Start
        initial_end = app.Selection.End

        # Loop through specified pages
        for pg in range(page_start, page_end + 1):
            try:
                # --- Step 1: Get range for this page ---
                rPage = doc.GoTo(What=C.wdGoToPage, Which=C.wdGoToAbsolute, Count=pg)
                rPage.Start = rPage.Start

                try:
                    rNext = rPage.GoTo(What=C.wdGoToPage, Which=C.wdGoToNext)
                    rPage.End = rNext.Start - 1
                except Exception:
                    # If there is no next page, go to end of document
                    rPage.End = doc.Content.End

                # --- Step 2: Convert numbering to literal text ---
                try:
                    rPage.ListFormat.ConvertNumbersToText()
                except Exception as e:
                    errors.append(f"Warning: Could not convert list numbers on page {pg}: {str(e)}")

                # --- Step 3: Process each paragraph on this page ---
                for para in rPage.Paragraphs:
                    try:
                        # Work on the first ~60 characters only
                        head = para.Range.Duplicate
                        head.End = head.Start + 60

                        # --- 3a) Replace "digits." with "digits—" ---
                        find = head.Find
                        find.ClearFormatting()
                        find.Replacement.ClearFormatting()
                        find.Text = "([0-9]{1,})."
                        find.Replacement.Text = r"\1—"
                        find.MatchWildcards = True
                        find.Forward = True
                        find.Wrap = C.wdFindStop
                        find.Format = False

                        if find.Execute(Replace=C.wdReplaceOne):
                            changed += 1

                        # --- 3b) Remove space/tab immediately after em dash ---
                        head = para.Range.Duplicate
                        head.End = head.Start + 60

                        find = head.Find
                        find.ClearFormatting()
                        find.Replacement.ClearFormatting()
                        find.Text = "—([ ^t]{1,})"   # any space(s) or tab(s) after em dash
                        find.Replacement.Text = "—"
                        find.MatchWildcards = True
                        find.Forward = True
                        find.Wrap = C.wdFindStop
                        find.Format = False
                        find.Execute(Replace=C.wdReplaceOne)

                    except Exception as e:
                        errors.append(f"Error processing paragraph on page {pg}: {str(e)}")

            except Exception as e:
                errors.append(f"Error processing page {pg}: {str(e)}")

        # Restore original selection
        app.Selection.SetRange(initial_start, initial_end)

        result = {
            "ok": True,
            "count_updated": changed,
            "page_range": f"{page_start}-{page_end}",
            "description": f"Converted {changed} list numbers to em dashes"
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
        if 'app' in locals():
            app.ScreenUpdating = True
        # Clean up COM
        pythoncom.CoUninitialize()
