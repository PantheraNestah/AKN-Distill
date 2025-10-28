from __future__ import annotations
from typing import Any
from win32com.client import constants as C

def lists_dot_to_emdash_py(doc: Any, page_start: int = 1, page_end: int = 4) -> dict:
    """
    Converts numbered list dots to em dashes and removes any following spaces.
    Example: "1." becomes "1—" (without space after dash)
    
    IMPORTANT: This rule mimics VBA behavior - ONLY processes pages 1-4!
    Even if page_end > 4 is passed, it will be clamped to 4 maximum.
    
    Args:
        doc: Word document object
        page_start: First page to process (default: 1)
        page_end: Last page to process (default: 4, max: 4)
    """
    try:
        app = doc.Application
        app.ScreenUpdating = False
        changed = 0
        errors = []
        
        # ENFORCE: Only pages 1-4 (matching VBA behavior)
        # Even if config passes page_end=999, we cap it at 4
        total_pages = doc.ComputeStatistics(C.wdStatisticPages)
        original_page_end = page_end
        page_end = min(page_end, 4, total_pages)  # ✅ HARD LIMIT: Max page 4
        
        # Validate page range
        if page_end < page_start:
            return {
                "ok": False,
                "error": "End page must be >= start page"
            }
        
        # Log if we clamped the page range
        if original_page_end > 4:
            errors.append(f"INFO: Page range clamped from 1-{original_page_end} to 1-{page_end} (matching VBA behavior)")

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
        if app:
            app.ScreenUpdating = True
