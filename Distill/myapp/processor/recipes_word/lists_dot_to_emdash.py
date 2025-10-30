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
        total_pages = doc.ComputeStatistics(C.wdStatisticPages)
        original_page_end = page_end
        page_end = min(page_end, 4, total_pages)
        
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
                    rPage.End = doc.Content.End

                # --- Step 2: Convert numbering to literal text on this page only ---
                try:
                    rPage.ListFormat.ConvertNumbersToText()
                    errors.append(f"INFO: Called ConvertNumbersToText on page {pg}")
                except Exception as e:
                    errors.append(f"Warning: ConvertNumbersToText failed on page {pg}: {str(e)}")

                # --- Step 3: Process each paragraph on this page ---
                for para in rPage.Paragraphs:
                    try:
                        pre_text = para.Range.Text[:60].replace('\r', ' ').replace('\n', ' ')
                        errors.append(f"DEBUG: Pg{pg} Para start: '{pre_text}'")

                        # PYTHON COM FIX: Cannot use .Duplicate like VBA
                        # Must work directly on para.Range for changes to persist
                        
                        # --- 3a) Replace "digits." with "digits—" ---
                        find = para.Range.Find
                        find.ClearFormatting()
                        find.Replacement.ClearFormatting()
                        find.Text = "([0-9]{1,})."
                        find.Replacement.Text = r"\1—"
                        find.MatchWildcards = True
                        find.Forward = True
                        find.Wrap = C.wdFindStop  # Stay within paragraph
                        find.Format = False

                        try:
                            if find.Execute(Replace=C.wdReplaceOne):
                                changed += 1
                                post_text = para.Range.Text[:60].replace('\r', ' ').replace('\n', ' ')
                                errors.append(f"SUCCESS: Pg{pg} replaced: '{post_text}'")
                        except Exception as e:
                            errors.append(f"DEBUG: find failed on pg{pg}: {str(e)}")

                        # --- 3b) Remove space/tab immediately after em dash ---
                        find = para.Range.Find
                        find.ClearFormatting()
                        find.Replacement.ClearFormatting()
                        find.Text = "—([ ^t]{1,})"
                        find.Replacement.Text = "—"
                        find.MatchWildcards = True
                        find.Forward = True
                        find.Wrap = C.wdFindStop
                        find.Format = False
                        
                        try:
                            find.Execute(Replace=C.wdReplaceOne)
                        except Exception:
                            pass

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
            "description": f"Converted {changed} list numbers to em dashes on pages {page_start}-{page_end}"
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