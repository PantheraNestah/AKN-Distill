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

                # --- Step 2: Convert numbering to literal text on this page only ---
                # Try the VBA-style range conversion first; record page sample for debugging
                try:
                    page_text_sample = rPage.Text[:200].replace('\r', ' ').replace('\n', ' ')
                    errors.append(f"DEBUG: Page {pg} sample before convert: '{page_text_sample}...'")
                except Exception:
                    page_text_sample = ''

                try:
                    # VBA calls ConvertNumbersToText on the page range — try that first
                    try:
                        rPage.ListFormat.ConvertNumbersToText()
                        errors.append(f"INFO: Called ConvertNumbersToText on page {pg}")
                    except Exception as e:
                        errors.append(f"Warning: Page-level ConvertNumbersToText failed on page {pg}: {str(e)}")

                    # After attempting page-level conversion, also attempt per-paragraph conversion
                    conv_count = 0
                    for para in rPage.Paragraphs:
                        try:
                            if para.Range.ListFormat.ListType != C.wdListNoNumbering:
                                # Convert per-paragraph if possible
                                try:
                                    para.Range.ListFormat.ConvertNumbersToText()
                                    conv_count += 1
                                except Exception:
                                    # ignore paragraphs that cannot be converted
                                    pass
                        except Exception:
                            pass

                    if conv_count:
                        errors.append(f"INFO: Converted {conv_count} paragraph list items to text on page {pg}")
                    else:
                        errors.append(f"INFO: No paragraph-level automatic numbering found on page {pg}")
                except Exception as e:
                    errors.append(f"Warning: Could not convert list numbers on page {pg}: {str(e)}")

                # --- Step 3: Process each paragraph on this page ---
                for para in rPage.Paragraphs:
                    try:
                        # Record paragraph-level debug info
                        try:
                            list_type = para.Range.ListFormat.ListType
                            level_num = para.Range.ListFormat.ListLevelNumber
                        except Exception:
                            list_type = None
                            level_num = None

                        pre_text = para.Range.Text[:120].replace('\r', ' ').replace('\n', ' ')
                        errors.append(f"DEBUG: Pg{pg} Para start (level={level_num}, type={list_type}): '{pre_text}...'")

                        # --- 3a) Replace "digits." with "digits—" ---
                        # CRITICAL: Work directly on para.Range, NOT a duplicate!
                        # Duplicates don't persist changes to the document
                        find = para.Range.Find
                        find.ClearFormatting()
                        find.Replacement.ClearFormatting()
                        find.Text = "([0-9]{1,})."
                        find.Replacement.Text = r"\1—"
                        find.MatchWildcards = True
                        find.Forward = True
                        find.Wrap = C.wdFindStop  # Don't wrap, stay in paragraph
                        find.Format = False

                        replaced = False
                        try:
                            if find.Execute(Replace=C.wdReplaceOne):
                                changed += 1
                                replaced = True
                        except Exception as e:
                            errors.append(f"DEBUG: find failed on pg{pg} para at start: {str(e)}")

                        # Log post-find snippet if replaced
                        if replaced:
                            post_text = para.Range.Text[:120].replace('\r', ' ').replace('\n', ' ')
                            errors.append(f"DEBUG: Pg{pg} Para AFTER replace: '{post_text}...'")

                        # --- FALLBACK: Catch plain-text numbered headings (e.g., "1. Section") ---
                        # If no replacement happened yet and paragraph starts with digits followed by dot,
                        # replace that dot with em-dash (pages 1-4 only; very conservative pattern)
                        if not replaced:
                            # Check if paragraph starts with optional whitespace + digits + dot
                            para_start = para.Range.Text[:20].strip()
                            if para_start and para_start[0].isdigit():
                                # Look for pattern like "1." or "2." at start
                                import re
                                match = re.match(r'^(\d+)\.\s*', para_start)
                                if match:
                                    # Found plain-text number pattern at paragraph start
                                    # CRITICAL: Work directly on para.Range, NOT a duplicate!
                                    find = para.Range.Find
                                    find.ClearFormatting()
                                    find.Replacement.ClearFormatting()
                                    # Match digits followed by dot at very start (with optional leading space)
                                    find.Text = "^[ ]{0,}([0-9]{1,})."
                                    find.Replacement.Text = r"\1—"
                                    find.MatchWildcards = True
                                    find.Forward = True
                                    find.Wrap = C.wdFindStop
                                    find.Format = False
                                    
                                    try:
                                        if find.Execute(Replace=C.wdReplaceOne):
                                            changed += 1
                                            replaced = True
                                            fallback_post = para.Range.Text[:120].replace('\r', ' ').replace('\n', ' ')
                                            errors.append(f"FALLBACK: Pg{pg} converted plain-text heading: '{fallback_post}...'")
                                    except Exception as e:
                                        errors.append(f"DEBUG: fallback find failed on pg{pg}: {str(e)}")

                        # --- 3b) Remove space/tab immediately after em dash ---
                        find = para.Range.Find
                        find.ClearFormatting()
                        find.Replacement.ClearFormatting()
                        find.Text = "—([ ^t]{1,})"   # any space(s) or tab(s) after em dash
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
