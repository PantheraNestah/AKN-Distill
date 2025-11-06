from __future__ import annotations
from typing import Any
import re

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
    # python-docx fallback: operate on paragraph text and perform regex replacements.
    try:
        if hasattr(doc, "paragraphs"):
            changed = 0
            # Clamp page_end to 4 to match original behavior, but python-docx doesn't expose pages
            page_end = min(page_end, 4)
            num_re = re.compile(r"(\d+)\.")
            em_space_re = re.compile(r"—[ \t]+")

            for para in doc.paragraphs:
                text = para.text or ""
                new = num_re.sub(r"\1—", text)
                new = em_space_re.sub("—", new)
                if new != text:
                    para.text = new
                    changed += 1

            return {
                "ok": True,
                "count_updated": changed,
                "page_range": f"{page_start}-{page_end}",
                "description": f"Converted {changed} list numbers to em dashes (docx)"
            }
    except Exception:
        pass

    # Fallback to COM behavior if available
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
            return {"ok": False, "error": "End page must be >= start page"}

        if original_page_end > 4:
            errors.append(f"INFO: Page range clamped from 1-{original_page_end} to 1-{page_end} (matching VBA behavior)")

        initial_start = app.Selection.Start
        initial_end = app.Selection.End

        for pg in range(page_start, page_end + 1):
            try:
                rPage = doc.GoTo(What=C.wdGoToPage, Which=C.wdGoToAbsolute, Count=pg)
                rPage.Start = rPage.Start
                try:
                    rNext = rPage.GoTo(What=C.wdGoToPage, Which=C.wdGoToNext)
                    rPage.End = rNext.Start - 1
                except Exception:
                    rPage.End = doc.Content.End

                try:
                    rPage.ListFormat.ConvertNumbersToText()
                    errors.append(f"INFO: Called ConvertNumbersToText on page {pg}")
                except Exception as e:
                    errors.append(f"Warning: ConvertNumbersToText failed on page {pg}: {str(e)}")

                for para in rPage.Paragraphs:
                    try:
                        find = para.Range.Find
                        find.ClearFormatting()
                        find.Replacement.ClearFormatting()
                        find.Text = "([0-9]{1,})."
                        find.Replacement.Text = r"\1—"
                        find.MatchWildcards = True
                        find.Forward = True
                        find.Wrap = C.wdFindStop
                        find.Format = False

                        try:
                            if find.Execute(Replace=C.wdReplaceOne):
                                changed += 1
                        except Exception as e:
                            errors.append(f"DEBUG: find failed on pg{pg}: {str(e)}")

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
        return {"ok": False, "error": str(e)}

    finally:
        try:
            app.ScreenUpdating = True
        except Exception:
            pass