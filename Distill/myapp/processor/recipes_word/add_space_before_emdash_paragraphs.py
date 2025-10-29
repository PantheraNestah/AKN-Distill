from __future__ import annotations
from typing import Any
from win32com.client import constants as C

def add_space_before_emdash_paragraphs_py(doc: Any) -> dict:
    """
    If a paragraph's second character is an em dash (—),
    append a space at the beginning of that paragraph.
    
    This ensures proper spacing for paragraphs that start with
    a number/letter followed immediately by an em dash.
    Example: "1—Text" becomes " 1—Text"
    
    Args:
        doc: Word document object
        
    Returns:
        dict: Result with count of paragraphs updated
    """
    try:
        app = doc.Application
        app.ScreenUpdating = False
        changed = 0
        errors = []
        
        # Store initial selection
        initial_start = app.Selection.Start
        initial_end = app.Selection.End
        
        # Process each paragraph in the document
        for para in doc.Paragraphs:
            try:
                txt = para.Range.Text
                
                # Ensure the paragraph has at least two characters
                if len(txt) >= 2:
                    # Check if second character is an em dash (Unicode 8212 / chr(8212))
                    # In Python: chr(8212) = '—'
                    if txt[1] == '\u2014' or txt[1] == '—':
                        # Check if paragraph doesn't already start with a space
                        if txt[0] != ' ':
                            # Insert a space at the beginning
                            para.Range.InsertBefore(" ")
                            changed += 1
                        
            except Exception as e:
                errors.append(f"Error processing paragraph at position {para.Range.Start}: {str(e)}")
        
        # Restore original selection
        app.Selection.SetRange(initial_start, initial_end)
        
        result = {
            "ok": True,
            "count_updated": changed,
            "description": f"Added leading space to {changed} paragraph(s) with em dash at position 2"
        }
        
        if errors:
            result["warnings"] = errors[:10]
            if len(errors) > 10:
                result["warnings"].append(f"... and {len(errors) - 10} more errors")
        
        return result
        
    except Exception as e:
        return {
            "ok": False,
            "error": str(e)
        }
    
    finally:
        app.ScreenUpdating = True
