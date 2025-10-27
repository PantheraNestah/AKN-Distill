"""
Tighten Level 3 spacing and preserve AOS sections.
Detects headings like 1.2.3, 2.4.1 etc. and removes extra spacing
without touching AOS, Appendices, or Schedules.
"""

import re

def tighten_level3_spacing_py(doc, log=None, **params):
    """
    Word recipe: tighten_level3_spacing
    """
    import re

    detect_pattern = re.compile(params.get("detect_pattern", r"^(\d+\.\d+\.\d+)"))
    exclude_patterns = [re.compile(pat, re.I) for pat in params.get("exclude_patterns", [])]

    spacing_before_pt = params.get("spacing_before_pt", 0)
    spacing_after_pt = params.get("spacing_after_pt", 0)
    line_spacing_multiple = params.get("line_spacing_multiple", 1.0)
    keep_with_next = params.get("keep_with_next", False)
    paragraph_alignment = params.get("paragraph_alignment", "left")

    count = 0
    for para in doc.Paragraphs:
        text = para.Range.Text.strip()
        if not text:
            continue

        # Skip excluded patterns
        if any(pat.search(text) for pat in exclude_patterns):
            continue

        # Match level 3 (1.2.3 etc.)
        if detect_pattern.match(text):
            fmt = para.Format
            fmt.SpaceBefore = spacing_before_pt
            fmt.SpaceAfter = spacing_after_pt
            fmt.LineSpacingRule = 4  # wdLineSpaceMultiple
            fmt.LineSpacing = line_spacing_multiple * 12
            fmt.KeepWithNext = keep_with_next
            fmt.Alignment = {"left": 0, "center": 1, "right": 2}.get(paragraph_alignment.lower(), 0)
            count += 1

    if log:
        log.info(f"Tightened spacing for {count} level-3 paragraphs.")
    return count
