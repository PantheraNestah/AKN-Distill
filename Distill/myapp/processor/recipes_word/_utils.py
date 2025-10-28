from __future__ import annotations
from typing import Any
from win32com.client import constants as C

PT_PER_CM = 28.3464567


def looks_like_manual_number(para) -> bool:
    """Detect paragraphs that begin with manual numbering (e.g., '20A.', '20B)', etc.)"""
    t = (para.Range.Text or "").replace("\r", "").strip()
    if not t:
        return False
    i = 0
    n = len(t)
    while i < n and t[i].isdigit():
        i += 1
    if i == 0:
        return False
    j = i
    k = 0
    while j < n and t[j].isalpha() and k < 3:
        j += 1
        k += 1
    if j < n and t[j] in (".", ")"):
        j += 1
    return j >= n or t[j] in (" ", "\t")


def replace_in_para(para, find_pat: str, repl: str) -> None:
    """Wildcard-safe replace inside a paragraph."""
    r = para.Range.Duplicate
    f = r.Find
    f.ClearFormatting()
    f.Replacement.ClearFormatting()
    f.Text = find_pat
    f.Replacement.Text = repl
    f.MatchWildcards = True
    f.MatchCase = False
    f.Wrap = C.wdFindStop
    f.Format = False
    f.Execute(Replace=C.wdReplaceAll)


def remove_all_spaces_after_dash(para) -> None:
    """Remove spaces/tabs after em dash."""
    replace_in_para(para, "—[ ^t]@", "—")
