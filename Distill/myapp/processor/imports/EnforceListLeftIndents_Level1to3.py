from __future__ import annotations
from typing import Any
from win32com.client import constants as C

def enforce_list_left_indents_level1to3_py(doc: Any) -> dict:
    """
    Sets consistent left indents for list levels 1-3:
    - Level 1: 0.3 inches
    - Level 2: 0.6 inches
    - Level 3: 0.9 inches
    Also removes spaces after numbers and aligns text with numbers.
    """
    app = doc.Application
    app.ScreenUpdating = False
    changed = 0
    errors = []

    wdListNoNumbering = 0
    wdTrailingNone = 0
    wdUndefined = -9999999

    inches_to_points = word.InchesToPoints

    for para in doc.Paragraphs:
        try:
            lf = para.Range.ListFormat
            if lf.ListType != wdListNoNumbering:
                level_num = lf.ListLevelNumber

                if 1 <= level_num <= 3:
                    try:
                        lvl = lf.ListTemplate.ListLevels(level_num)

                        # Set NumberPosition and TextPosition depending on list level
                        if level_num == 1:
                            lvl.NumberPosition = inches_to_points(0.3)
                            lvl.TextPosition = inches_to_points(0.3)
                        elif level_num == 2:
                            lvl.NumberPosition = inches_to_points(0.6)
                            lvl.TextPosition = inches_to_points(0.6)
                        elif level_num == 3:
                            lvl.NumberPosition = inches_to_points(0.9)
                            lvl.TextPosition = inches_to_points(0.9)

                        # Remove tab/space after number
                        lvl.TrailingCharacter = wdTrailingNone
                        # Keep text and number aligned
                        lvl.TabPosition = wdUndefined

                        # Match paragraph indent to level depth
                        para.Format.LeftIndent = inches_to_points(level_num * 0.3)
                        para.Format.FirstLineIndent = 0

                        changed += 1

                    except Exception:
                        # Skip problematic list templates/levels
                        pass
        except Exception:
            # Skip paragraphs that fail ListFormat access
            pass

    word.Application.ActiveWindow.Application.Run(
        f'MsgBox "{changed} list paragraph(s) adjusted for Level 1–3 indents.", 64, "Indent Enforcement Complete"'
    )

if __name__ == "__main__":
    enforce_list_left_indents_level1to3()
