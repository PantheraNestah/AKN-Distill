import win32com.client

def enforce_numeric_alignment_all_lists():
    # Connect to Word
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True  # optional for debugging

    doc = word.ActiveDocument

    wdListNoNumbering = 0
    wdListLevelAlignRight = 2
    wdTrailingNone = 0
    wdTrailingTab = 1
    wdUndefined = -9999999

    inches_to_points = word.InchesToPoints

    changed = 0

    # Define numeric alignment columns (same as VBA)
    numPos = {
        1: inches_to_points(0.5),  # Level 1 number column
        2: inches_to_points(0.8),  # Level 2 number column
        3: inches_to_points(1.1),  # Level 3 number column
    }
    gap = inches_to_points(0.1)    # small space between number and text

    # Loop through all paragraphs
    for para in doc.Paragraphs:
        try:
            lf = para.Range.ListFormat
            if lf.ListType != wdListNoNumbering:
                level_num = lf.ListLevelNumber

                if 1 <= level_num <= 3:
                    try:
                        lvl = lf.ListTemplate.ListLevels(level_num)

                        # --- Key alignment rule ---
                        lvl.Alignment = wdListLevelAlignRight
                        lvl.NumberPosition = numPos[level_num]
                        lvl.TextPosition = numPos[level_num] + gap
                        lvl.TrailingCharacter = (
                            wdTrailingNone if gap == 0 else wdTrailingTab
                        )
                        lvl.TabPosition = wdUndefined

                        changed += 1
                    except Exception:
                        # Ignore bad list levels/templates
                        pass
        except Exception:
            # Skip non-list paragraphs
            pass

    # Completion message (same as VBA MsgBox)
    word.Application.ActiveWindow.Application.Run(
        f'MsgBox "{changed} list paragraph(s) adjusted for numeric alignment.", 64, "Done"'
    )

if __name__ == "__main__":
    enforce_numeric_alignment_all_lists()

