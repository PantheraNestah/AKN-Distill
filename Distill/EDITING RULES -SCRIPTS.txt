'(1)===============================================================
' Fix: Set "Follow number with: Nothing" for all list levels
' Works reliably by looping through paragraphs.
'===============================================================
Sub NoSpaceAfterNumber_AllLists_Fix()
    Dim p As Paragraph
    Dim lvl As ListLevel
    Dim changed As Long

    For Each p In ActiveDocument.Paragraphs
        If p.Range.ListFormat.ListType <> wdListNoNumbering Then
            On Error Resume Next
            Set lvl = p.Range.ListFormat.ListTemplate.ListLevels(p.Range.ListFormat.ListLevelNumber)
            With lvl
                ' Remove space or tab after the number
                .TrailingCharacter = wdTrailingNone
                ' Align text with number (no gap)
                .TextPosition = .NumberPosition
                .TabPosition = wdUndefined
            End With
            On Error GoTo 0
            changed = changed + 1
        End If
    Next p

    MsgBox changed & " list paragraphs updated (no space after number).", _
           vbInformation, "Done"
End Sub


======================================================================

'(2)====================================================================
' Enforce Proper LeftIndent for Level 1–3 Lists
'====================================================================
Sub EnforceListLeftIndents_Level1to3()
    Dim p As Paragraph
    Dim lvl As ListLevel
    Dim levelNum As Long
    Dim changed As Long
    
    ' Loop through each paragraph in the document
    For Each p In ActiveDocument.Paragraphs
        If p.Range.ListFormat.ListType <> wdListNoNumbering Then
            levelNum = p.Range.ListFormat.ListLevelNumber
            
            ' Only apply to level 1–3
            If levelNum >= 1 And levelNum <= 3 Then
                On Error Resume Next
                Set lvl = p.Range.ListFormat.ListTemplate.ListLevels(levelNum)
                With lvl
                    Select Case levelNum
                        Case 1
                            .NumberPosition = InchesToPoints(0.3)
                            .TextPosition = InchesToPoints(0.3)
                        Case 2
                            .NumberPosition = InchesToPoints(0.6)
                            .TextPosition = InchesToPoints(0.6)
                        Case 3
                            .NumberPosition = InchesToPoints(0.9)
                            .TextPosition = InchesToPoints(0.9)
                    End Select
                    ' Remove tab/space after number
                    .TrailingCharacter = wdTrailingNone
                    ' Keep text and number aligned
                    .TabPosition = wdUndefined
                End With
                On Error GoTo 0
                
                ' Optional: also ensure paragraph indent matches
                With p.Range.ParagraphFormat
                    .LeftIndent = InchesToPoints(levelNum * 0.3)
                    .FirstLineIndent = 0
                End With
                
                changed = changed + 1
            End If
        End If
    Next p
    
    MsgBox changed & " list paragraph(s) adjusted for Level 1–3 indents.", _
           vbInformation, "Indent Enforcement Complete"
End Sub


======================

'(3)===============================================================
' Remove all tab characters from the entire active document
'===============================================================
Sub RemoveAllTabs()
    Dim rng As Range
    Dim replaced As Long

    Set rng = ActiveDocument.Content
    rng.Find.ClearFormatting
    rng.Find.Replacement.ClearFormatting

    With rng.Find
        .Text = "^t"                  ' Tab character
        .Replacement.Text = ""        ' Remove it (replace with nothing)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
    End With

    replaced = rng.Find.Execute(Replace:=wdReplaceAll)

    MsgBox "All tabs have been removed from the document.", _
           vbInformation, "Done"
End Sub

======================================================
'(4)=====================================================================
' Remove any space before or after em dash (—) across the document
'=====================================================================
Sub RemoveSpacesAroundEmDash()
    Dim rng As Range
    Dim totalReplacements As Long
    
    Set rng = ActiveDocument.Content
    rng.Find.ClearFormatting
    rng.Find.Replacement.ClearFormatting
    
    ' Remove space before or after em dash
    With rng.Find
        .Text = "([  ]{1,})—([  ]{1,})"     ' match space(s) around em dash
        .Replacement.Text = "—"              ' replace with tight em dash
        .MatchWildcards = True
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
    End With
    totalReplacements = rng.Find.Execute(Replace:=wdReplaceAll)
    
    ' Handle cases: space before only ( " —word" → "—word" )
    With rng.Find
        .Text = "([  ]{1,})—"
        .Replacement.Text = "—"
        .MatchWildcards = True
    End With
    totalReplacements = totalReplacements + rng.Find.Execute(Replace:=wdReplaceAll)
    
    ' Handle cases: space after only ( "word— " → "word—" )
    With rng.Find
        .Text = "—([  ]{1,})"
        .Replacement.Text = "—"
        .MatchWildcards = True
    End With
    totalReplacements = totalReplacements + rng.Find.Execute(Replace:=wdReplaceAll)
    
    MsgBox "Spaces around em dashes cleaned (" & totalReplacements & " replacements).", _
           vbInformation, "Done"
End Sub
======================================================

'(5)=====================================================================
' Enforce right-aligned numbering for all numbered lists (Levels 1–3)
' Makes ones, tens, hundreds align vertically
'=====================================================================
Sub EnforceNumericAlignment_AllLists()
    Dim p As Paragraph
    Dim lvl As ListLevel
    Dim levelNum As Long
    Dim changed As Long

    ' Define numeric alignment columns (in inches)
    Dim numPos(1 To 3) As Single
    Dim textPos(1 To 3) As Single
    Dim gap As Single

    ' Adjust these values to fine-tune alignment
    numPos(1) = InchesToPoints(0.5)   ' Level 1 number column
    numPos(2) = InchesToPoints(0.8)   ' Level 2 number column
    numPos(3) = InchesToPoints(1.1)   ' Level 3 number column
    gap = InchesToPoints(0.1)         ' small space between number and text

    ' Loop through all paragraphs in the document
    For Each p In ActiveDocument.Paragraphs
        If p.Range.ListFormat.ListType <> wdListNoNumbering Then
            levelNum = p.Range.ListFormat.ListLevelNumber

            If levelNum >= 1 And levelNum <= 3 Then
                On Error Resume Next
                Set lvl = p.Range.ListFormat.ListTemplate.ListLevels(levelNum)

                With lvl
                    ' Key alignment rule
                    .Alignment = wdListLevelAlignRight
                    .NumberPosition = numPos(levelNum)
                    .TextPosition = numPos(levelNum) + gap
                    .TrailingCharacter = IIf(gap = 0, wdTrailingNone, wdTrailingTab)
                    .TabPosition = wdUndefined
                End With
                On Error GoTo 0
                changed = changed + 1
            End If
        End If
    Next p

    MsgBox changed & " list paragraph(s) adjusted for numeric alignment.", _
           vbInformation, "Done"
End Sub
========================================================================



'(6)=====================================================================
' Pages 1–4:
'   - Freeze list numbering to text (page-scoped)
'   - Replace the period after the number with an em dash (—)
'   - Remove any tab/space immediately after the em dash
'=====================================================================
Sub ListsDotToEmDash_Pages1to4_NoGap()
    Dim pg As Long
    Dim rPage As Range, rNext As Range
    Dim para As Paragraph
    Dim head As Range
    Dim did As Long

    For pg = 1 To 4
        '--- Get range for this page ---
        Set rPage = ActiveDocument.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=pg)
        rPage.Start = rPage.Start
        On Error Resume Next
        Set rNext = rPage.GoTo(What:=wdGoToPage, Which:=wdGoToNext)
        On Error GoTo 0
        If Not rNext Is Nothing Then
            rPage.End = rNext.Start - 1
        Else
            rPage.End = ActiveDocument.Content.End
        End If

        '--- Step 1: Convert numbering to literal text on this page only ---
        rPage.ListFormat.ConvertNumbersToText

        '--- Step 2: For each paragraph, fix the leading number marker ---
        For Each para In rPage.Paragraphs
            'Work on the first ~60 chars only (start of paragraph)
            Set head = para.Range.Duplicate
            head.End = head.Start + 60

            ' 2a) Replace "digits." with "digits—" (first occurrence near start)
            With head.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = "([0-9]{1,})."
                .Replacement.Text = "\1—"
                .MatchWildcards = True
                .Forward = True
                .Wrap = wdFindStop
                .Format = False
            End With
            If head.Find.Execute(Replace:=wdReplaceOne) Then
                did = did + 1
            End If

            ' 2b) Remove any TAB/space immediately after the em dash at start
            Set head = para.Range.Duplicate
            head.End = head.Start + 60
            With head.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = "—([ ^t]{1,})"   ' any space(s) or tab(s) after em dash
                .Replacement.Text = "—"
                .MatchWildcards = True
                .Forward = True
                .Wrap = wdFindStop
                .Format = False
            End With
            head.Find.Execute Replace:=wdReplaceOne
        Next para
    Next pg

    MsgBox "Updated " & did & " item(s) on pages 1–4. Em dashes are now tight (no space/tab).", _
           vbInformation, "Done"
End Sub
===============================================================

