'Send relevant columns to WORK-FILE sheet'

  'Disable screen update for performance'
  Application.ScreenUpdating = False

    Dim sheetname As String
    sheetname = ActiveSheet.Name
    Sheets(sheetname).Select

    'Clear contents of WORK-FILE below the headers'
    Sheets("WORK-FILE").Select
    Rows("2:" & Rows.Count).ClearContents

    'Select COMPANY column
    Sheets(sheetname).Select
    Range("H1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

    'Paste COMPANY column in WORK-FILE
    Sheets("WORK-FILE").Select
    Range("G1").Select
    ActiveSheet.Paste

    'Copy COMMENT column'
    Sheets(sheetname).Select
    Range("D1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

    'Paste COMMENT column into WORK-FILE'
    Sheets("WORK-FILE").Select
    Range("A1").Select
    ActiveSheet.Paste

    'Copy Q1-Q7 range'
    Sheets(sheetname).Select
    Range("N1:T1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

    'Paste Q1-Q7 range into WORK-FILE'
    Sheets("WORK-FILE").Select
    Range("H1").Select
    ActiveSheet.Paste

    'Copy SUBMITTER column'
    Sheets(sheetname).Select
    Range("X1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

    'Paste SUBMITTER column into WORK-FILE'
    Sheets("WORK-FILE").Select
    Range("O1").Select
    ActiveSheet.Paste

    'Copy ASSIGNEE column'
    Sheets(sheetname).Select
    Range("Z1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

    'Paste ASSIGNEE column into WORK-FILE'
    Sheets("WORK-FILE").Select
    Range("Q1").Select
    ActiveSheet.Paste

    'Copy VIP column'
    Sheets(sheetname).Select
    Range("AI1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

    'Paste VIP column into WORK-FILE'
    Sheets("WORK-FILE").Select
    Range("R1").Select
    ActiveSheet.Paste

    'Copy INCIDENTID column'
    Sheets(sheetname).Select
    Range("B1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

    'Paste INCIDENTID column'
    Sheets("WORK-FILE").Select
    Range("S1").Select
    ActiveSheet.Paste

    'Copy SUMMARY column'
    Sheets(sheetname).Select
    Range("W1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

    'Paste SUMMARY column into WORK-FILE'
    Sheets("WORK-FILE").Select
    Range("T1").Select
    ActiveSheet.Paste

    'Copy ASSIGNEDGROUP column'
    Sheets(sheetname).Select
    Range("Y1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

    'Paste ASSIGNEDGROUP column into WORK-FILE'
    Sheets("WORK-FILE").Select
    Range("U1").Select
    ActiveSheet.Paste

    'Copy SUBMITDATE column'
    Sheets(sheetname).Select
    Range("E1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

    'Paste SUBMITDATE column into WORK-FILE'
    Sheets("WORK-FILE").Select
    Range("W1").Select
    ActiveSheet.Paste

    'Copy CATE1 and CATE2 columns'
    Sheets(sheetname).Select
    Range("L1:M1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

    'Paste CATE1 and CATE2 columns into WORK-FILE'
    Sheets("WORK-FILE").Select
    Range("X1").Select
    ActiveSheet.Paste

    'Format column headers in WORK-FILE'
    Sheets("WORK-FILE").Select
    Range("B1").Value = "Positive"
    Range("C1").Value = "Neutral"
    Range("D1").Value = "Negative"
    Range("A1:Y1").Select
    With Selection.Interior
       .Pattern = xlSolid
       .PatternColorIndex = xlAutomatic
       .ThemeColor = xlThemeColorAccent5
       .TintAndShade = 0.799981688894314
       .PatternTintAndShade = 0
    End With
    With Selection.Font
        .Name = "Calibri"
        .Size = 14
        .Bold = True
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("B1").Select
    With Selection.Font
        .Color = -11489280
    End With
    Range("C1").Select
    With Selection.Font
        .Color = -16711681
    End With
    Range("D1").Select
    With Selection.Font
        .Color = -16776961
    End With

    'Resize row height'
    Dim lastRow As Integer
    lastRow = ActiveSheet.Range("A1", Range("A1").End(xlDown)).Rows.Count
    Columns("A:Y").AutoFit
    Rows("2:" & lastRow).EntireRow.AutoFit

    'Move to top-left'
    Range("A1").Select

    'Turn screen updating back on'
    Application.ScreenUpdating = True

    'Deselect'
    Application.CutCopyMode = False
