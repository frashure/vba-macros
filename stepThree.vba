'Turn off screen updating for performance'
Application.ScreenUpdating = False


'Declare vairables'
Dim sheetname As String
sheetname = "WORK-FILE"
Dim lastRow As Integer
lastRow = ActiveSheet.Range("A1", Range("A1").End(xlDown)).Rows.Count

'Copy relevant columns into Comments w Incidents sheet'
    'Copy COMPANY column'
    Sheets(sheetname).Select
    Range("G2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

    'Paste COMPANY column into Comments w Incidents as Agency Name'
    Sheets("Comments w Incidents").Select
    Range("A2").Select
    ActiveSheet.Paste

    'Copy COMMENT column'
    Sheets(sheetname).Select
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

    'Paste COMMNENT column into Comments w Incidents'
    Sheets("Comments w Incidents").Select
    Range("B2").Select
    ActiveSheet.Paste

    'Copy Q1-Q7 range'
    Sheets(sheetname).Select
    Range("H2:N"&lastRow).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

    'Paste Q1-Q7 range into Comments w Incidents'
    Sheets("Comments w Incidents").Select
    Range("C2").Select
    ActiveSheet.Paste

    'Copy SUBMITTER column'
    Sheets(sheetname).Select
    Range("O2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

    'Paste SUBMITTER column into Comments w Incidents'
    Sheets("Comments w Incidents").Select
    Range("J2").Select
    ActiveSheet.Paste

    'Copy VIP colum'
    Sheets(sheetname).Select
    Range("R2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

    'Paste VIP column into Comments w Incidents'
    Sheets("Comments w Incidents").Select
    Range("K2").Select
    ActiveSheet.Paste

    'Copy INCIDENTID column'
    Sheets(sheetname).Select
    Range("S2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

    'Paste INCIDENTID column into Comments w Incidents'
    Sheets("Comments w Incidents").Select
    Range("L2").Select
    ActiveSheet.Paste

    'Copy Positive, Neutral, Negative columns'
    Sheets(sheetname).Select
    Range("B2:D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

    'Paste Positive, Neutral, Negative columns'
    Sheets("Comments w Incidents").Select
    Range("M2").Select
    ActiveSheet.Paste

    'Auto-format cell height'
    Columns("A:O").AutoFit
    Rows("2:" & lastRow).EntireRow.AutoFit

    'Sort by Agency Name'
    Sheets("Comments w Incidents").Select
    Columns("A:O").Sort key1:=Range("A:A"), order1:=xlAscending, Header:=xlNo

'Copy relevant columns into Comments sheet'

    'Copy Comments column from WORK-FILE tab
    Sheets(sheetname).Select
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

    'Paste into Comments tab
    Sheets("Comments").Select
    Range("B2").Select
    ActiveSheet.Paste

    'Copy COMPANY column from WORK-FILE'
    Sheets(sheetname).Select
    Range("G2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

    'Paste COMPANY column into Comments as Agency Name'
    Sheets("Comments").Select
    Range("A2").Select
    ActiveSheet.Paste

    'Copy Q7 score from WORK-FILE sheet'
    Sheets(sheetname).Select
    Range("N2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

    'Paste into Comments tab
    Sheets("Comments").Select
    Range("C2").Select
    ActiveSheet.Paste

    'Copy Positive, Neutral, and Negative scores from Comments w Incidents tab
    Sheets(sheetname).Select
    Range("B2:D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

    'Paste into Comments tab
    Sheets("Comments").Select
    Range("D2").Select
    ActiveSheet.Paste

    'Auto-format cell height'
    Columns("A:O").AutoFit
    Rows("2:" & lastRow).EntireRow.AutoFit

    'Sort by Agency Name'
    Sheets("Comments").Select
    Columns("A:F").Sort key1:=Range("A:A"), order1:=xlAscending, Header:=xlNo

    'Deselect'
    Application.CutCopyMode = False

'Create new workbooks, copy Comments and Comments w Incidents into new workbooks'
    Dim sourceWB As Workbook
    Set sourceWB = ActiveWorkbook

    Dim incidentsWB As Workbook
    Set incidentsWB = Workbooks.Add

    Dim commentsWB As Workbook
    Set commentsWB = Workbooks.Add

    sourceWB.Sheets("Comments w Incidents").Copy Before:=incidentsWB.Sheets("Sheet1")
    incidentsWB.Sheets("Sheet1").Name = "Removed"
    sourceWB.Sheets("Comments").Copy Before:=commentsWB.Sheets("Sheet1")
    commentsWB.Sheets("Sheet1").Name = "Removed"

    'Move to A1'
    Range("A1").Select

    'Turn on screen updating'
    Application.ScreenUpdating = True
