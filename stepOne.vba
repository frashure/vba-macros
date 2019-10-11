'TEMPORARY'
'ActiveCell.Offset(1).Formula = "=sum(U2, U" & lastRow & ")"'


'Turn off screen updating for speed performance'
Application.ScreenUpdating = False

'Rename original sheet'
ActiveSheet.Name = "All Surveys and Comments"

'Create month and year variables via user input'
Dim monthVar As String
Dim yearVar As String
monthVar = InputBox("Input month (ex. June): ")
yearVar = InputBox("Input year (ex. 2018): ")

'Add 4 under 4 and Q7 under 4 columns, fill with formula'
Sheets("All Surveys and Comments").Select
Range("U1").Select
Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
Range("U1").Value = "4 under 4"
Range("V1").Value = "Q7 under 4"

Dim lastRow As Integer
lastRow = ActiveSheet.Range("I1", ActiveSheet.Range("I1").End(xlDown)).Rows.Count

Range("U2").Formula = "=IF(COUNTIF(N2:T2, ""<5"")>3, 1, """")"
Range("U2", Range("U"&lastRow)).FillDown

Range("V2").Formula = "=IF(T2 <5, 1, """")"
Range("V2", Range("V"&lastRow)).FillDown

'Create month column, fill with formula'
Range("AJ1").Value = "MONTH"
Range("AJ2").Formula = "=text(AF2, ""mmmm"")"
Range("AJ2", Range("AJ"&lastRow)).FillDown

'Recode EEOC (for FBMS only)'
For i = 1 To lastRow
  If Range("H" & i).Value = "EEOC" Then
    If Range("AD" & i).Value = "FBMS" Then
      If Range("M" & i).Value = "FM-Travel" Then
      Range("H" & i).Value = "EEOC - Travel"
      ElseIf Range("M" & i).Value = "OFF" Then
        Range("H" & i).Value = "EEOC - Oracle"
        End if
    End If
  End If
Next i

'Recode ONRR and IBC (for FBMS only)'
For i = 1 To lastRow
  If (Range("H" & i).Value = "ONRR") Or (Range("H" & i).Value = "IBC") Then
    If Range("AD" & i).Value = "FBMS" Then
      Range("H" & i).Value = "DOI-OS"
      End If
    End If
Next i

'Recode Department of Interior to DOI-OS'
For i = 1 to lastRow
  If Range("H" & i).Value = "Dept Of Interior" Then
    Range("H" & i).Value = "DOI-OS"
  End If
  Next i

'Recode SOL for IT'
For i = 1 To lastRow
  If Range("AD" & i).Value = "IT" Then
    If InStr(1, Range("I" & i).Value, "@sol.doi.gov") > 0 Then
    Range("H" & i).Value = "SOL"
  End If
End If
Next i



'Create individual tabs for LoBs'
With ActiveWorkbook
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "All Surveys " & monthVar & " " & yearVar
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "FBMS Surveys " & monthVar & " " & yearVar
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "FM Surveys " & monthVar & " " & yearVar
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "HR Surveys " & monthVar & " " & yearVar
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "IT Surveys " & monthVar & " " & yearVar
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "FBMS Comments " & monthVar & " " & yearVar
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "FM Comments " & monthVar & " " & yearVar
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "HR Comments " & monthVar & " " & yearVar
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "IT Comments " & monthVar & " " & yearVar
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "WORK-FILE"
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Comments"
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Comments w Incidents"
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "4 below 4"
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Pie Chart"
  End With

'Filter and select data from target month; paste to All Surveyrs for Month sheet'
  Sheets("All Surveys and Comments").Select
  Range("$A:AJ").AutoFilter Field:=36, Criteria1:=monthVar
  Cells.SpecialCells(xlCellTypeVisible).Select
  Selection.Copy
  Sheets("All Surveys " & monthVar & " " & yearVar).Select
  ActiveSheet.Paste

  'Count number of rows; '

'Copy all FBMS data from monthly All Surveys sheet'
  Sheets("All Surveys " & monthVar & " " & yearVar).Select
  Range("$A:AJ").AutoFilter Field:=30, Criteria1:="FBMS"
  Cells.SpecialCells(xlCellTypeVisible).Select
  Selection.Copy
  Sheets("FBMS Surveys " & monthVar & " " & yearVar).Select
  ActiveSheet.Paste

'Copy all FM data from All Surveys sheet'
  Sheets("All Surveys " & monthVar & " " & yearVar).Select
  Range("$A:AJ").AutoFilter Field:=30, Criteria1:="FMD"
  Cells.SpecialCells(xlCellTypeVisible).Select
  Selection.Copy
  Sheets("FM Surveys " & monthVar & " " & yearVar).Select
  ActiveSheet.Paste

'Copy all HR data from All Surveys sheet'
  Sheets("All Surveys " & monthVar & " " & yearVar).Select
  Range("$A:AJ").AutoFilter Field:=30, Criteria1:="HRD"
  Cells.SpecialCells(xlCellTypeVisible).Select
  Selection.Copy
  Sheets("HR Surveys " & monthVar & " " & yearVar).Select
  ActiveSheet.Paste

'Copy all IT data from All Surveys sheet'
  Sheets("All Surveys " & monthVar & " " & yearVar).Select
  Range("$A:AJ").AutoFilter Field:=30, Criteria1:="ITD"
  Cells.SpecialCells(xlCellTypeVisible).Select
  Selection.Copy
  Sheets("IT Surveys " & monthVar & " " & yearVar).Select
  ActiveSheet.Paste

  'Turn off autofilter'
  Sheets("All Surveys and Comments").Select
  ActiveSheet.AutoFilterMode = False

'Copy all FBMS Comments data from FBMS Surveys sheet'
  Sheets("FBMS Surveys " & monthVar & " " & yearVar).Select
  Range("$A:AJ").AutoFilter Field:=4, Criteria1:="<>"
  Cells.SpecialCells(xlCellTypeVisible).Select
  Selection.Copy
  Sheets("FBMS Comments " & monthVar & " " & yearVar).Select
  ActiveSheet.Paste
  Sheets("FBMS Surveys " & monthVar & " " & yearVar).Select
  ActiveSheet.AutoFilterMode = False

'Copy all FM Comments data from FM Surveys sheet'
  Sheets("FM Surveys " & monthVar & " " & yearVar).Select
  Range("$A:AJ").AutoFilter Field:=4, Criteria1:="<>"
  Cells.SpecialCells(xlCellTypeVisible).Select
  Selection.Copy
  Sheets("FM Comments " & monthVar & " " & yearVar).Select
  ActiveSheet.Paste
  Sheets("FM Surveys " & monthVar & " " & yearVar).Select
  ActiveSheet.AutoFilterMode = False

'Copy all HR Comments data from HR Surveys sheet'
  Sheets("HR Surveys " & monthVar & " " & yearVar).Select
  Range("$A:AJ").AutoFilter Field:=4, Criteria1:="<>"
  Cells.SpecialCells(xlCellTypeVisible).Select
  Selection.Copy
  Sheets("HR Comments " & monthVar & " " & yearVar).Select
  ActiveSheet.Paste
  Sheets("HR Surveys " & monthVar & " " & yearVar).Select
  ActiveSheet.AutoFilterMode = False

'Copy all IT Comments data from FBMS Surveys sheet'
  Sheets("IT Surveys " & monthVar & " " & yearVar).Select
  Range("$A:AJ").AutoFilter Field:=4, Criteria1:="<>"
  Cells.SpecialCells(xlCellTypeVisible).Select
  Selection.Copy
  Sheets("IT Comments " & monthVar & " " & yearVar).Select
  ActiveSheet.Paste
  Sheets("IT Surveys " & monthVar & " " & yearVar).Select
  ActiveSheet.AutoFilterMode = False


  'Fill missing Positive, Negative, Neutral column headers; format'
  Sheets("Comments").Select
  Range("A1").Value = "Agency Name"
  Range("B1").Value = "Survey Comments"
  Range("C1").Value = "Q7 Score"
  Range("D1").Value = "Positive"
  Range("E1").Value = "Neutral"
  Range("F1").Value = "Negative"
  Range("A1:F1").Select
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
  Range("D1").Select
  With Selection.Font
      .Color = -11489280
  End With
  Range("E1").Select
  With Selection.Font
      .Color = -16711681
  End With
  Range("F1").Select
  With Selection.Font
      .Color = -16776961
  End With
  Columns("A:F").AutoFit

'Fill column headers in Comments w Incidents sheet; format'
Sheets("Comments w Incidents").Select
Range("A1").Value = "Agency Name"
Range("B1").Value = "Survey Comments"
Range("C1").Value = "Q1"
Range("D1").Value = "Q2"
Range("E1").Value = "Q3"
Range("F1").Value = "Q4"
Range("G1").Value = "Q5"
Range("H1").Value = "Q6"
Range("I1").Value = "Q7"
Range("J1").Value = "Submitter"
Range("K1").Value = "VIP"
Range("L1").Value = "Incident #"
Range("M1").Value = "Positive"
Range("N1").Value = "Neutral"
Range("O1").Value = "Negative"
Range("A1:O1").Select
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
Range("M1").Select
With Selection.Font
    .Color = -11489280
End With
Range("N1").Select
With Selection.Font
    .Color = -16711681
End With
Range("O1").Select
With Selection.Font
    .Color = -16776961
End With
Columns("A:O").AutoFit

'Return to All Surveys and Comments sheet, unfilter and delete added columns'
Sheets("All Surveys and Comments").Select
'Range("U:U,V:V,AJ:AJ").Delete'

  'Deselect'
  Application.CutCopyMode = False

  'Turn screen updating back on'
  Application.ScreenUpdating = True
