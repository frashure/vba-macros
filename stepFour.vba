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
