 ```VBA

Private Sub CommandButton1_Click()

Dim sht As Worksheet
Dim lngRow As Long
Dim LoopDate As Date
Set sht = Sheet1

'lngRow = sht.Cells.SpecialCells(xlCellTypeLastCell).Row + 1
lngRow = LastRowColumn(ActiveSheet, "Row") + 1

 If StartDate.Value > EndDate.Value Then
 
 MsgBox "Start Date needs to be before end date"
 Else
 LoopDate = StartDate.Value
 Do While LoopDate <= EndDate.Value
 
 If Weekday(LoopDate, vbMonday) < 6 Then
 'lngRow = sht.Cells.SpecialCells(xlCellTypeLastCell).Row + 1

Cells(lngRow, 1) = Format(LoopDate, "yyyymmdd") & "-" & Initials(NameStaff.Value)
 Cells(lngRow, 2) = Format(LoopDate, "dddd")
 Cells(lngRow, 3) = Format(LoopDate, "dd MMMM yyyy")
 Cells(lngRow, 4) = WorksheetFunction.Proper(NameStaff.Value)
 'Cells(lngRow, 5) = EndDate.Value
 'Cells(lngRow, 6) = CBSaturday.Value
 'Cells(lngRow, 7) = CBSaturday.Value
 'Cells(lngRow, 8) = Initials(NameStaff.Value)
 Cells(lngRow, 4).Activate
 LoopDate = LoopDate + 1
 lngRow = lngRow + 1
 
 
 'Is it a Saturday?
 
 ElseIf Weekday(LoopDate, vbMonday) = 6 And CBSaturday.Value = True Then
  Cells(lngRow, 1) = Format(LoopDate, "yyyymmdd") & "-" & Initials(NameStaff.Value)
 Cells(lngRow, 2) = Format(LoopDate, "dddd")
 Cells(lngRow, 3) = Format(LoopDate, "dd MMMM yyyy")
 Cells(lngRow, 4) = WorksheetFunction.Proper(NameStaff.Value)
 Cells(lngRow, 4).Activate
 LoopDate = LoopDate + 1
 lngRow = lngRow + 1
 
 'Is it a Sunday
 
 ElseIf Weekday(LoopDate, vbMonday) = 7 And CBSunday.Value = True Then
  Cells(lngRow, 1) = Format(LoopDate, "yyyymmdd") & "-" & Initials(NameStaff.Value)
 Cells(lngRow, 2) = Format(LoopDate, "dddd")
 Cells(lngRow, 3) = Format(LoopDate, "dd MMMM yyyy")
 Cells(lngRow, 4) = WorksheetFunction.Proper(NameStaff.Value)
 Cells(lngRow, 4).Activate
 LoopDate = LoopDate + 1
 lngRow = lngRow + 1
 
 Else
 LoopDate = LoopDate + 1
 
 End If
 
 
 Loop
 End If

 
End Sub



Sub LoopDate()

LoopDate = DateDiff("d", StartDate.Value, EndDate.Value)
 MsgBox "Start Date needs to be before end date"
'For a = lngRow To lngRow + LoopDate

'If a = lngRow Then
'Cells(a, 1) = Date
'Else
'Cells(a, 1) = Date + a - 1
'End If
'Next a

End Sub

Private Sub CommandButton2_Click()

Unload Me

End Sub

Private Sub CommandButton3_Click()
'
' ClearTimesheetPage Macro
'

'
    Cells.Select
    Selection.ClearContents
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Ref"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Day"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Date"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Name"
    Range("A2").Select
End Sub


Private Sub DTPickerStart_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub EndDate_Change()
    If EndDate.Text = "" Then
    EndDate.Text = Format(Now(), "dd/mm/yyyy")
    End If
End Sub


Private Sub OpenWord_Click()
Dim wordApp As Object

    Set wordApp = CreateObject("word.Application")

    wordApp.documents.Open "\\tsclient\Link to Remote\Timesheets\Timesheets Main Template.docx"

    wordApp.Visible = True
    wordApp.Activate
    wordApp.WindowState = xlNormal
    
End Sub

```
