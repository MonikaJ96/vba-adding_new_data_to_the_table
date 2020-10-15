Option Explicit

Sub adding_data()

Dim information As String
Dim newdate As String
Dim newdealer As String
Dim newproduct As String
Dim newprofit As String

Range("A1").End(xlDown).Offset(1, 0).Select

AddDate:
    newdate = InputBox("Add date")
    
On Error GoTo WrongDate
newdate = CDate(newdate)
On Error GoTo 0

newdealer = InputBox("Add dealer")
newproduct = InputBox("Add product")

AddProfit:
    newprofit = InputBox("Add profit")
    
On Error GoTo WrongProfit
newprofit = CInt(newprofit)
On Error GoTo 0

ActiveCell.Value = newdate
ActiveCell.Offset(0, 1).Value = newdealer
ActiveCell.Offset(0, 2).Value = newproduct
ActiveCell.Offset(0, 3) = newprofit

ActiveCell.Offset(0, 3).Select
Selection.NumberFormat = "#,##0.00 $"

With Range("A1", Range("A1").End(xlDown).End(xlToRight)).Borders
    .LineStyle = xlContinuous
    .Weight = xlThin
End With

Exit Sub

WrongDate:
MsgBox "Add date again", vbRetryCancel
Resume AddDate

WrongProfit:
MsgBox "Add profit again", vbRetryCancel
Resume AddProfit

End Sub
