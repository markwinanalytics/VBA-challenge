Sub vbachallenge()
Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
ws.Activate
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).row
Cells(1, "i").Value = "Ticker"
Cells(1, "j").Value = "Yearly Change"
Cells(1, "k").Value = "Percent Change"
Cells(1, "l").Value = "Total Stock Volume"
Dim openp As Double
Dim closep As Double
Dim yearly_change As Double
Dim ticker As String
Dim percent_change As Double
Dim vol As Double
vol = 0
Dim row As Double
row = 2
Dim column As Integer
column = 1
openp = Cells(2, column + 2).Value
For i = 2 To LastRow
If Cells(i + 1, column).Value <> Cells(i, column).Value Then
ticker = Cells(i, column).Value
Cells(row, column + 8).Value = ticker
closep = Cells(i, column + 5).Value
yearly_change = closep - openp
Cells(row, column + 9).Value = yearly_change
If (openp = 0 And closep = 0) Then
percent_change = 0
ElseIf (openp = 0 And closep <> 0) Then
percent_change = 1
Else
percent_change = yearly_change / openp
Cells(row, column + 10).Value = percent_change
Cells(row, column + 10).NumberFormat = "0.00%"
End If
vol = vol + Cells(i, column + 6).Value
Cells(row, column + 11).Value = vol
row = row + 1
openp = Cells(i + 1, column + 2)
vol = 0
Else
vol = vol + Cells(i, column + 6).Value
End If
Next i
YearlyChangeLastRow = ws.Cells(Rows.Count, column + 8).End(xlUp).row
For j = 2 To YearlyChangeLastRow
If (Cells(j, column + 9).Value > 0 Or Cells(j, column + 9).Value = 0) Then
Cells(j, column + 9).Interior.ColorIndex = 4
Cells(j, columm + 9).Font.ColorIndex = 1
ElseIf Cells(j, column + 9).Value < 0 Then
Cells(j, column + 9).Interior.ColorIndex = 3
Cells(j, column + 9).Font.ColorIndex = 1
End If
Next j
Cells(2, column + 14).Value = "Greatest % Increase"
Cells(3, column + 14).Value = "Greatest % Decrease"
Cells(4, column + 14).Value = "Greatest Total Volume"
Cells(1, column + 15).Value = "Ticker"
Cells(1, column + 16).Value = "Value"
For Z = 2 To YearlyChangeLastRow
If Cells(Z, column + 10).Value = Application.WorksheetFunction.Max(ws.Range("K2:k" & YearlyChangeLastRow)) Then
Cells(2, column + 15).Value = Cells(Z, column + 8).Value
Cells(2, column + 16).Value = Cells(Z, column + 10).Value
Cells(2, column + 16).NumberFormat = "0.00%"
ElseIf Cells(Z, column + 10).Value = Application.WorksheetFunction.Min(ws.Range("K2:k" & YearlyChangeLastRow)) Then
Cells(3, column + 15).Value = Cells(Z, column + 8).Value
Cells(3, column + 16).Value = Cells(Z, column + 10).Value
Cells(3, column + 16).NumberFormat = "0.00%"
ElseIf Cells(Z, column + 11).Value = Application.WorksheetFunction.Max(ws.Range("L2:l" & YearlyChangeLastRow)) Then
Cells(4, column + 15).Value = Cells(Z, column + 8).Value
Cells(4, column + 16).Value = Cells(Z, column + 11).Value
End If
Next Z
Next ws
End Sub
