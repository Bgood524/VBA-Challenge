Sub Testing()

Dim TickerCounter As Integer
Dim TickerList As String
Dim lastrow As Long
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
TickerCounter = 0
TickerList = Range("I:I")



For i = 2 To lastrow


If Cells(i, 1) <> TickerList Then
   TickerCounter = TickerCounter + 1
   Cells(TickerCounter + 1, 9).Value = Cells(i, 1).Value

End If

Next i


