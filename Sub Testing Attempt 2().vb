Sub Testing()

Dim TickerCounter As Integer
Dim TickerList As String
Dim lastrow As Long
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
TickerCounter = 2



For i = 2 To lastrow


If Cells(i+1, 1) <> Cells(i,1).Value Then
   TickerList = Cells(i,1)Value
   Range("I" & TickerCounter) = TickerList
   TickerCounter = TickerCounter +1
   

End If

Next i

End Sub
