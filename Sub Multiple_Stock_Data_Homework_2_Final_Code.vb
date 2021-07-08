Sub Multiple_Stock()

Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets
ws.Activate

Dim TickerCounter As Integer
Dim TickerList As String
Dim lastrow As Long
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
TickerCounter = 2
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim YearlyChange As Double
Dim RunningTotal As Double
ClosePrice = 0
RunningTotal = 0
OpenPrice = Cells(2, 3).Value


For i = 2 To lastrow


If Cells(i + 1, 1) <> Cells(i, 1).Value Then
   TickerList = Cells(i, 1).Value
   Range("I" & TickerCounter) = TickerList
   ClosePrice = ClosePrice + Cells(i, 6).Value
   RunningTotal = Cells(i, 7).Value + RunningTotal
   Range("J" & TickerCounter) = ClosePrice - OpenPrice
   YearlyChange = ClosePrice - OpenPrice
   If OpenPrice <> 0 Then
    Range("K" & TickerCounter) = YearlyChange / OpenPrice
   Else
    Range("K" & TickerCounter) = 0
   End If
   
   Range("L" & TickerCounter) = RunningTotal


TickerCounter = TickerCounter + 1
OpenPrice = Cells(i + 1, 3)
ClosePrice = 0
RunningTotal = 0



Else
RunningTotal = RunningTotal + Cells(i, 7).Value




End If

Next i

Cells(1, 10).Value = "Yearly Change"
Cells(1, 9).Value = "Ticker"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Volume"

For i = 2 To lastrow


If Cells(i, 10).Value >= 0 Then
    Cells(i, 10).Interior.ColorIndex = 4

Else
    Cells(i, 10).Interior.ColorIndex = 3
    
    
End If

Next i

Next ws

End Sub
