Sub Challenge()

' Loop through all sheets
    For Each ws In Worksheets

'Add the column headings
ws.Range("O1").Value = "Ticker"
ws.Range("P1").Value = "Value"
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Lowest % Increase"
ws.Range("N4").Value = "Greatest Total Volume"

Dim Max_Percent As Double
Dim Min_Percent As Double
Dim Max_Volume As Double

Max_Percent = 0
Min_Percent = 0
Max_Volume = 0

'Count the last number of rows
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Loop through all of the tickers
For i = 2 To lastrow
    
    'Find the next value that is larger than the old max percent value
   If ws.Cells(i, 11).Value > Max_Percent Then
   
   Max_Percent = ws.Cells(i, 11).Value
   
   ws.Range("P2") = Max_Percent
   ws.Range("O2") = ws.Cells(i, 9).Value
   
    'Format column K
      ws.Range("P2:P3").NumberFormat = "0.00%"

End If

Next i

For i = 2 To lastrow
 'Find the next value that is larger than the old min value
   If ws.Cells(i, 11).Value < Min_Percent Then
   
   Min_Percent = ws.Cells(i, 11).Value
   
   ws.Range("P3") = Min_Percent
   ws.Range("O3") = ws.Cells(i, 9).Value

End If

Next i

For i = 2 To lastrow
 'Find the next value that is larger than the old max value
   If ws.Cells(i, 12).Value > Max_Volume Then
   
   Max_Volume = ws.Cells(i, 12).Value
   
   ws.Range("P4") = Max_Volume
   ws.Range("O4") = ws.Cells(i, 9).Value
   
   End If
   
   Next i
   
Next ws

End Sub

