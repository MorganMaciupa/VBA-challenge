Sub Conditional_Formatting():

' Loop through all sheets
    For Each ws In Worksheets

'Count the last number of rows
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Loop through all of the tickers
    For i = 2 To lastrow
  
  ' Check if the Yearly_Change is greater than zero
  If ws.Cells(i, 10).Value >= 0 Then
  
    ws.Cells(i, 10).Interior.ColorIndex = 4
    
    'if the value is less than zero
    Else
    
    ws.Cells(i, 10).Interior.ColorIndex = 3
 
End If

Next i

Next

End Sub




