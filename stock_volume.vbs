Sub StockData_Volume():

' Loop through all sheets
    For Each ws In Worksheets

'Add the column headings
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percentage Change"
ws.Range("L1").Value = "Total Stock Volume"

 ' Set an initial variable for holding the ticker name
  Dim Ticker As String

'------------------------------------------------
' CALCULATING TOTAL STOCK VOLUME
'------------------------------------------------
  ' Set an initial variable for holding the total per ticker
  Dim Ticker_Volume As Double
    Ticker_Volume = 0

  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  'Count the last number of rows
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Loop through all of the tickers
    For i = 2 To lastrow

    ' Check if we are still within the same ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the Ticker name
      Ticker_Name = ws.Cells(i, 1).Value

      ' Add to the Ticker Volume Total
      Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value

      ' Print the Ticker Name in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Ticker_Name

      ' Print the Ticker Volume to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = Ticker_Volume

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Brand Total
      Ticker_Volume = 0

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Brand Total
      Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value

    End If

  Next i
  
Next ws

End Sub

