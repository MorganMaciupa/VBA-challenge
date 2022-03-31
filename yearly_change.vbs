Sub StockData_YearlyChange():

' Loop through all sheets
    For Each ws In Worksheets

'--------------------------------------------------
'CALCULATING YEARLY CHANGE
'--------------------------------------------------

  ' Set an initial variable for holding the total per ticker
'   Dim i As Double
'   Dim Yearly_Change As Double
'   Dim Year_Open As Double
'   Dim Year_Closes As Double
'   Dim Percentage_Change As Double
  Yearly_Change = 0
  Percentage_Change = 0

  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Double
  Summary_Table_Row = 2
  
  'Count the last number of rows
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Loop through all of the tickers
    For i = 2 To lastrow

 'Find the first line of each ticker
    If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then

      ' Set the Year Open
      Year_Open = ws.Cells(i, 3).Value
      
      End If
    
      'Find the last line of each ticker
      If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
      
      'Set the Year Close
      Year_Close = ws.Cells(i, 6).Value
      
        'Calculate the yearly change
      If Year_Open = 0 Or Year_Close = 0 Then
        Yearly_Change = 0
      
      Else: Yearly_Change = Year_Close - Year_Open
      
      'Calculate the percentage change
      Percentage_Change = (Yearly_Change / Year_Open)

      ' Print the Yearly Change to the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
      
      'Print the percentage change to the Summary Table
      ws.Range("K" & Summary_Table_Row).Value = Percentage_Change
      
      'Format column K
      ws.Range("K:K").NumberFormat = "0.00%"

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Ticker Totals
      Yearly_Change = 0
      Percentage_Change = 0

    End If
    
    End If

  Next i
  
Next ws

End Sub


