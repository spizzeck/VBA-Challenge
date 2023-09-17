# VBA-Challenge
Sub TickerLoop()


 ' Loop through all of the worksheets in the active workbook.
    Dim ws As Worksheet
         
    For Each ws In Worksheets


' Set an initial variable for holding Ticker Symbol
    Dim Ticker As String
    
  ' Set an initial variable for Yearly Change
  Dim YearlyChange As Double
  YearlyChange = 0
  
  'Set variable for Percent Change
  Dim PercentChange As Double
  PercentChange = 0
  
  'Set initial variable for Total Volume
  Dim TotalVolume As Double
  TotalVolume = 0

  ' Keep track of the location for each Ticker Symbol, Yearly Change, Percent Change, and Total Volume
  Dim Ticker_Yearly_Row As Integer
  Ticker_Yearly_Row = 2
  
  Dim Open_Price_Row As Double
  Open_Price_Row = 2
  
  Dim LastRow As Double
  LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

  ' Loop through all Tickers for Yearly Change,Percent Change, and Total Volume
  For i = 2 To LastRow

    ' Check if the same Ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the Ticker Symbol
      Ticker = ws.Cells(i, 1).Value

      ' Add to the Yearly Change
      YearlyChange = ((ws.Cells(i, 6).Value - ws.Cells(Open_Price_Row, 3).Value)) '+ YearlyChange
      
      'Add to the Yearly Percent Change
      PercentChange = (((ws.Cells(i, 6).Value) - (ws.Cells(Open_Price_Row, 3).Value)) / ws.Cells(Open_Price_Row, 3).Value)
      
      'Add to Total Volume
      TotalVolume = TotalVolume + ws.Cells(i, 7).Value
      
      ' Print the ticker Symbol in the Summary Table
      ws.Range("I" & Ticker_Yearly_Row).Value = Ticker

      ' Print the Yearly Change in the Summary Table
      ws.Range("J" & Ticker_Yearly_Row).Value = YearlyChange
      
      'Print the Yearly Percent Change in the SUmmary Table
      ws.Range("K" & Ticker_Yearly_Row).Value = PercentChange
      
      'Format to Percentage to two decimal places
      ws.Range("K" & Ticker_Yearly_Row).NumberFormat = "0.00%"
      
      'Print the Total Volume in Summary Table
      ws.Range("L" & Ticker_Yearly_Row).Value = TotalVolume
      
      'If Else for Conditional Yearly Change
    If ws.Range("J" & Ticker_Yearly_Row).Value > 0 Then
        
        ws.Range("J" & Ticker_Yearly_Row).Interior.ColorIndex = 4
    
    ElseIf ws.Range("J" & Ticker_Yearly_Row).Value < 0 Then
        
        ws.Range("J" & Ticker_Yearly_Row).Interior.ColorIndex = 3
    
    End If

      ' Add one to the summary table row
      Ticker_Yearly_Row = Ticker_Yearly_Row + 1
      
     ' Reset the Yearly Change
      YearlyChange = 0
      
      'Reset the Total Percentage
      PercentChange = 0
      
      'Reset the Total Volume
      TotalVolume = 0
      
      Open_Price_Row = i + 1

    ' If the cell immediately following a row is the same Ticker...
    Else

      
      'Add to the Total Volume
      TotalVolume = TotalVolume + ws.Cells(i, 7).Value
       

    End If

  Next i
  
  'Find Max and Min Values of Data
  
  ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & LastRow))
  ws.Range("Q2").NumberFormat = "0.00%"
    
    ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & LastRow))
    ws.Range("Q3").NumberFormat = "0.00%"
    
    ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & LastRow))
    
    max_ticker = WorksheetFunction.Match(ws.Range("Q2").Value, ws.Range("K2:K" & LastRow), 0)
    
    min_ticker = WorksheetFunction.Match(ws.Range("Q3").Value, ws.Range("K2:K" & LastRow), 0)
    
    max_volume_ticker = WorksheetFunction.Match(ws.Range("Q4").Value, ws.Range("L2:L" & LastRow), 0)
    
    ws.Range("P2").Value = ws.Cells(max_ticker + 1, 9).Value
    
    ws.Range("P3").Value = ws.Cells(min_ticker + 1, 9).Value
    
    ws.Range("P4").Value = ws.Cells(max_volume_ticker + 1, 9).Value
    
  Next ws
    
End Sub
Sub Column()

' Loop through all of the worksheets in the active workbook.
    Dim ws As Worksheet
         
    For Each ws In Worksheets
    
        'Set value for header and format
        ws.Range("A1:G1") = Array("Ticker", "Date", "Open", "High", "Low", "Close", "Volume")
        ws.Range("A1:G1").Font.Bold = True
        ws.Range("A1:G1").HorizontalAlignment = xlCenter
        'New Columns and Format
        ws.Range("I1:L1") = Array("Ticker_Symbol", "Yearly_Change", "Yearly_%_Change", "Total_Stock_Volume")
        ws.Range("I1:L1").Font.Bold = True
        ws.Range("I1:L1").HorizontalAlignment = xlCenter
        'Headers for Max and Min
        ws.Range("P1:Q1") = Array("Ticker", "Result")
        ws.Range("P1:Q1").Font.Bold = True
        ws.Range("P1:Q1").HorizontalAlignment = xlCenter
        'Headers for Max and Min
        'Headers for Max and Min
        ws.Cells(2, 15).Value = "Max % Change"
        ws.Cells(3, 15).Value = "Min % Change"
        ws.Cells(4, 15).Value = "Max Stock Volume"
        ws.Range("O2:O4").Font.Bold = True
        ws.Range("O2:O4").HorizontalAlignment = xlCenter
        'Set Column Width for Max % Change, Min % Change, Max Stock Volume
        ws.Columns("A:Q").ColumnWidth = 18
        'Center Data for All Worksheets
        ws.Columns("A:Q").HorizontalAlignment = xlCenter
        
        
    Next ws

End Sub
