Sub LoopThroughWS()

Dim ws As Worksheet
    ' Loop through all sheets
    For Each ws In ThisWorkbook.Sheets
        Call StockTicker(ws)
    Next ws

End Sub



Sub StockTicker(ws As Worksheet)

 ' Set an initial variables for stocks
  Dim stockSymbol As String
  Dim yearlyDelta, openingValue, closingValue, percentDelta As Double
  yearlyDelta = 0
  openingValue = 0
  closingValue = 0
  percentDelta = 0
  
  Dim hasValue As Boolean
  Dim volumeTotal As Double
  volumeTotal = 0
  
  Dim marketDay, firstDayOfYear, lastDayOfYear As Date
  firstDayOfYear = 1 / 1 / 2019
  lastDayOfYear = 12 / 20 / 2019

      
  ' Keep track of the location for each stock in the summary table
  Dim summaryTableRow, lastRowMain, lastRowSummary As Double
  summaryTableRow = 2
  lastRowMain = ws.Cells(Rows.Count, 1).End(xlUp).Row

  ' Loop through all stock market days
  For i = 2 To lastRowMain

    ' compare stock symbols for same value, if not then add summary item
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the stock symbol name
      stockSymbol = ws.Cells(i, 1).Value
      
      'Get last day of the year
      marketDay = CDate(Format$(ws.Cells(i, 2).Value, "####-##-##"))
      lastDayOfYear = DateSerial(Year(marketDay), 12, 31)
      
      'Get last day of the year
        closingValue = ws.Cells(i, 3).Value
        volumeTotal = volumeTotal + ws.Cells(i, 7).Value
      
      'compute difference if end of year
      yearlyDelta = closingValue - openingValue
      
      'check to see if the yearly delta or opening value is 0, if it is assert the value of 0% for
      'percentdelta or set percent delta to yearly to avoid divide by 0 error, else calculate percent change
        'stock does not exist ever and set to N/A
      If openingValue = 0 And closingValue = 0 Then
        hasValue = False
      ElseIf yearlyDelta = 0 Then
        percentDelta = 0
        hasValue = True
      Else
        percentDelta = yearlyDelta / openingValue
        hasValue = True
      End If
        ' Print the stock symbol in the Summary Table
        ws.Range("I1").Value = "Ticker"
        ws.Range("I" & summaryTableRow).Value = stockSymbol
      
      'Add Delta to summary table
      ws.Range("J1").Value = "Yearly Change"
      If ws.Range("J" & summaryTableRow).Value > 0 Then
        ws.Range("J" & summaryTableRow).Interior.Color = vbGreen
        ws.Range("J" & summaryTableRow).Value = yearlyDelta
      ElseIf ws.Range("J" & summaryTableRow).Value < 0 Then
        ws.Range("J" & summaryTableRow).Interior.Color = vbRed
        ws.Range("J" & summaryTableRow).Value = yearlyDelta
      Else
        ws.Range("J" & summaryTableRow).Interior.Color = vbYellow
        ws.Range("J" & summaryTableRow).Value = yearlyDelta
      End If
    
      ' add percent delta to summary table if value is N/A skip formatting
      ws.Range("K1").Value = "Percent Change"
      ws.Range("K" & summaryTableRow).Value = percentDelta
      If hasValue = False Then
        ws.Range("K" & summaryTableRow).Value = "N/A"
      Else
        ws.Range("K" & summaryTableRow).Value = percentDelta
        ws.Range("K" & summaryTableRow).NumberFormat = "0.00%"
      End If
      ' total stock volume
      ws.Range("L1").Value = "Total Stock Volume"
      ws.Range("L" & summaryTableRow).Value = volumeTotal

      ' Add one to the summary table row
      summaryTableRow = summaryTableRow + 1
      
      ' Reset the values at end of year
        yearlyDelta = 0
        percentDelta = 0
        volumeTotal = 0
        stockSymbol = ""
        openingValue = 0
        closingValue = 0
        
    ' If the cell immediately following a row is the same stock symbol
    Else
      'Get first day of the year
      marketDay = CDate(Format$(ws.Cells(i, 2).Value, "####-##-##"))
      firstDayOfYear = DateSerial(Year(marketDay), 1, 1)
      
      
      'Grab opening day of year Market Price checking for any 0 value opening
      'if the first day value is 0 do nothing else if the first the first value of the stock
      'is not 0 while the opening value remains 0, set price to first non-zero value
      'else just total the volume, opening price already grabbed
        If marketDay = firstDayOfYear And ws.Cells(i, 3).Value <> 0 Then
            openingValue = ws.Cells(i, 3).Value
            volumeTotal = volumeTotal + ws.Cells(i, 7).Value
        ElseIf marketDay = firstDayOfYear And ws.Cells(i, 3).Value = 0 Then
            'do nothing
        ElseIf openingValue = 0 And ws.Cells(i, 3).Value <> 0 Then
            openingValue = ws.Cells(i, 3).Value
            volumeTotal = volumeTotal + ws.Cells(i, 7).Value
        Else
            volumeTotal = volumeTotal + ws.Cells(i, 7).Value
        End If
    End If
  Next i
  
  
  'summary of greatest increase, decrease, total volume
  ws.Range("O2").Value = "Greatest % Increase:"
  ws.Range("O3").Value = "Greatest % Decrease:"
  ws.Range("O4").Value = "Greatest Total Volume:"
  ws.Range("P1").Value = "Ticker"
  ws.Range("Q1").Value = "Value"
  
  Dim greatestIncrease, greatestDecrease, greatestTotalVolume As Double
  greatestIncrease = 0
  greatestTotalVolume = 0
  greatestDecrease = 10000000
  Dim stockSymbolIncrease, stockSymbolDecrease, stockSymbolVolume As String
  
  'loop through summary table to compare values of highest or lowest values
  lastRowSummary = ws.Cells(Rows.Count, 11).End(xlUp).Row
  For i = 2 To lastRowSummary
    If ws.Range("K" & i).Value <> "N/A" Then
        If ws.Range("K" & i).Value > greatestIncrease Then
            greatestIncrease = ws.Range("K" & i).Value
            stockSymbolIncrease = ws.Range("I" & i).Value
        End If
        If ws.Range("K" & i).Value < greatestDecrease Then
            greatestDecrease = ws.Range("K" & i).Value
            stockSymbolDecrease = ws.Range("I" & i).Value
        End If
        If ws.Range("L" & i).Value > greatestTotalVolume Then
            greatestTotalVolume = ws.Range("L" & i).Value
            stockSymbolVolume = ws.Range("I" & i).Value
        End If
    End If
  Next i
  
  ws.Range("P2").Value = stockSymbolIncrease
  ws.Range("Q2").Value = greatestIncrease
  ws.Range("Q2").NumberFormat = "0.00%"
  ws.Range("P3").Value = stockSymbolDecrease
  ws.Range("Q3").Value = greatestDecrease
  ws.Range("Q3").NumberFormat = "0.00%"
  ws.Range("P4").Value = stockSymbolVolume
  ws.Range("Q4").Value = greatestTotalVolume

End Sub


