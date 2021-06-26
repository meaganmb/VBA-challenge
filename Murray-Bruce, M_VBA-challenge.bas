Attribute VB_Name = "Module1"
Sub ticker()

'Set all variables
Dim ticker_name As String
Dim yearly_open As Double
Dim yearly_end As Double
Dim counter As Double
Dim open_close As Double
Dim stock_volume As Double

'define variables
'using ws for each line for it to work all worksheets


For Each ws In Worksheets
    counter = 2
    stock_volume = 0
    open_close = 2
    
    ' Determine the Last Row
   LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    
    'Bonus part
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    
    For i = 2 To LastRow
        stock_volume = stock_volume + ws.Cells(i, 7).Value
        ticker_name = ws.Cells(i, 1).Value
        yearly_open = ws.Cells(open_close, 3).Value
        
         If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
         yearly_end = ws.Cells(i, 6)
         ws.Cells(counter, 9).Value = ticker_name
         ws.Cells(counter, 10).Value = yearly_end - yearly_open
         
          ' Print the Ticker name to summary
      ws.Range("I" & counter).Value = ticker_name
         
         If yearly_open = 0 Then
            ws.Cells(counter, 11).Value = Null
        Else
            ws.Cells(counter, 11).Value = ((yearly_end - yearly_open) / yearly_open)
        End If
        ws.Cells(counter, 12).Value = stock_volume
        
        'colors
        If ws.Cells(counter, 10).Value > 0 Then
            ws.Cells(counter, 10).Interior.ColorIndex = 4
        Else
            ws.Cells(counter, 10).Interior.ColorIndex = 3
        End If
        
        ' Change formatting for percent change column
        ws.Cells(counter, 11).NumberFormat = "0.00%"
        
        stock_volume = 0
        counter = counter + 1
        open_close = i + 1
        
        End If
        
    Next i
    
    'Greatest % decrease
    Set myRange = ws.Range("K2:K" & counter)
    min_value = Application.WorksheetFunction.Min(myRange)
    ws.Range("Q3") = min_value
    ws.Range("Q3").NumberFormat = "0.00%"
    
    min_ticker = Application.WorksheetFunction.Match(min_value, ws.Range("K2:K" & counter))
    ws.Range("P3") = ws.Cells(min_ticker + 1, 9)
    
    'Greatest % increase (using same range as greatest % decrease)
    max_value = Application.WorksheetFunction.Max(myRange)
    ws.Range("Q2") = max_value
    ws.Range("Q2").NumberFormat = "0.00%"
    
      max_ticker = Application.WorksheetFunction.Match(max_value, ws.Range("K2:K" & counter))
    ws.Range("P2") = ws.Cells(max_ticker + 1, 9)
    
    'Greatest total stock volume
    Set myRange2 = ws.Range("L2:L" & counter)
    max_stock_value = Application.WorksheetFunction.Max(myRange2)
    ws.Range("Q4") = max_stock_value
    
    max_stock_volume_ticker = Application.WorksheetFunction.Match(max_stock_value, ws.Range("L2:L" & counter))
    ws.Range("P1") = ws.Cells(max_stock_volume_ticker + 1, 9)
    
    
    
    
    
    AutoFit Columns
    combined_sheet.Columns("I:L").AutoFit
    combined_sheet.Columns("O:Q").AutoFit
    
    Next ws

End Sub

