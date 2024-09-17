Sub StockData()

Dim r As Long 'row variable
Dim lastrow As Long 'last row variable
Dim qchangelastrow As Long 'quarterly change last row variable
Dim ws As Worksheet 'worksheet variable
Dim ws2 As Worksheet 'worksheet variable
Dim WorksheetName As String 'worksheet name variable
Dim ticker As String 'ticker variable
Dim QuarterOpen As Double 'quarter price open change variable
Dim QuarterClose As Double 'quarter price close change variable
Dim QuarterChange As Double 'quarter open - close variable
Dim QuarterPercent As Double 'quarter percent change variable
Dim biggest_ticker As String
Dim maxValue As Double
Dim smallest_ticker As String
Dim minValue As Double
Dim maxTotalVol As String


'Loop through all the sheets
For Each ws In Worksheets

    ' Count rows to last row instead of specifying
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    qchangelastrow = ws.Cells(Rows.Count, 10).End(xlUp).Row
    
    'Determine row counter one for the ticker
    rtick = 2
    rquarter = 2
    volume = 0
    
    'Grab Worksheet Name
    WorksheetName = ws.Name
    
    'Add the titles to the columns
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    'Create a script that loops through the stock tickers for each quarter
        For r = 2 To lastrow                                                         'For row 2 to the last row in column 1
            currentcell = ws.Cells(r, 1).Value                                  'the value in the current row cell in column 1
            nextcell = ws.Cells(r + 1, 1).Value                               'the value in the next row cell in column 1
            ticker = ws.Cells(r, 1).Value                                          'the value in the row in column 1 is a string value, ticker
            volume = volume + ws.Cells(r, 7).Value
            'formattedDate = Format(ws.Cells(r, 2).Value, "mm/dd/yyyy")
            
            If currentcell <> nextcell Then                                         'If the current cell does not equal the next cell
                ws.Cells(rtick, 9).Value = ticker                                'then use the counter rtick to place the value of the ticker column 1 into column 9
                QuarterOpen = ws.Cells(rquarter, 3).Value                     'the quarter open stock price value is in column 3
                rquarter = r + 1
                QuarterClose = ws.Cells(r, 6).Value                            'the quarter close stock price value is in column 6
                QuarterChange = QuarterClose - QuarterOpen          'the quarter change between the closing and opening stock price
                PercentChange = QuarterChange / QuarterOpen
                ws.Cells(rtick, 10).Value = QuarterChange                       'the row counter in column 10 will enter the quarter change values
                ws.Cells(rtick, 10).NumberFormat = "0.00"
                ws.Cells(rtick, 11).Value = PercentChange
                ws.Cells(rtick, 11).NumberFormat = "0.00%"
                ws.Cells(rtick, 12).Value = volume
                volume = 0
                rtick = rtick + 1                                                        'rtick counter adds one each iteration to fill in the values in column 9
            End If
            
            If ws.Cells(r, 10).Value > 0 Then
                ws.Cells(r, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(r, 10).Value < 0 Then
                ws.Cells(r, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(r, 10).Interior.ColorIndex = 0
            End If
      
      Next r
      
            'Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume"
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"
            ws.Cells(2, 15).Value = "Greatest % Increase"
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(4, 15).Value = "Greatest Total Volume"
    
            lastrowfx = ws.Cells(Rows.Count, 11).End(xlUp).Row
    
            'Greatest % increase
            maxValueRow = Application.WorksheetFunction.Match(Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrowfx)), ws.Range("K2:K" & lastrowfx), 0)
            ws.Range("Q2").Value = ws.Cells(maxValueRow + 1, 11).Value
            ws.Range("Q2").NumberFormat = "0.00%"
            biggest_ticker = ws.Cells(maxValueRow + 1, 1).Value
            ws.Range("P2").Value = biggest_ticker
            
            'Greatest % decrease
            minValueRow = Application.WorksheetFunction.Match(Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrowfx)), ws.Range("K2:K" & lastrowfx), 0)
            ws.Range("Q3").Value = ws.Cells(minValueRow + 1, 11).Value
            ws.Range("Q3").NumberFormat = "0.00%"
            smallest_ticker = ws.Cells(minValueRow + 1, 1).Value
            ws.Range("P3").Value = smallest_ticker
        
            'Greatest total volume
            maxVolRow = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrowfx))
            ws.Range("Q4").Value = maxVolRow
            ws.Range("P4").Value = maxTotalVol
            ws.Range("Q4").NumberFormat = "0"
            
Next ws

End Sub
