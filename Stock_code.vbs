Sub Stock_Market()

For Each ws In Worksheets

    'Set variable for ticker symbol
    Dim Ticker As String

    'Set variable for ticker count
    Dim TickerCount As Integer
    TickerCount = 0
    
    'Set variable for stock volume
    Dim VolumeTotal As Double
        Volume = 0

    'Set row where summary data will start
    Dim SummaryRow As Integer
    SummaryRow = 2

    'Set summary area headings
    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Yearly Change"
    ws.Cells(1, 11) = "Percent Change"
    ws.Cells(1, 12) = "Total Stock Volume"
    
    'Count total number of rows
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Loop through all rows
    For i = 2 To LastRow

        'Check to see if ticker symbol has changed in next row, if not then
         If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            'Set ticker
            Ticker = ws.Cells(i, 1).Value

            'Add to the total ticker Volume
            VolumeTotal = VolumeTotal + ws.Cells(i, 7).Value

            'Print ticket symbol in summary area
            ws.Cells(SummaryRow, 9).Value = Ticker

            'Print total ticker Volume
            ws.Cells(SummaryRow, 12).Value = VolumeTotal

            'Print yearly change
            ws.Cells(SummaryRow, 10).Value = ws.Cells(i, 6).Value - ws.Cells(i - TickerCount, 3).Value
            
                'Format color of yearly change cell if value is positive
                If ws.Cells(SummaryRow, 10).Value > 0 Then
                
                    ws.Cells(SummaryRow, 10).Interior.ColorIndex = 4
            
                'Format color of yearly change cell if value is negative
                ElseIf ws.Cells(SummaryRow, 10).Value < 0 Then
                
                    ws.Cells(SummaryRow, 10).Interior.ColorIndex = 3
            
                'No cell color format if there is no yearly change
                Else
            
                End If
                                            
                'Determine if percent change is going to cause an error, if so
                If ws.Cells(i - TickerCount, 3).Value = 0 Then
                
                    ws.Cells(SummaryRow, 11).Value = 0
                
               'Print percentage change
                Else
                ws.Cells(SummaryRow, 11).Value = Format((ws.Cells(i, 6).Value - ws.Cells(i - TickerCount, 3).Value) / ws.Cells(i - TickerCount, 3).Value, "Percent")
            
                End If
            
            'Add 1 to SummaryRow for next value
            SummaryRow = SummaryRow + 1

            'Reset total ticker volume
            VolumeTotal = 0
            
            'Reset ticker count
            TickerCount = 0
                
    'If the next row has the same ticker symbol
        Else
            'Add to the total ticker volume
            VolumeTotal = VolumeTotal + ws.Cells(i, 7).Value
            
            'Add to the ticker count
            TickerCount = TickerCount + 1
                
        End If
            
    Next i

    'Set greatest value row and column headings
    ws.Cells(2, 15) = "Greatest % Increase"
    ws.Cells(3, 15) = "Greatest % Decrease"
    ws.Cells(4, 15) = "Greatest Total Volume"
    ws.Cells(1, 16) = "Ticker"
    ws.Cells(1, 17) = "Value"

    'Set variables for greatest values
    Dim GreatestVolume, GreatestIncrease, GreatestDecrease As Double
        GreatestVolume = Application.WorksheetFunction.Max(ws.Range("L:L"))
        GreatestIncrease = Application.WorksheetFunction.Max(ws.Range("K:K"))
        GreatestDecrease = Application.WorksheetFunction.Min(ws.Range("K:K"))
                
    'Print greatest percentage increase
    ws.Cells(2, 17) = Format(GreatestIncrease, "Percent")
    
    'Print greatest percentage decrease
    ws.Cells(3, 17) = Format(GreatestDecrease, "Percent")
    
    'Print greatest total volume
    ws.Cells(4, 17) = GreatestVolume
         
    'Print ticker value of Greatest Increase
    ws.Range("P2").Value = WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(ws.Range("Q2").Value, ws.Range("K:K"), 0))
    
    'Print ticker value of Greatest Decrease
    ws.Range("P3").Value = WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(ws.Range("Q3").Value, ws.Range("K:K"), 0))
    
    'Print ticker value of Greatest Volume
    ws.Range("P4").Value = WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(ws.Range("Q4").Value, ws.Range("L:L"), 0))
   
    'Autofit new columns
    ws.Columns("I:Q").AutoFit
            
Next ws

End Sub