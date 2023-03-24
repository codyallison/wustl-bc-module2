Attribute VB_Name = "Module1"
Sub stock_data_yearly()

    Dim row As Double
    Dim summaryRow As Double
    Dim lookUpRow As Integer
    
    Dim openPrice As Double
    Dim closePrice As Double
    
    Dim yearChange As Double
    Dim percentChange As Double
    Dim ticker As String
    Dim totalVolume As Double
    
    Dim maxPercentIncrease As Double
    Dim maxPercentDecrease As Double
    Dim maxTotalVolume As Double
    Dim maxIncreaseTicker As String
    Dim maxDecreaseTicker As String
    Dim maxVolumeTicker As String
    Dim ws As Worksheet
    
    'loops through each sheet
    For Each ws In ThisWorkbook.Worksheets
    
        'first stock opening price and volume
        
        openPrice = ws.Cells(2, 3).Value
        totalVolume = ws.Cells(2, 7).Value
        
        'establish starting row for summary information/calculate last row
        summaryRow = 2
        LastRow = ws.Range("A" & Rows.Count).End(xlUp).row
        
        'creates headers/labels
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        
        'loop through the rows
        For row = 2 To LastRow
            
            'if next ticker different than current
            If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
            
                'stores final closing price
                closePrice = ws.Cells(row, 6).Value
            
                'stores ticker from final row or current stock
                ticker = ws.Cells(row, 1).Value
            
                'calculates yearly change/percent change
                yearChange = (closePrice - openPrice)
                
                    'to avoid 0/0 calculation and establish percent change
                    If openPrice = 0 Then
                        percentChange = 0
                        
                    Else
                        percentChange = (yearChange / openPrice)
                        
                    End If
                    
                'prints ticker to summary table
                ws.Range("I" & summaryRow).Value = ticker
            
                'prints year Change w/ formatting to summary table
                ws.Range("J" & summaryRow).Value = yearChange
                    If ws.Range("J" & summaryRow).Value > 0 Then
                        ws.Range("J" & summaryRow).Interior.ColorIndex = 4
                    ElseIf ws.Range("J" & summaryRow).Value < 0 Then
                        ws.Range("J" & summaryRow).Interior.ColorIndex = 3
                    End If
                'prints percentage change w/formatting to summary table
                ws.Range("K" & summaryRow).Value = percentChange
                ws.Range("K" & summaryRow).NumberFormat = "0.00%"
                'prints total volume to summary
                ws.Range("L" & summaryRow).Value = totalVolume
            
            
            
                'advances summary to next row
                summaryRow = summaryRow + 1
            
                'reset ticker, volume, and established Next opening price
            
                ticker = ""
                totalVolume = 0
                openPrice = ws.Cells(row + 1, 3).Value
                
            'if next row ticker matches, just add the volume to count
            Else
            totalVolume = totalVolume + ws.Cells(row, 7).Value
            
            
            End If
    
        Next row
        
            'tracking max stats/ applies percentage formating
            maxPercentIncrease = Application.WorksheetFunction.Max(ws.Range("K:K"))
            ws.Cells(2, 16).Value = maxPercentIncrease
            lookUpRow = Application.WorksheetFunction.Match(maxPercentIncrease, ws.Range("K:K"), 0)
            ws.Cells(2, 15).Value = ws.Cells(lookUpRow, 9)
            ws.Cells(2, 16).NumberFormat = "0.00%"
            
            maxPercentDecrease = Application.WorksheetFunction.Min(ws.Range("K:K"))
            ws.Cells(3, 16).Value = maxPercentDecrease
            lookUpRow = WorksheetFunction.Match(maxPercentDecrease, ws.Range("K:K"), 0)
            ws.Cells(3, 15).Value = ws.Cells(lookUpRow, 9)
            ws.Cells(3, 16).NumberFormat = "0.00%"
            
            maxTotalVolume = Application.WorksheetFunction.Max(ws.Range("L:L"))
            ws.Cells(4, 16).Value = maxTotalVolume
            lookUpRow = WorksheetFunction.Match(maxTotalVolume, ws.Range("L:L"), 0)
            ws.Cells(4, 15).Value = ws.Cells(lookUpRow, 9)
            
    Next ws
            
End Sub
