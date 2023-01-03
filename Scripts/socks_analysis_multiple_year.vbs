Sub StockInfo():
    
'Repeat code for each sheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
    
        'Make headers for Column I to L
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"
        
    
        'While loop
        Dim i As Long
    
        i = 2
        While IsEmpty(ws.Cells(i, 1)) = False
            
            'get ticker symbol
            ws.Cells(i, 9).Value = ws.Cells(i, 1).Value
            
            'open to end (yearly) change
            ws.Cells(i, 10).Value = ws.Cells(i, 6).Value - ws.Cells(i, 3).Value
            
            'percent change from open to close
            ws.Cells(i, 11).Value = ws.Cells(i, 10).Value / ws.Cells(i, 3).Value
            
            'total volume
            ws.Cells(i, 12).Value = ws.Cells(i, 7).Value
            
            'Conditional Formatting +/- change
            If ws.Cells(i, 10).Value < 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 22
            
            ElseIf ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 43
            
            End If
            
    
            i = i + 1
            
        Wend
        
        'Bonus = get biggest % increase, % decrease, Volume
        
    
    
        'Labels
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
    
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        'Values
        ws.Cells(2, 17).Value = WorksheetFunction.Max(ws.Range("K:K"))
        ws.Cells(3, 17).Value = WorksheetFunction.Min(ws.Range("K:K"))
        ws.Cells(4, 17).Value = WorksheetFunction.Max(ws.Range("L:L"))
        ws.Range("P2") = WorksheetFunction.XLookup(ws.Range("Q2"), ws.Range("K:K"), ws.Range("I:I"))
        ws.Range("P3") = WorksheetFunction.XLookup(ws.Range("Q3"), ws.Range("K:K"), ws.Range("I:I"))
        ws.Range("P4") = WorksheetFunction.XLookup(ws.Range("Q4"), ws.Range("L:L"), ws.Range("I:I"))
        
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("Q4").NumberFormat = "0"
        
        
        'Autofit and Number Format
        ws.Cells.EntireColumn.AutoFit
        ws.Range("K:K").NumberFormat = "0.00%"
        
        Next ws

End Sub

