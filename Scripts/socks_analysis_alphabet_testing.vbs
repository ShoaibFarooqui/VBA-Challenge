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
        
        ws.Cells.EntireColumn.AutoFit
        
        'Search for Bonus Values (biggest % increase, % decrease, Volume)
        Dim inc As Double
        Dim dec As Double
        Dim vol As LongLong
        'ticker names
        Dim inc_tik As String
        Dim dec_tik As String
        Dim vol_tik As String
        
        'bonus values of current sheet
        inc = WorksheetFunction.Max(ws.Range("K:K"))
        dec = WorksheetFunction.Min(ws.Range("K:K"))
        vol = WorksheetFunction.Max(ws.Range("L:L"))

        If inc > increase Then
        inc_tik = WorksheetFunction.XLookup(inc, ws.Range("K:K"), ws.Range("I:I"))
        increase = inc
        End If
        
        If dec < decrease Then
        dec_tik = WorksheetFunction.XLookup(dec, ws.Range("K:K"), ws.Range("I:I"))
        decrease = dec
        End If
        
        If vol > volume Then
        vol_tik = WorksheetFunction.XLookup(vol, ws.Range("L:L"), ws.Range("I:I"))
        volume = vol
        End If
        

        
    Next ws
        
    Sheets(1).Select
    
        'Bonus = get biggest % increase, % decrease, Volume
        
    
    
    'Labels
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"

    
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    'Values
    Cells(2, 17).Value = increase
    Cells(3, 17).Value = decrease
    Cells(4, 17).Value = volume
    Cells(2, 16).Value = inc_tik
    Cells(3, 16).Value = dec_tik
    Cells(4, 16).Value = vol_tik
    
    Range("Q4").NumberFormat = "0"
    
    
        'Autofit
    Cells.EntireColumn.AutoFit

End Sub
