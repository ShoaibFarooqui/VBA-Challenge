
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
        Dim j As Integer
        i = 2
        j = 2
        
        Dim total_close As Double
        total_close = 0
        
        Dim total_open As Double
        total_open = 0
        
        Dim total_volume As LongLong
        total_volume = 0
        
        While IsEmpty(ws.Cells(i, 1)) = False
            
            'open to end
            total_open = total_open + ws.Cells(i, 3).Value
            total_close = total_close + ws.Cells(i, 6).Value
            total_volume = total_volume + ws.Cells(i, 7).Value
            
            'get ticker symbol
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ws.Cells(j, 9).Value = ws.Cells(i, 1).Value
                'get yearly change
                ws.Cells(j, 10).Value = total_close - total_open
                'percent change
                ws.Cells(j, 11).Value = ws.Cells(j, 10).Value / total_open
                'total Volume
                ws.Cells(j, 12).Value = total_volume
                
                'Conditional Formatting +/- change
                If ws.Cells(j, 10).Value < 0 Then
                    ws.Cells(j, 10).Interior.ColorIndex = 22
    
                ElseIf ws.Cells(j, 10).Value > 0 Then
                    ws.Cells(j, 10).Interior.ColorIndex = 43
    
                End If
            
                j = j + 1
            End If



            
            total_open = 0
            total_close = 0
            total_volume = 0
            i = i + 1
            
        Wend
        ws.Range("J:J").NumberFormat = "0.00"
        ws.Range("K:K").NumberFormat = "0.00%"

        

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
        Range("Q2:Q3").NumberFormat = "0.00%"
    
    
        Cells.EntireColumn.AutoFit
    Next ws
    
End Sub



