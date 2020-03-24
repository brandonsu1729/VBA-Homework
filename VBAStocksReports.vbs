Sub worksheet_loop()
    Dim sheet As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Dim lastrow As Long
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    'Ticker
        tickerRow = 2
        ws.Range("I1").Value = "Ticker"
        For i = 2 To lastrow
            If ws.Cells(i, 1) <> ws.Cells(tickerRow - 1, 9) Then
                ws.Cells(tickerRow, 9) = ws.Cells(i, 1)
                tickerRow = tickerRow + 1
            End If
        Next i
    
    'Yearly Change
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
        Dim open_stock As Double
    
        open_stock = ws.Range("C2")
    'lastTickerRow = Cells(Rows.Count, 9).End(xlUp).Row 'last row that has  a ticker
        CurrentTickerRow = 2
        stock_volume = 0
        For x = 2 To lastrow
        'open_row = Range("A:A").Find(what = Cells(x, 9)).Row
        'close_row = Range("A:A").Find(what = Cells(x + 1, 9)).Row - 1
        'MsgBox (open_row + close_row)
            If ws.Cells(x, 1) <> ws.Cells(x - 1, 1) Then 'new ticker
                If x <> 2 Then 'anything but first change in ticker
                    ws.Cells(CurrentTickerRow, 10) = ws.Cells(x - 1, 6) - open_stock 'yearly change, taking last close and minus first open
                    'conditional formatting
                    If ws.Cells(CurrentTickerRow, 10) < 0 Then
                        ws.Cells(CurrentTickerRow, 10).Interior.ColorIndex = 22
                    ElseIf ws.Cells(CurrentTickerRow, 10) > 0 Then
                        ws.Cells(CurrentTickerRow, 10).Interior.ColorIndex = 43
                    Else
                        ws.Cells(CurrentTickerRow, 10).Interior.ColorIndex = 36
                    End If
                    
                    
                    If open_stock = 0 Then 'division by zero handled
                        ws.Cells(CurrentTickerRow, 11) = 0
                    Else
                        ws.Cells(CurrentTickerRow, 11) = ws.Cells(CurrentTickerRow, 10) / open_stock 'percent change
                    End If
                
                    ws.Cells(CurrentTickerRow, 12) = stock_volume
                    stock_volume = ws.Cells(x, 7) 'set stock volume to new ticker
                    open_stock = ws.Cells(x, 3) 'resetting openstock to the next ticker
                    CurrentTickerRow = CurrentTickerRow + 1
                End If
            ElseIf (x = lastrow) Then 'for the last row since x,1<>x-1,1 does not account for it
                ws.Cells(CurrentTickerRow, 10) = ws.Cells(x, 6) - open_stock
                'conditional format
                If ws.Cells(CurrentTickerRow, 10) < 0 Then
                    ws.Cells(CurrentTickerRow, 10).Interior.ColorIndex = 22
                ElseIf ws.Cells(CurrentTickerRow, 10) > 0 Then
                    ws.Cells(CurrentTickerRow, 10).Interior.ColorIndex = 43
                Else
                    ws.Cells(CurrentTickerRow, 10).Interior.ColorIndex = 36
                End If
                
                
                If open_stock = 0 Then  'division by zero handled
                    ws.Cells(CurrentTickerRow, 11) = 0
                Else
                    ws.Cells(CurrentTickerRow, 11) = ws.Cells(CurrentTickerRow, 10) / open_stock
                    
                ws.Cells(CurrentTickerRow, 12) = stock_volume
                End If
            Else
                stock_volume = stock_volume + ws.Cells(x, 7) 'accumulate stock volume within the ticker
            End If
        
        Next x
    'greatest percent increase, decrease and total volume highest
        greatest_increase = 0
        ws.Range("N2") = "Greatest % Increase"
        greatest_decrease = 0
        ws.Range("N3") = "Greatest % Decrease"
        highest_volume = 0
        ws.Range("N4") = "Greatest Total Volume"
        ws.Range("O1") = "Ticker"
        ws.Range("P1") = "Value"
    
        For y = 2 To CurrentTickerRow
            If ws.Cells(y, 11) < greatest_decrease Then
                greatest_decrease = ws.Cells(y, 11)
                ws.Range("O3") = ws.Cells(y, 9)
            ElseIf ws.Cells(y, 11) > greatest_increase Then
                greatest_increase = ws.Cells(y, 11)
                ws.Range("O2") = ws.Cells(y, 9)
            End If
            If ws.Cells(y, 12) > highest_volume Then 'compare total volume
                ws.Range("O4") = ws.Cells(y, 9)
                highest_volume = ws.Cells(y, 12)
            End If
        Next y
        ws.Range("P2") = greatest_increase
        ws.Range("P3") = greatest_decrease
        ws.Range("P4") = highest_volume
        
    Next
    
End Sub
