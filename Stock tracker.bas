Sub StockTracker()

For Each ws In Worksheets

    Dim i As Long
    Dim ticker As String
    
    Dim volume As Double
    volume = 0
    
    Dim tickrow As Double
    tick_row = 2
    
    Dim lastrow As Double
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

    
    For i = 2 To last_row
        
    current_tick = ws.Cells(i, 1).Value
    next_tick = ws.Cells(i + 1, 1).Value
    
    volume = volume + ws.Cells(i, 7).Value
            
        If current_tick <> next_tick Then
                
            ws.Range("I" & tick_row).Value = current_tick
            ws.Range("J" & tick_row).Value = volume
            
            tick_row = tick_row + 1
            
            volume = 0
                
        End If
             
    Next i

    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Total Stock Volume"
        
Next ws

End Sub
