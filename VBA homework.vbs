Sub stocks()

For Each ws In Worksheets

'variables
Dim Rows As Long
    Rows = ws.Range("A1").End(xlDown).Row
    'MsgBox Rows
    
    'Total Volume
Dim total_vol As LongLong
    total_vol = 0 'set initial value
'Stock open variable
Dim stock_open As Double
    stock_open = ws.Cells(2, 3).Value 'set initial value
'stock close variable
Dim stock_close As Double
    stock_close = 0 'set initial value
'counter to keep track of row for the current ticker
Dim ticker_count As Integer
    ticker_count = 1 'set initial value
Dim titles As Variant

'set titles for new columns and rows
titles = VBA.Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
titles2 = VBA.Array("Ticker", "Value")
titles3 = VBA.Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume")

ws.Range("I1:L1").Value = titles
ws.Range("p1:q1").Value = titles2
transposetitles3 = Application.WorksheetFunction.Transpose(titles3)
ws.Range("O2:O4").Value = transposetitles3



'loop for all rows with values of interest
For i = 2 To Rows

    If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
        total_vol = total_vol + ws.Cells(i, 7).Value 'add up total volume
    Else
        total_vol = total_vol + ws.Cells(i, 7).Value 'add total volume final time
        ticker_count = ticker_count + 1 'increase ticker to move to next output row
        stock_close = ws.Cells(i, 6).Value 'get stock close value
        ws.Cells(ticker_count, 10).Value = stock_open - stock_close 'get stock change
        ws.Cells(ticker_count, 11).Value = (ws.Cells(ticker_count, 10).Value / stock_open) 'get percent change
        ws.Cells(ticker_count, 11).NumberFormat = "0.00%" 'format %
        ws.Cells(ticker_count, 12).Value = total_vol 'post total volume
        ws.Cells(ticker_count, 9).Value = ws.Cells(i, 1).Value 'post ticker name
            If ws.Cells(ticker_count, 10).Value > 0 Then 'format increase/decrese color
                ws.Cells(ticker_count, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(ticker_count, 10).Interior.ColorIndex = 3
            End If
                
        stock_open = ws.Cells(i + 1, 3).Value 'reset to new stock open value for next stock
    End If
Next i

'Bonus
    'variables
    Dim best_inc_Per As Double
    
    Dim best_inc_tic As String
    
    best_inc_Per = ws.Cells(2, 11).Value
    best_inc_tick = ws.Cells(2, 9).Value
    

Dim BRows As Integer
    BRows = ws.Range("I1").End(xlDown).Row
ws.Cells(2, 16).Value = best_inc_tick
ws.Cells(2, 17).Value = best_inc_Per
    For i = 2 To BRows
        If ws.Cells(i, 11).Value > best_inc_Per Then 'find greatest value
            best_inc_Per = ws.Cells(i, 11).Value
            best_inc_tick = ws.Cells(i, 9).Value
            ws.Cells(2, 16).Value = best_inc_tick
            ws.Cells(2, 17).Value = best_inc_Per
            ws.Cells(2, 17).NumberFormat = "0.00%"
        End If
    Next i
        
        
        
    Dim best_dec_per As Double
    Dim best_dec_tic As String
    
    best_dec_per = ws.Cells(2, 11).Value
    best_dec_tick = ws.Cells(2, 9).Value

ws.Cells(3, 16).Value = best_dec_tick
ws.Cells(3, 17).Value = best_dec_per
    For i = 2 To BRows
        If ws.Cells(i, 11).Value < best_dec_per Then 'find greatest value
            best_dec_per = ws.Cells(i, 11).Value
            best_dec_tick = ws.Cells(i, 9).Value
            ws.Cells(3, 16).Value = best_dec_tick
            ws.Cells(3, 17).Value = best_dec_per
            ws.Cells(3, 17).NumberFormat = "0.00%"
        End If
    Next i
    
    
    Dim great_vol As LongLong
    Dim great_vol_tick As String
    
    great_vol = ws.Cells(2, 12).Value
    great_vol_tick = ws.Cells(2, 9).Value

ws.Cells(4, 16).Value = great_vol_tick
ws.Cells(4, 17).Value = great_vol
    For i = 2 To BRows
        If ws.Cells(i, 13).Value > great_vol Then 'find greatest value
            great_vol = ws.Cells(i, 12).Value
            great_vol_tick = ws.Cells(i, 9).Value
            ws.Cells(4, 16).Value = great_vol_tick
            ws.Cells(4, 17).Value = great_vol
        End If
    Next i
Next ws

End Sub


