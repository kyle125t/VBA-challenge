Sub VBAwallst()

Dim tickerSymbol As String '<ticker>
Dim tickerCount As Double
Dim tickerYearTotal As Double
Dim tickStart As Double '<open>
Dim tickEnd As Double '<close>
Dim totalVol As Double

For Each ws In Worksheets ' cycle through each sheet
    totalVol = 0 ' sets count to zero to add each ticker value
    tickerCount = 2 ' allows counting of each symbol after header
    tickerYearTotal = 2 ' allows counting for open/close after header
    
    ws.Range("I1").Value = "Ticker" ' following four lines set column header names for I1:L1
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    For r = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row ' goes down row by row in each sheet
        totalVol = totalVol + ws.Cells(r, 7).Value ' adds column G value to total volume as each row is checked
        tickerSymbol = ws.Cells(r, 1).Value ' sets tickerSymbol to the stock symbol in column A
        tickStart = ws.Cells(tickerYearTotal, 3) ' pulls cell value for <open> from column C
        
        If ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value Then ' essentially runs following loop onces stock symbol changes
            tickEnd = ws.Cells(r, 6) ' sets <close> value
            ws.Cells(tickerCount, 9).Value = tickerSymbol ' places stock symbol in column I
            ws.Cells(tickerCount, 10).Value = tickEnd - tickStart ' gives yearly change for whole stock
            
            If tickStart = 0 Then
                ws.Cells(tickerCount, 11).Value = Null ' prevents dividing by 0
            Else
                ws.Cells(tickerCount, 11).Value = (tickEnd - tickStart) / tickStart ' gives percent change for column K
            End If
            ws.Cells(tickerCount, 12).Value = totalVol ' gives total volume for all cells in column L
            
            If ws.Cells(tickerCount, 10).Value > 0 Then
                ws.Cells(tickerCount, 10).Interior.ColorIndex = 4 ' green for positive change
            Else
                ws.Cells(tickerCount, 10).Interior.ColorIndex = 3 ' red for negative change
            End If
            
            ws.Cells(tickerCount, 11).NumberFormat = "0.00%" ' sets formatting for column K
            
            totalVol = 0 ' resets volume count as we move to next symbol
            tickerCount = tickerCount + 1 ' moves through ticker symbols
            tickerYearTotal = r + 1 ' next yearly change row
        End If
    Next r
Next ws
       
End Sub
