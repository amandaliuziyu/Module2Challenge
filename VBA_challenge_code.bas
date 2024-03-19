Sub loop_change_total():
    For Each ws In Worksheets
        Dim lastRow As Long, i As Long
        Dim openPrice As Double, closePrice As Double
        Dim totalVolume As Double
        Dim tickerSymbol As String
        Dim startOfTheYear As Long, endOfTheYear As Long
        Dim tickernum As Long
        tickernum = 1
        Dim maxin As Double
        maxin = 0
        Dim maxde As Double
        maxde = 0
        Dim maxvol As Double
        maxvol = 0
        Dim maxinTicker As String
        Dim maxdeTicker As String
        Dim maxvolTicker As String
        
        
        lastRow = Cells(Rows.Count, "A").End(xlUp).Row
        totalVolume = 0
    
    ' Setup headers
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percentage Change"
        Cells(1, 12).Value = "Total Volume"
        
    ' Other functionality headers and labels
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15) = "Greatest Total Volume"
        
    
        For i = 2 To lastRow
            If i = 2 Or Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
                tickernum = tickernum + 1
                tickerSymbol = Cells(i, 1).Value
                openPrice = Cells(i, 3).Value
                totalVolume = 0
            End If
        
            totalVolume = totalVolume + Cells(i, 7).Value
        
            If i = lastRow Or Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                closePrice = Cells(i, 6).Value
        
            ' Output the results next to the original data
                Cells(tickernum, 9).Value = tickerSymbol
            
                Cells(tickernum, 10).Value = closePrice - openPrice
                
                If openPrice <> 0 Then
                    Cells(tickernum, 11).Value = (closePrice - openPrice) / openPrice
                Else
                    Cells(tickernum, 11).Value = 0
                End If
                Cells(tickernum, 11).NumberFormat = "0.00%"
                Cells(tickernum, 12).Value = totalVolume
            End If
            
            If Cells(tickernum, 10).Value > 0 Then
                Cells(tickernum, 10).Interior.ColorIndex = 4
            Else:
                Cells(tickernum, 10).Interior.ColorIndex = 3
                
            End If
            
        Next i
        
        'Find Max Increase, Max Decrease, and Max Volume
        
        Application.WorksheetFunction.Max(Range("K2:K")) = maxin
        Application.WorksheetFunction.Min(Range("K2:K")) = maxde
        Application.WorksheetFunction.Max(Range("L2:L")) = maxin
        
        'Output the values we found
        Cells(2, 17).Value = maxin
        Cells(2, 17).NumberFormat = "0.00%"
        Cells(3, 17).Value = maxde
        Cells(3, 17).NumberFormat = "0.00%"
        Cells(4, 17).Value = maxvol
            
        
    
    Next ws
    
    
End Sub