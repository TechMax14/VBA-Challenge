Sub HW2()
    Dim ticker As String
    Dim index, i, j, g, h, tickerCount As Integer
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim stockVol, yearlyChange, percentChange, greatestPercentage, lowestPercentage, greatestVolume As Double

   ' Set ws = ThisWorkbook.Worksheets("2018")
    For Each ws In ThisWorkbook.Worksheets
    
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row - 1
        
        index = 2
        stockVol = 0
        yearlyChange = 0
        percentChange = 0
        tickerCount = 1
        
        ws.Cells(1, "I").Value = "Ticker"
        ws.Cells(1, "J").Value = "Yearly Change"
        ws.Cells(1, "K").Value = "Percent Change"
        ws.Cells(1, "L").Value = "Volume"
        ws.Cells(1, "P").Value = "Ticker"
        ws.Cells(1, "Q").Value = "Value"
        ws.Cells(2, "O").Value = "Greatest % Increase"
        ws.Cells(3, "O").Value = "Greatest % Decrease"
        ws.Cells(4, "O").Value = "Greatest Total Volume"
        
        For i = 2 To lastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Ticker Symbol
                ticker = ws.Cells(i, 1).Value
                ws.Cells(index, "I").Value = ticker
                
                ' Stock Volume
                stockVol = stockVol + ws.Cells(i, 7).Value
                ws.Cells(index, "L").Value = stockVol
                
                ' Yearly Change
                yearlyChange = ws.Cells(i, "F").Value - ws.Cells(i - tickerCount + 1, "C").Value
                ws.Cells(index, "J").Value = yearlyChange
                If yearlyChange > 0 Then
                    ws.Cells(index, "J").Interior.Color = RGB(0, 255, 0) ' Green color
                ElseIf yearlyChange < 0 Then
                    ws.Cells(index, "J").Interior.Color = RGB(255, 0, 0) ' Red color
                End If
                
                ' Percent Change
                percentChange = (ws.Cells(i, "F").Value - ws.Cells(i - tickerCount + 1, "C").Value) / ws.Cells(i - tickerCount + 1, "C").Value * 100
                ws.Cells(index, "K").Value = percentChange
                
                index = index + 1
                stockVol = 0
                tickerCount = 1
            Else
                stockVol = stockVol + ws.Cells(i, 7).Value
                tickerCount = tickerCount + 1
            End If
        Next i
        
        ' Find the greatest percentage increase
        greatestPercentage = ws.Cells(2, "K").Value
        For j = 3 To lastRow
            If ws.Cells(j, "K").Value > greatestPercentage Then
                greatestPercentage = ws.Cells(j, "K").Value
                ticker = ws.Cells(j, "A").Value
            End If
        Next j
        ws.Cells(2, "Q") = greatestPercentage
        ws.Cells(2, "P") = ticker
        
        ' Find the greatest percentage decrease
        lowestPercentage = ws.Cells(2, "K").Value
        For g = 3 To lastRow
            If ws.Cells(g, "K").Value < lowestPercentage Then
                lowestPercentage = ws.Cells(g, "K").Value
                 ticker = ws.Cells(g, "A").Value
            End If
        Next g
        ws.Cells(3, "Q") = lowestPercentage
        ws.Cells(3, "P") = ticker
        
        ' Find the greatest volume
        greatestVolume = ws.Cells(2, "G").Value
        For h = 3 To lastRow
            If ws.Cells(h, "G").Value > greatestVolume Then
                greatestVolume = ws.Cells(h, "G").Value
                ticker = ws.Cells(h, "A").Value
            End If
        Next h
        ws.Cells(4, "Q") = greatestVolume
        ws.Cells(4, "P") = ticker
    
    Next ws
    
End Sub
