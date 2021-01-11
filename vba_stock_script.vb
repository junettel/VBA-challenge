Sub StockLoop()

    For Each ws in Worksheets

    Dim WorksheetName As String
    Dim i As Long
    Dim j As Long
    Dim LastRow As Long
    Dim LastRowAlt As Long
    Dim TickerCount As Long
    Dim PctChange As Double

    WorksheetName = ws.Name

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    TickerCount = 2
    j = 2
    LastRow = ws.UsedRange.Rows.Count
    ' LastRowAlt = ws.Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To LastRow

            ' Record ticker symbols if ticker name changed in col 9
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ws.Cells(TickerCount, 9).Value = ws.Cells(i, 1).Value
            ' Record yearly change in col 10
            ws.Cells(TickerCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value

                ' Percent change from opening price to closing price
                If ws.Cells(j, 3).Value <> 0 Then
                PctChange = (( ws.Cells(i, 6).Value - ws.Cells(j, 3).Value ) / ws.Cells(j, 3).Value)
                ' Record percent change in col 11 and format as percent
                ws.Cells(TickerCount, 11).Value = Format(PctChange, "Percent")

                Else
                ws.Cells(TickerCount, 11).Value = Format(0, "Percent")

                End If

                ' Highlight negative yearly change in red
                If ws.Cells(TickerCount, 10).Value < 0 Then
                ws.Cells(TickerCount, 10).Interior.ColorIndex = 3

                ' Highlight positive yearly change in green
                Else
                ws.Cells(TickerCount, 10).Interior.ColorIndex = 4

                End If 

            ' Calculate total stock volume
            ws.Cells(TickerCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
            
            TickerCount = TickerCount + 1

            j = i + 1

            End If

        Next i
    
        ws.Columns("I:L").AutoFit

    Next ws

End Sub