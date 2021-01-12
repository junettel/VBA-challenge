Sub StockLoop()

    For Each ws in Worksheets

        Dim WorksheetName As String
        Dim i As Long
        Dim j As Long
        Dim k As Long
        Dim LastRow As Long
        Dim LastRowAlt As Long
        Dim TickerCount As Long
        Dim PctChange As Double
        Dim GreatestIncrease As Double
        Dim GreatestDecrease As Double
        Dim GreatestVolume As Double


        WorksheetName = ws.Name

        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"

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


        ' Bonus
        SummaryLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        GreatestIncrease = ws.Cells(2, 11).Value
        GreatestDecrease = ws.Cells(2, 11).Value
        GreatestVolume = ws.Cells(2, 12).Value

        For k = 2 To SummaryLastRow

            ' Greatest % increase
            If ws.Cells(k, 11).Value > GreatestIncrease Then
                GreatestIncrease = ws.Cells(k, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(k, 9)

                Else
                GreatestIncrease = GreatestIncrease

            End If
            
            ' Greatest % decrease
            If ws.Cells(k, 11).Value < GreatestDecrease Then
                GreatestDecrease = ws.Cells(k, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(k, 9)

                Else
                GreatestDecrease = GreatestDecrease

            End If

            ' Greatest total volume
            If ws.Cells(k, 12).Value > GreatestVolume Then
                GreatestVolume = ws.Cells(k, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(k, 9)

                Else
                GreatestVolume = GreatestVolume

            End If

            ws.Cells(2, 17).Value = Format(GreatestIncrease, "Percent")
            ws.Cells(3, 17).Value = Format(GreatestDecrease, "Percent")
            ws.Cells(4, 17).Value = Format(GreatestVolume, "Scientific")

        Next k

        ws.Columns("I:Q").AutoFit

    Next ws

End Sub