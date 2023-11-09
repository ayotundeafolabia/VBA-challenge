Sub CalculatePriceDifferenceWithFormattingAndTableOnAllWorksheets()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim dataRange As Range
    Dim resultRow As Long
    Dim currentTicker As String
    Dim firstDayOpeningPrice As Double
    Dim lastDayClosingPrice As Double
    Dim totalVolume As Double
    Dim priceDifference As Double
    Dim percentageChange As Double
    Dim maxIncreasePercentage As Double
    Dim minDecreasePercentage As Double
    Dim greatestVolume As Double
    Dim tickerMaxIncrease As String
    Dim tickerMinDecrease As String
    Dim tickerGreatestVolume As String

    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Sheets(Array("2018", "2019", "2020"))
        ' Find the last row in column A
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

        ' Define the data range
        Set dataRange = ws.Range("A1:G" & lastRow)

        ' Add headers to the result columns
        ws.Cells(1, 8).Value = "Ticker"
        ws.Cells(1, 9).Value = "Price Difference"
        ws.Cells(1, 10).Value = "Percentage Change"
        ws.Cells(1, 11).Value = "Total Volume"

        ' Initialize variables
        resultRow = 2 ' Assuming your data starts from row 2
        maxIncreasePercentage = 0
        minDecreasePercentage = 0
        greatestVolume = 0

        ' Loop through each row
        For i = 2 To lastRow
            currentTicker = dataRange.Cells(i, 1).Value

            ' Check if the ticker has changed
            If currentTicker <> dataRange.Cells(i - 1, 1).Value Then
                ' Set the currentTicker
                currentTicker = currentTicker

                ' Add to the Total Volume
                totalVolume = totalVolume + dataRange.Cells(i, 7).Value

                ' Print the currentTicker in the Summary Table
                ws.Range("H" & resultRow).Value = currentTicker

                ' Print the Total Volume to the Summary Table
                ws.Range("I" & resultRow).Value = totalVolume

                ' Calculate and display the difference for the previous ticker
                If i > 2 Then
                    priceDifference = lastDayClosingPrice - firstDayOpeningPrice
                    percentageChange = (lastDayClosingPrice / firstDayOpeningPrice) - 1

                    ' Insert the result in a new column (for example, columns J, K, L, M)
                    ws.Cells(resultRow, 8).Value = currentTicker
                    ws.Cells(resultRow, 9).Value = priceDifference
                    ws.Cells(resultRow, 10).Value = Format(percentageChange, "0.00%")
                    ws.Cells(resultRow, 11).Value = totalVolume

                    ' Apply conditional formatting to "Price Difference" column
                    If priceDifference > 0 Then
                        ws.Cells(resultRow, 9).Interior.Color = RGB(144, 238, 144) ' Green fill for positive change
                    ElseIf priceDifference < 0 Then
                        ws.Cells(resultRow, 9).Interior.Color = RGB(255, 99, 71) ' Red fill for negative change
                    End If

                    ' Apply conditional formatting to "Percentage Change" column
                    If percentageChange > 0 Then
                        ws.Cells(resultRow, 10).Interior.Color = RGB(144, 238, 144) ' Green fill for positive change
                    ElseIf percentageChange < 0 Then
                        ws.Cells(resultRow, 10).Interior.Color = RGB(255, 99, 71) ' Red fill for negative change
                    End If

                    ' Check for the Greatest % Increase, Greatest % Decrease, and Greatest Total Volume
                    If percentageChange > maxIncreasePercentage Then
                        maxIncreasePercentage = percentageChange
                        tickerMaxIncrease = currentTicker
                    End If

                    If percentageChange < minDecreasePercentage Then
                        minDecreasePercentage = percentageChange
                        tickerMinDecrease = currentTicker
                    End If

                    If totalVolume > greatestVolume Then
                        greatestVolume = totalVolume
                        tickerGreatestVolume = currentTicker
                    End If

                    resultRow = resultRow + 1 ' Move to the next row for the result
                End If

                ' Reset the Total Volume
                totalVolume = 0

                ' Update currentTicker and reset variables
                firstDayOpeningPrice = dataRange.Cells(i, 3).Value ' Opening price on the first day
                lastDayClosingPrice = dataRange.Cells(i, 6).Value ' Closing price on the last day
            Else
                ' Add to the Total Volume
                totalVolume = totalVolume + dataRange.Cells(i, 7).Value

                ' Update last day closing price for the currentTicker
                lastDayClosingPrice = dataRange.Cells(i, 6).Value
            End If
        Next i

        ' Calculate and display the difference for the last ticker in the table
        If currentTicker <> "" Then
            priceDifference = lastDayClosingPrice - firstDayOpeningPrice
            percentageChange = (lastDayClosingPrice / firstDayOpeningPrice) - 1

            ' Insert the result in a new column (for example, columns J, K, L, M)
            ws.Cells(resultRow, 8).Value = currentTicker
            ws.Cells(resultRow, 9).Value = priceDifference
            ws.Cells(resultRow, 10).Value = Format(percentageChange, "0.00%")
            ws.Cells(resultRow, 11).Value = totalVolume

            ' Apply conditional formatting to "Price Difference" column
            If priceDifference > 0 Then
                ws.Cells(resultRow, 9).Interior.Color = RGB(144, 238, 144) ' Green fill for positive change
            ElseIf priceDifference < 0 Then
                ws.Cells(resultRow, 9).Interior.Color = RGB(255, 99, 71) ' Red fill for negative change
            End If

            ' Apply conditional formatting to "Percentage Change" column
            If percentageChange > 0 Then
                ws.Cells(resultRow, 10).Interior.Color = RGB(144, 238, 144) ' Green fill for positive change
            ElseIf percentageChange < 0 Then
                ws.Cells(resultRow, 10).Interior.Color = RGB(255, 99, 71) ' Red fill for negative change
            End If
        End If

        ' Display the results for "Greatest % Increase," "Greatest % Decrease," and "Greatest Total Volume"
        ws.Cells(1, 13).Value = "Result Type"
        ws.Cells(1, 14).Value = "Ticker"
        ws.Cells(1, 15).Value = "Value"

        ws.Cells(2, 13).Value = "Greatest % Increase"
        ws.Cells(2, 14).Value = tickerMaxIncrease
        ws.Cells(2, 15).Value = Format(maxIncreasePercentage, "0.00%")

        ws.Cells(3, 13).Value = "Greatest % Decrease"
        ws.Cells(3, 14).Value = tickerMinDecrease
        ws.Cells(3, 15).Value = Format(minDecreasePercentage, "0.00%")

        ws.Cells(4, 13).Value = "Greatest Total Volume"
        ws.Cells(4, 14).Value = tickerGreatestVolume
        ws.Cells(4, 15).Value = greatestVolume
    Next ws
End Sub

