Attribute VB_Name = "Module1"
Sub stocks()
    ' Variables
    Dim ws As Worksheet
    Dim row As Long
    Dim rowCount As Long
    Dim tickerRow As Long
    Dim total As Double
    Dim startPrice As Double
    Dim endPrice As Double
    Dim priceChange As Double
    Dim percentChange As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    Dim greatestIncreaseValue As Double
    Dim greatestDecreaseValue As Double
    Dim greatestVolumeValue As Double

    ' Loop through each worksheet (tab) in the Excel file
    For Each ws In ThisWorkbook.Worksheets
        ' Add title rows
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"

        ' Set initial values for ticker data
        tickerRow = 2
        total = 0
        startPrice = 0

        ' Set initial values for summary data table
        greatestIncreaseTicker = ""
        greatestDecreaseTicker = ""
        greatestVolumeTicker = ""
        greatestIncreaseValue = 0
        greatestDecreaseValue = 0
        greatestVolumeValue = 0

        ' Get the row number of the last data row
        rowCount = Cells(Rows.Count, "A").End(xlUp).row

        For row = 2 To rowCount
            ' Get the total volume
            total = total + Cells(row, 7).Value

            ' Find first nonzero start
            If startPrice = 0 Then
                startPrice = Cells(row, 3).Value
            End If

            ' Once the ticker changes, calculate totals and print stats
            If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
                ' Add ticker to the first column of the Summary row
                Cells(tickerRow, 9).Value = Cells(row, 1).Value

                ' Get Stats for Summary Table
                If total = 0 Then
                    ' Give results
                    Cells(tickerRow, 10).Value = 0
                    Cells(tickerRow, 11).Value = "%" & 0
                    Cells(tickerRow, 12).Value = 0
                Else
                    ' Calculate Yearly Summaries
                    endPrice = Cells(row, 6).Value
                    priceChange = endPrice - startPrice
                    percentChange = priceChange / startPrice

                    ' Print Yearly Summary Stats for each ticker
                    ws.Cells(tickerRow, 10).Value = priceChange
                    ws.Cells(tickerRow, 10).NumberFormat = "0.00"
                    ws.Cells(tickerRow, 11).Value = percentChange
                    ws.Cells(tickerRow, 11).NumberFormat = "0.00%"
                    ws.Cells(tickerRow, 12).Value = total
                    ws.Cells(tickerRow, 12).NumberFormat = "#,###"

                    ' Color code positive and negative
                    If priceChange > 0 Then
                        Cells(tickerRow, 10).Interior.ColorIndex = 4
                    ElseIf priceChange < 0 Then
                        Cells(tickerRow, 10).Interior.ColorIndex = 3
                    End If

                    ' Check if ticker has the greatest total volume
                    If total > greatestVolumeValue Then
                        greatestVolumeValue = total
                        greatestVolumeTicker = Cells(row, 1).Value
                    End If

                    ' Find if it has the greatest % increase or decrease so far
                    If percentChange > greatestIncreaseValue Then
                        greatestIncreaseValue = percentChange
                        greatestIncreaseTicker = Cells(row, 1).Value
                    ElseIf percentChange < greatestDecreaseValue Then
                        greatestDecreaseValue = percentChange
                        greatestDecreaseTicker = Cells(row, 1).Value
                    End If
                End If

                ' Reset ticker summary counts
                tickerRow = tickerRow + 1
                total = 0
                startPrice = 0
            End If
        Next row

        ' Print standout ticker stats after looping through all
        ws.Cells(2, 16).Value = greatestIncreaseTicker
        ws.Cells(2, 17).Value = greatestIncreaseValue
        ws.Cells(2, 17).NumberFormat = "0.00%"

        ws.Cells(3, 16).Value = greatestDecreaseTicker
        ws.Cells(3, 17).Value = greatestDecreaseValue
        ws.Cells(3, 17).NumberFormat = "0.00%"

        ws.Cells(4, 16).Value = greatestVolumeTicker
        ws.Cells(4, 17).Value = greatestVolumeValue
        ws.Cells(4, 17).NumberFormat = "#,###"
    Next ws
End Sub

