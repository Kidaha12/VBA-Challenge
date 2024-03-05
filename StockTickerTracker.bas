Attribute VB_Name = "StockTickerTracker"
Sub StockTickerTracker()
    Dim ws As Worksheet
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets


'Turn off screen updating to speed up process
Application.ScreenUpdating = False

'set variables
 
    Dim lastrow As Long
    Dim i As Long
    Dim counter As Long
    Dim percentMaxTicker As String
    Dim percentMinTicker As String
    Dim volumeMaxTicker As String
    Dim sum As Double
    Dim annualChange As Double
    Dim percentMin As Double
    Dim percentMax As Double
    Dim volumeMax As Double
    Dim priceFlag As Boolean
    
    ' Initialize variables before for loop, and across wks.
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    counter = 2
    sum = 0
    percentMin = 1E+99
    percentMax = -1E+99
    volumeMax = -1E+99
    priceFlag = True
    
    For i = 2 To lastrow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ' Generate the unique ticker symbol from column A onto column I.
            ws.Cells(counter, 9).Value = ws.Cells(i, 1).Value
            ' Calculate Yearly Change and save in column J, and highlight cell red for negative or green for positive changes.
            closePrice = ws.Cells(i, 6).Value
            annualChange = closePrice - openPrice
            ws.Cells(counter, 10).Value = annualChange
            If annualChange > 0 Then
                ws.Cells(counter, 10).Interior.ColorIndex = 4
                ws.Cells(counter, 11).Interior.ColorIndex = 4
            ElseIf annualChange < 0 Then
                ws.Cells(counter, 10).Interior.ColorIndex = 3
                ws.Cells(counter, 11).Interior.ColorIndex = 3
            End If
            ' Calculate percent change and save in column K.
            If annualChange = 0 Or openPrice = 0 Then
                ws.Cells(counter, 11).Value = 0
            Else
                ws.Cells(counter, 11).Value = Format(annualChange / openPrice, "#.##%")
            End If
            ' Annual Total Volume per ticker in column L.
            sum = sum + ws.Cells(i, 7).Value
            ws.Cells(counter, 12).Value = sum
            ' Find the values for greatest decrease/increase and greatest volume.
            If ws.Cells(counter, 11).Value > percentMax Then
                If ws.Cells(counter, 11).Value = ".%" Then
                Else
                    percentMax = ws.Cells(counter, 11).Value
                    percentMaxTicker = ws.Cells(counter, 9).Value
                End If
            ElseIf ws.Cells(counter, 11).Value < percentMin Then
                percentMin = ws.Cells(counter, 11).Value
                percentMinTicker = ws.Cells(counter, 9).Value
            ElseIf ws.Cells(counter, 12).Value > volumeMax Then
                volumeMax = ws.Cells(counter, 12).Value
                volumeMaxTicker = ws.Cells(counter, 9).Value
            End If
            ' Rinse and repeat for the next ticker symbol.
            counter = counter + 1
            sum = 0
            priceFlag = True
        Else
            ' Use Pflag to save the open price value at the start of the year.
            If priceFlag Then
                openPrice = ws.Cells(i, 3).Value
                priceFlag = False
            End If
            ' If adjacent ticker symbols are the same, then save volume value to sum total annual volume.
            sum = sum + ws.Cells(i, 7).Value
        End If
    Next i
    ' Fill in the values for greatest increase&decrease and greatest volume.
    ws.Cells(2, 17).Value = Format(percentMax, "#.##%")
    ws.Cells(3, 17).Value = Format(percentMin, "#.##%")
    ws.Cells(4, 17).Value = volumeMax
    ' Fill in the column headers names.
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Volume"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    ' Place corresponding ticker symbol to weighted change values.
    ws.Cells(2, 16).Value = percentMaxTicker
    ws.Cells(3, 16).Value = percentMinTicker
    ws.Cells(4, 16).Value = volumeMaxTicker
    ' Apply autofit to adjust column widths
    ws.Columns.AutoFit
Next ws
    
End Sub
