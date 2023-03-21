Attribute VB_Name = "stockcode"
Sub stockAnalysis()
    
    Dim ws As Worksheet
    Dim tickerSymbol As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    Dim outputRow As Long
    Dim maxPercentIncrease As Double
    Dim maxPercentDecrease As Double
    Dim maxTotalVolume As Double
    Dim maxPercentIncreaseTicker As String
    Dim maxPercentDecreaseTicker As String
    Dim maxTotalVolumeTicker As String
    
   ' Loop through each worksheet in the workbook
    For Each ws In Worksheets
        
    
        ' Find last row of data
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        ' Set up output headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"
    
        ' Initialize output row counter
        outputRow = 2
    
        ' Loop through each row of data
        For i = 2 To lastRow
        
            ' Check if we've moved to a new ticker symbol
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            
                ' Output the data for the previous ticker symbol
                ws.Cells(outputRow, 9).Value = tickerSymbol
                ws.Cells(outputRow, 10).Value = yearlyChange
                ws.Cells(outputRow, 11).Value = percentChange
                ws.Cells(outputRow, 12).Value = totalVolume
            
                ' Reset variables for the new ticker symbol
                tickerSymbol = ws.Cells(i, 1).Value
                openingPrice = ws.Cells(i, 3).Value
                totalVolume = 0
                outputRow = outputRow + 1
            
            End If
        
            ' Calculate total volume for the ticker symbol
            totalVolume = totalVolume + ws.Cells(i, 7).Value
        
            ' Check if we've reached the last row of data for the current ticker symbol
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
                ' Calculate yearly change and percent change for the ticker symbol
                closingPrice = ws.Cells(i, 6).Value
                yearlyChange = closingPrice - openingPrice
                percentChange = yearlyChange / openingPrice
            
            End If
        
        Next i
        
        maxPercentIncrease = WorksheetFunction.Max(ws.Range("K:K"))
        maxPercentDecrease = WorksheetFunction.Min(ws.Range("K:K"))
        maxTotalVolume = WorksheetFunction.Max(ws.Range("L:L"))
        
        
        
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(1, 15).Value = "Ticker"
        
        ' Search for the value "ABC" in column A
        Dim searchVal As Variant
        searchVal = maxPercentIncrease
        searchVal2 = maxPercentDecrease
        searchVal3 = maxTotalVolume
        Dim rowNum As Long
        Dim rowNum2 As Long
        Dim rowNum3 As Long
        
        rowNum = Application.Match(searchVal, ws.Range("K:K"), 0)
        rowNum2 = Application.Match(searchVal2, ws.Range("K:K"), 0)
        rowNum3 = Application.Match(searchVal3, ws.Range("L:L"), 0)
        
        ws.Cells(2, 15).Value = ws.Cells(rowNum, 9)
        ws.Cells(3, 15).Value = ws.Cells(rowNum2, 9)
        ws.Cells(4, 15).Value = ws.Cells(rowNum3, 9)
        
        ws.Cells(2, 16).Value = maxPercentIncrease
        ws.Cells(3, 16).Value = maxPercentDecrease
        ws.Cells(4, 16).Value = maxTotalVolume
        
    Next ws
    
End Sub
