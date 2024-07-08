# 02_VBAChallenge

Code Solution:
Sub StockAnalysis()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summaryRow As Integer
    Dim i As Long
    
    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Initialize variables
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        summaryRow = 2
        totalVolume = 0
        
        ' Add headers to the summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Get the opening price for the first stock
        openPrice = ws.Cells(2, 3).Value
        
        ' Loop through all rows
        For i = 2 To lastRow
            ' Check if we're still within the same stock
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Set the ticker symbol
                ticker = ws.Cells(i, 1).Value
                
                ' Get the closing price
                closePrice = ws.Cells(i, 6).Value
                
                ' Calculate yearly change
                yearlyChange = closePrice - openPrice
                
                ' Calculate percent change
                If openPrice <> 0 Then
                    percentChange = yearlyChange / openPrice
                Else
                    percentChange = 0
                End If
                
                ' Add to the total volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                
                ' Output the results
                ws.Cells(summaryRow, 9).Value = ticker
                ws.Cells(summaryRow, 10).Value = yearlyChange
                ws.Cells(summaryRow, 11).Value = percentChange
                ws.Cells(summaryRow, 12).Value = totalVolume
                
                ' Format the percent change as a percentage
                ws.Cells(summaryRow, 11).NumberFormat = "0.00%"
                
                ' Color code the yearly change
                If yearlyChange > 0 Then
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(0, 255, 0)
                ElseIf yearlyChange < 0 Then
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(255, 0, 0)
                End If
                
                ' Reset variables for next stock
                summaryRow = summaryRow + 1
                totalVolume = 0
                openPrice = ws.Cells(i + 1, 3).Value
            Else
                ' Add to the total volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
        Next i
        
        ' Autofit the columns in the summary table
        ws.Columns("I:L").AutoFit
    Next ws
    
End Sub