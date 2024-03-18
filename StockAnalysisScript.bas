Attribute VB_Name = "Module1"

Sub StockResults()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summaryRow As Integer

    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    
    For Each ws In ThisWorkbook.Sheets
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Volume"
    
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    summaryRow = 2

    For i = 2 To lastRow
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Or i = 2 Then
            ticker = ws.Cells(i, 1).Value
            openingPrice = ws.Cells(i, 3).Value
            totalVolume = 0
            closingPrice = 0
        End If

        totalVolume = totalVolume + ws.Cells(i, 7).Value
        
        If ws.Cells(i + 1, 1).Value <> ticker Or i = lastRow Then
            closingPrice = ws.Cells(i, 6).Value
            yearlyChange = closingPrice - openingPrice
            ws.Cells(summaryRow, 11).NumberFormat = "0.00"
                If yearlyChange > 0 Then
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(0, 255, 0)
                ElseIf yearlyChange < 0 Then
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(255, 0, 0)
                End If
            If openingPrice <> 0 Then
                percentChange = (yearlyChange / openingPrice)
                ws.Cells(summaryRow, 11).NumberFormat = "0.00%"
                If percentChange > 0 Then
                    ws.Cells(summaryRow, 11).Interior.Color = RGB(0, 255, 0)
                ElseIf yearlyChange < 0 Then
                    ws.Cells(summaryRow, 11).Interior.Color = RGB(255, 0, 0)
                End If
                 
            Else
                percentChange = 0
            End If
            ws.Cells(summaryRow, 9).Value = ticker
            ws.Cells(summaryRow, 10).Value = yearlyChange
            ws.Cells(summaryRow, 11).Value = percentChange
            ws.Cells(summaryRow, 12).Value = totalVolume
 
            If percentChange > greatestIncrease Then
                greatestIncrease = percentChange
                greatestIncreaseTicker = ticker
            End If
                
            If percentChange < greatestDecrease Then
                greatestDecrease = percentChange
                greatestDecreaseTicker = ticker
            End If
                
            If totalVolume > greatestVolume Then
                greatestVolume = totalVolume
                greatestVolumeTicker = ticker
            End If
                summaryRow = summaryRow + 1
            End If
        Next i
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = greatestIncreaseTicker
        ws.Cells(2, 17).Value = greatestIncrease
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = greatestDecreaseTicker
        ws.Cells(3, 17).Value = greatestDecrease
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = greatestVolumeTicker
        ws.Cells(4, 17).Value = greatestVolume

    Next ws
    

End Sub
