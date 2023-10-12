VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stock()
    'define all variables
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summaryRow As Long
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    
For Each ws In ThisWorkbook.Worksheets
    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    'create headers on first row for each worksheet
    
    summaryRow = 2 'start writing results from row 2
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    openingPrice = ws.Cells(2, 3).Value 'initialize opening price for each ws
    totalVolume = 0 'initialize total volume for each ws
    For i = 2 To lastRow
        
        If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
            totalVolume = totalVolume + ws.Cells(i, 7).Value
        ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            ticker = ws.Cells(i, 1).Value 'store ticker symbol
            closingPrice = ws.Cells(i, 6).Value 'store closing price
            yearlyChange = closingPrice - openingPrice 'calculate yearly change
            If yearlyChange > 0 Then 'color format for yearly change
                ws.Cells(summaryRow, 10).Interior.ColorIndex = 4
            ElseIf yearlyChange < 0 Then
                ws.Cells(summaryRow, 10).Interior.ColorIndex = 3
            End If
            
            'calculate the percent change and percentage format
            If openingPrice <> 0 Then
                percentChange = yearlyChange / openingPrice
            Else
                percentChange = 0
            End If
            
            ws.Cells(summaryRow, 11).NumberFormat = "0.00%"
            
            'still have to add to total volume
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            'print results into table
            ws.Cells(summaryRow, 9).Value = ticker
            ws.Cells(summaryRow, 10).Value = yearlyChange
            ws.Cells(summaryRow, 11).Value = percentChange
            ws.Cells(summaryRow, 12).Value = totalVolume
            
            'check for greatest percent increase, greatest percent decrease, and greatest total volume
            If percentChange > greatestIncrease Then
                greatestIncrease = percentChange
                greatestIncreaseTicker = ticker
            ElseIf percentChange < greatestDecrease Then
                greatestDecrease = percentChange
                greatestDecreaseTicker = ticker
            End If
            If totalVolume > greatestVolume Then
                greatestVolume = totalVolume
                greatestVolumeTicker = ticker
            End If

            
            'move to the next summary row
            summaryRow = summaryRow + 1
            
            'set new opening price
            openingPrice = ws.Cells(i + 1, 3).Value

            'reset total volume for new ticker
            totalVolume = 0
        End If
        
    Next i
    ws.Cells(2, 16).Value = greatestIncreaseTicker
    ws.Cells(2, 17).Value = greatestIncrease
    ws.Cells(3, 16).Value = greatestDecreaseTicker
    ws.Cells(3, 17).Value = greatestDecrease
    ws.Cells(4, 16).Value = greatestVolumeTicker
    ws.Cells(4, 17).Value = greatestVolume
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"
    ws.Columns("J:L").AutoFit 'autofit column
    ws.Columns("O:Q").AutoFit
Next ws
End Sub

