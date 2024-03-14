' VBA-challenge
' Module 2 Challenge - Due Thursday by 11:59pm

' Project Description:

' Developed a VBA script to analyze stock market data efficiently, providing insights into various stock metrics over the course of a year. The script iterates through all stocks for a given year and generates key information, 'including the ticker symbol, yearly change from opening to closing price, percentage change, and total stock volume.

' Enhanced the functionality of the script to identify stocks with the greatest percentage increase, decrease, and total volume, aiding in the identification of significant market trends and outliers.

' Implemented conditional formatting within the script to visually highlight positive changes in green and negative changes in red, allowing for quick and intuitive data interpretation.

' The script is designed to run seamlessly across multiple worksheets, facilitating the analysis of stock data across different years simultaneously, thus streamlining the analytical process.

' Code Source: The VBA script development process heavily utilized the Xpert Learning Assistant platform, leveraging its vast knowledge base and interactive capabilities. Additionally, some cross-comparisons were made with ' ChatGPT to address specific challenges encountered during script development.

' Below are the VBA scripts for both the "Alphabetical_testing" and "AnalyzeStockData" functionalities:
' 1.) Alphabetical_testing VBA Script
' 2.) AnalyzeStockData VBA Script

' Separate VBA script files Alphabetical_testing
Sub AnalyzeStockData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
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
    Dim greatestVolumeWritten As Boolean
    
    greatestIncrease = -1 ' Set initial values for comparison
    greatestDecrease = 999999 ' Set initial values for comparison
    greatestVolume = 0 ' Set initial values for comparison
    
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Summary" Then
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            summaryRow = 2
            greatestVolumeWritten = False ' Initialize flag to indicate if greatest volume has been written
           
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Yearly Change"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"
           
            openPrice = ws.Cells(2, 3).Value
           
            For i = 2 To lastRow
                If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                    ticker = ws.Cells(i, 1).Value
                    closePrice = ws.Cells(i, 6).Value
                    yearlyChange = closePrice - openPrice
                   
                    If openPrice <> 0 Then
                        percentChange = yearlyChange / openPrice
                    Else
                        percentChange = 0
                    End If
                   
                    totalVolume = Application.WorksheetFunction.SumIf(ws.Range("A:A"), ticker, ws.Range("G:G"))
                   
                    ws.Cells(summaryRow, 9).Value = ticker
                    ws.Cells(summaryRow, 10).Value = yearlyChange
                    ws.Cells(summaryRow, 11).Value = percentChange
                    ws.Cells(summaryRow, 12).Value = totalVolume
                   
                    ' Check for greatest % increase
                    If percentChange > greatestIncrease Then
                        greatestIncrease = percentChange
                        greatestIncreaseTicker = ticker
                    End If
                   
                    ' Check for greatest % decrease
                    If percentChange < greatestDecrease Then
                        greatestDecrease = percentChange
                        greatestDecreaseTicker = ticker
                    End If
                   
                    ' Check for greatest total volume if it's not already written
                    If totalVolume > greatestVolume And Not greatestVolumeWritten Then
                        greatestVolume = totalVolume
                        greatestVolumeTicker = ticker
                        greatestVolumeWritten = True ' Set flag to indicate greatest volume has been written
                    End If
                   
                    summaryRow = summaryRow + 1
                    openPrice = ws.Cells(i + 1, 3).Value
                End If
            Next i
           
            ' Write greatest % increase and greatest % decrease to output destination
            ws.Cells(1, 15).Value = "Summary Statistics"
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"
            ws.Cells(2, 15).Value = "Greatest % Increase"
            ws.Cells(2, 16).Value = greatestIncreaseTicker
            ws.Cells(2, 17).Value = Format(greatestIncrease, "0.00%")
           
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(3, 16).Value = greatestDecreaseTicker
            ws.Cells(3, 17).Value = Format(greatestDecrease, "0.00%")
           
            ws.Cells(4, 15).Value = "Greatest Total Volume"
            ws.Cells(4, 16).Value = greatestVolumeTicker
            ws.Cells(4, 17).Value = greatestVolume
            
            ' Apply conditional formatting to highlight positive change in green and negative change in red
            Dim rng As Range
            Set rng = ws.Range(ws.Cells(2, 11), ws.Cells(lastRow, 11)) ' Column K for percent change
            rng.NumberFormat = "0.00%" ' Format as percentage
            Set rng = ws.Range(ws.Cells(2, 10), ws.Cells(lastRow, 10)) ' Column J for yearly change
            rng.FormatConditions.Delete
            With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="=0")
                .Interior.Color = RGB(0, 255, 0) ' Green for positive change
            End With
            With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0")
                .Interior.Color = RGB(255, 0, 0) ' Red for negative change
            End With
        End If
    Next ws
End Sub

'Separate VBA script files for AnalyzeStockData
Sub AnalyzeStockData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
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
    Dim greatestVolumeWritten As Boolean
    
    greatestIncrease = -1 ' Set initial values for comparison
    greatestDecrease = 999999 ' Set initial values for comparison
    greatestVolume = 0 ' Set initial values for comparison
    
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Summary" Then
            ws.Activate
            lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            summaryRow = 2
            greatestVolumeWritten = False ' Initialize flag to indicate if greatest volume has been written
           
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Yearly Change"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"
           
            openPrice = ws.Cells(2, 3).Value
           
            For i = 2 To lastRow
                If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                    ticker = ws.Cells(i, 1).Value
                    closePrice = ws.Cells(i, 6).Value
                    yearlyChange = closePrice - openPrice
                   
                    If openPrice <> 0 Then
                        percentChange = yearlyChange / openPrice
                    Else
                        percentChange = 0
                    End If
                   
                    totalVolume = Application.WorksheetFunction.SumIf(ws.Range("A:A"), ticker, ws.Range("G:G"))
                   
                    ws.Cells(summaryRow, 9).Value = ticker
                    ws.Cells(summaryRow, 10).Value = yearlyChange
                    ws.Cells(summaryRow, 11).Value = percentChange
                    ws.Cells(summaryRow, 12).Value = totalVolume
                   
                    ' Check for greatest % increase
                    If percentChange > greatestIncrease Then
                        greatestIncrease = percentChange
                        greatestIncreaseTicker = ticker
                    End If
                   
                    ' Check for greatest % decrease
                    If percentChange < greatestDecrease Then
                        greatestDecrease = percentChange
                        greatestDecreaseTicker = ticker
                    End If
                   
                    ' Check for greatest total volume if it's not already written
                    If totalVolume > greatestVolume And Not greatestVolumeWritten Then
                        greatestVolume = totalVolume
                        greatestVolumeTicker = ticker
                        greatestVolumeWritten = True ' Set flag to indicate greatest volume has been written
                    End If
                   
                    summaryRow = summaryRow + 1
                    openPrice = ws.Cells(i + 1, 3).Value
                End If
            Next i
           
            ' Write greatest % increase and greatest % decrease to output destination
            ws.Cells(1, 15).Value = "Summary Statistics"
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"
            ws.Cells(2, 15).Value = "Greatest % Increase"
            ws.Cells(2, 16).Value = greatestIncreaseTicker
            ws.Cells(2, 17).Value = Format(greatestIncrease, "0.00%")
           
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(3, 16).Value = greatestDecreaseTicker
            ws.Cells(3, 17).Value = Format(greatestDecrease, "0.00%")
           
            ' Write greatest total volume if it's not already written
            If Not greatestVolumeWritten Then
                ws.Cells(4, 15).Value = "Greatest Total Volume"
                ws.Cells(4, 16).Value = greatestVolumeTicker
                ws.Cells(4, 17).Value = greatestVolume
            End If
            
            ' Apply conditional formatting to highlight positive change in green and negative change in red
            Dim rng As Range
            Set rng = ws.Range(ws.Cells(2, 10), ws.Cells(lastRow, 10))
            rng.FormatConditions.Delete
            With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="=0")
                .Interior.Color = RGB(0, 255, 0) ' Green for positive change
            End With
            With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0")
                .Interior.Color = RGB(255, 0, 0) ' Red for negative change
            End With
        End If
    Next ws
End Sub
