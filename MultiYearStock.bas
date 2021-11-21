Attribute VB_Name = "MultiYearStock"

Sub run_all_worksheets()
' Main sub to run the entire program

    Dim ws As Worksheet
    Dim first_ws As Worksheet
    Set first_ws = ActiveSheet 'set a current worksheet as first active worksheet in the beginning
    
    ' Repeat calculating for all worksheets in the current workbook
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        Call get_ticker_summary
        Call format_table
    Next
    
    first_ws.Activate 'reactivate the worksheet

End Sub


Sub get_ticker_summary()

    ' Declare and initialise all required variables
    Dim tickerName As String
    Dim inputRow, outputRow, lastRow As Long
    Dim totalVolume As Currency: totalVolume = 0
    
    Dim openPrice As Double: openPrice = Cells(2, 3).Value
    Dim closePrice As Double: closePrice = 0
    Dim priceDiff As Double: priceDiff = 0
    Dim priceDiff_percent As Double: priceDiff_percent = 0

    Dim highestTicker As String: highestTicker = " "
    Dim lowestTicker As String: lowestTicker = " "
    Dim maxPercent As Double: maxPercent = 0
    Dim minPercent As Double: minPercent = 0
    Dim highestVolumeTicker As String: highestVolumeTicker = " "
    Dim highestVolume As Double: highestVolume = 0

    ' Get last row count of a worksheet
    lastRow = ActiveWorkbook.ActiveSheet.Cells(ActiveWorkbook.ActiveSheet.Rows.Count, 1).End(xlUp).Row
  
    ' Start to work on row 2 of a worksheet
    outputRow = 2
  
    ' Starting at 2nd row (i.e. ignore column header row), loop through all rows in the columns
    For inputRow = 2 To lastRow
                
        ' Check when the value of the next cell is different than that of the current cell
        If Cells(inputRow + 1, 1).Value <> Cells(inputRow, 1).Value Then
            
            ' Get the ticker name and volume
            tickerName = Cells(inputRow, 1).Value
            totalVolume = totalVolume + Cells(inputRow, 7).Value
        
            ' Populate data in column I - for ticker symbol
            Cells(outputRow, 9).Value = Cells(inputRow, 1).Value
            
            ' Calculate yearly priceDiff and priceDiff_percent from price of first day openning to last day closing
            closePrice = Cells(inputRow, 6).Value
            priceDiff = closePrice - openPrice
            
            ' Populate data in column J - for price change
            Cells(outputRow, 10).Value = priceDiff
            
            ' Capture the scenario when a price divides by 0
            If openPrice <> 0 Then
                priceDiff_percent = (priceDiff / openPrice) * 100
            End If
                        
            ' Populate data in column K and format the cells with percentage - for yearly percentage price change
            Cells(outputRow, 11) = priceDiff_percent
            Cells(outputRow, 11).NumberFormat = "0.00\%"
            
            'Format cell color with green - for positive value; or red - for negative value
            If (priceDiff > 0) Then
                Cells(outputRow, 10).Interior.ColorIndex = 4
            ElseIf (priceDiff <= 0) Then
                Cells(outputRow, 10).Interior.ColorIndex = 3
            End If
                    
            ' Populate data in column L - for total stock volume
            Cells(outputRow, 12).Value = Str(totalVolume)
            
            ' Set next row as the current row
            outputRow = outputRow + 1

            ' Capture the ticker with the highest increase & decrease value
            If (priceDiff_percent > maxPercent) Then
                maxPercent = priceDiff_percent
                highestTicker = tickerName
            ElseIf (priceDiff_percent < minPercent) Then
                minPercent = priceDiff_percent
                lowestTicker = tickerName
            End If
                
            ' Capture the ticker with the highest volume
            If (totalVolume > highestVolume) Then
                highestVolume = totalVolume
                highestVolumeTicker = tickerName
            End If
            
            ' Reset variables
            openPrice = Cells(inputRow + 1, 3).Value
            closePrice = 0
            priceDiff = 0
            priceDiff_percent = 0
            totalVolume = 0
        
        Else
            'Calculate total volume of stock of a ticker
            totalVolume = totalVolume + Cells(inputRow, 7).Value
            
        End If

    Next inputRow
    
    ' Populate data in column O to Q
    Range("Q2").Value = maxPercent
    Range("Q2").NumberFormat = "0.00\%"
    Range("Q3").Value = minPercent
    Range("Q3").NumberFormat = "0.00\%"
    Range("P2").Value = highestTicker
    Range("P3").Value = lowestTicker
    Range("Q4").Value = highestVolume
    Range("P4").Value = highestVolumeTicker

End Sub


Sub format_table()
        
    'Label headers for the summary columns & fields
    Range("I1").Value = "Ticker Symbol"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("I1:Q1").Font.Bold = True
    Range("I1:L1").Interior.ColorIndex = 15
    Range("P1:Q1").Interior.ColorIndex = 15
    
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("O2:O4").Font.Bold = True
    Range("O2:O4").Interior.ColorIndex = 15

    'Autofit all columns
    Columns("A:Q").AutoFit

End Sub


Sub clean_all_sheets_data()
' Remove all output & formatted columns
    
    Dim ws As Worksheet
    Dim first_ws As Worksheet
    Set first_ws = ActiveSheet 'set the current worksheet as first active worksheet in the beginning

    ' Repeat deletion for all worksheets in the current workbook
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        Range("I:Q").Delete
    Next
    
    first_ws.Activate
    
End Sub
