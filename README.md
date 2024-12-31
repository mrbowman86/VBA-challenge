Sub StockAnalysisLoop()

    'Set worksheet variable
        Dim Worksheet As Worksheet

    'Begin worksheet loop
            For Each Worksheet In ActiveWorkbook.Worksheets
            Worksheet.Activate

    'Setting variables
                Dim TickerSymbol As String
                Dim TickerVolume As Double
                TickerVolume = 0
                Dim TickerSummaryRow As Integer
                TickerSummaryRow = 2
                Dim OpeningPrice As Double
                OpeningPrice = Cells(2, 3).Value
                Dim ClosingPrice As Double
                Dim QuarterlyChange As Double
                Dim PercentageChange As Double
        
    'Setting summary row labels
                Cells(1, 9).Value = "Ticker"
                Cells(1, 10).Value = "Quarterly Change"
                Cells(1, 11).Value = "Percent Change"
                Cells(1, 12).Value = "Total Stock Volume"
            
    'Count the number of rows
                LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
    'Begin looping through rows
                For i = 2 To LastRow
                    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                        TickerSymbol = Cells(i, 1).Value
                        TickerVolume = TickerVolume + Cells(i, 7).Value
                        Range("I" & TickerSummaryRow).Value = TickerSymbol
                        Range("L" & TickerSummaryRow).Value = TickerVolume
                        ClosingPrice = Cells(i, 6).Value
                        QuarterlyChange = (ClosingPrice - OpeningPrice)
                        Range("J" & TickerSummaryRow).Value = QuarterlyChange
               
    'Adjusting instances of dividing by zero
                    If (OpeningPrice = 0) Then
                        PercentageChange = 0
                    Else
                        PercentageChange = QuarterlyChange / OpeningPrice
                    End If
                
    'Add quarterly change to each ticker symbol in summary section
                Range("K" & TickerSummaryRow).Value = PercentageChange
                Range("K" & TickerSummaryRow).NumberFormat = "0.00%"
               
    'Reset row counter
                TickerSummaryRow = TickerSummaryRow + 1
               
    'Reset trade volume
                TickerVolume = 0
               
    'Reset opening price
                OpeningPrice = Cells(i + 1, 3)
               
    'Add ticker volume
                Else
                    TickerVolume = TickerVolume + Cells(i, 7).Value
                End If
            
            Next i
                
    'Find last ticker summary row
            TickerSummaryLastRow = Worksheet.Cells(Rows.Count, 9).End(xlUp).Row
                
    'Add conditional formatting to yearly change column
            TickerSummaryRow = Cells(Rows.Count, 9).End(xlUp).Row
            For j = 2 To TickerSummaryLastRow
                If Cells(j, 10).Value > 0 Then
                    Cells(j, 10).Interior.ColorIndex = 4
                ElseIf Cells(j, 10).Value < 0 Then
                    Cells(j, 10).Interior.ColorIndex = 3
                ElseIf Cells(j, 10).Value = 0 Then
                    Cells(j, 10).Interior.ColorIndex = 2
                End If
        
            Next j
        
    'Create headers with greatest % increase, greatest % decrease, and greatest total volume
            Cells(1, 16).Value = "Ticker"
            Cells(1, 17).Value = "Value"
            Cells(2, 15).Value = "Greatest % Increase"
            Cells(3, 15).Value = "Greatest % Decrease"
            Cells(4, 15).Value = "Greatest Total Volume"
        
    'Return values for greatest % increase, greatest % decrease, and greatest total volume
        
            For k = 2 To TickerSummaryLastRow
                If Cells(k, 11).Value = Application.WorksheetFunction.Max(Worksheet.Range("K2:K" & TickerSummaryLastRow)) Then
                    Cells(2, 16).Value = Cells(k, 9).Value
                    Cells(2, 17).Value = Cells(k, 11).Value
                    Cells(2, 17).NumberFormat = "0.00%"
                ElseIf Cells(k, 11).Value = Application.WorksheetFunction.Min(Worksheet.Range("K2:K" & TickerSummaryLastRow)) Then
                    Cells(3, 16).Value = Cells(k, 9).Value
                    Cells(3, 17).Value = Cells(k, 11).Value
                    Cells(3, 17).NumberFormat = "0.00%"
                ElseIf Cells(k, 12).Value = Application.WorksheetFunction.Max(Worksheet.Range("L2:L" & TickerSummaryLastRow)) Then
                    Cells(4, 16).Value = Cells(k, 9).Value
                    Cells(4, 17).Value = Cells(k, 12).Value
                End If
            
            Next k
        
            Worksheets("Q1").Select
        
        Next Worksheet
        
End Sub
