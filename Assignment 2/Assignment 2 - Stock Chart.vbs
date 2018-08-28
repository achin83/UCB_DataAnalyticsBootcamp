'MEDIUM METHOD
'----------------------------------------------------------------------------------
Sub TotalStockVolume():
    'Observe all worksheets in workbook
    For Each ws In Worksheets
        'Declare total volume to accept a value greater than Long
        Dim TotalVolume As Double
        'Store incremental row number in memory to aggregate stock volume by ticker symbolDim SummaryRow As Integer
        Dim SummaryRow As Integer
        'Summary information begins after header on line 2
        SummaryRow = 2
        'Set column headers for summary information
        ws.Range("I1").Value = "Stock Ticker"
        ws.Range("J1").Value = "Opening Year"
        ws.Range("K1").Value = "Opening Price"
        ws.Range("L1").Value = "Closing Year"
        ws.Range("M1").Value = "Closing Price"
        ws.Range("N1").Value = "Yearly Change"
        ws.Range("O1").Value = "Percent Change"
        ws.Range("P1").Value = "Total Volume"
        'Declare last row value for every worksheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            For i = 2 To LastRow
                'If the next value in column 1 isn't the same as the current value in column 1
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    'Set the current sheet's ticker symbol and calculated total volume on incremental row
                    ws.Range("I" & SummaryRow) = ws.Cells(i, 1).Value
                    ws.Range("P" & SummaryRow) = TotalVolume + Cells(i, 7).Value
                    'Increment the summary row for next ticker symbol information
                    SummaryRow = SummaryRow + 1
                    'Reset total volume for next ticker symbol
                    TotalVolume = 0
                'If first condition is not met (next value in colulmn 1 matches current in column 1)
                Else
                    'Increment total volume based on current ticker symbol
                    TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                End If
            Next i
        'Reset SummaryRow to 2 so that I can loop through dataset again starting from first value  
        SummaryRow = 2
        For j = 2 To LastRow
            'Only take the first value of the stock ticker to get the opening price and date information
            If ws.Cells(SummaryRow, 9).Value = ws.Cells(j, 1).Value And IsEmpty(ws.Cells(SummaryRow, 10)) = True Then
                ws.Cells(SummaryRow, 10).Value = ws.Cells(j, 2).Value
                ws.Cells(SummaryRow, 11).Value = ws.Cells(j, 3).Value
            'If the next aggregate value matches the first itemized value, grab the opening price and date information
            ElseIf ws.Cells(SummaryRow + 1, 9).Value = ws.Cells(j, 1).Value Then
                SummaryRow = SummaryRow + 1
                ws.Cells(SummaryRow, 10).Value = ws.Cells(j, 2).Value
                ws.Cells(SummaryRow, 11).Value = ws.Cells(j, 3).Value
            End If
        Next j
        'Reset SummaryRow to 2 so that I can loop through dataset again starting from first value 
        SummaryRow = 2
        For k = 2 To LastRow
            'To get the closing price and date information, loop through itemized list and choose the last value
            'where the aggregate value doesn't match the last itemized value
            If ws.Cells(SummaryRow, 9).Value <> ws.Cells(k + 1, 1).Value And IsEmpty(ws.Cells(SummaryRow, 12)) = True Then
                ws.Cells(SummaryRow, 12).Value = ws.Cells(k, 2).Value
                ws.Cells(SummaryRow, 13).Value = ws.Cells(k, 6).Value
                ws.Cells(SummaryRow, 14).Value = ws.Cells(SummaryRow, 13).Value - ws.Cells(SummaryRow, 11).Value
                'Account for the divide by zero scenario as this throws an error
                If ws.Cells(SummaryRow, 11).Value = 0 Then
                    ws.Cells(SummaryRow, 15).Value = ws.Cells(SummaryRow, 13).Value * 100
                Else
                    ws.Cells(SummaryRow, 15).Value = ws.Cells(SummaryRow, 13).Value / ws.Cells(SummaryRow, 11).Value
                End If
                SummaryRow = SummaryRow + 1
            End If
        Next k
    'Apply formatting
    ws.Range("K2:K" & LastRow & ", " & "M2:N" & LastRow).Style = "Currency"
    ws.Columns("A:P").AutoFit
    Next ws
End Sub
