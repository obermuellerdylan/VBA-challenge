Sub Stock_Ticker()

For Each ws in Worksheets

    dim SummaryTableRow As Integer
        SummaryTableRow = 1
    dim YearlyChange As Double
    dim TotalVolume 
    dim LastRow As Long
        LastRow = ws.Cells(Rows.Count,1).End(xlUp).Row
   
    
    
    'Cell Formatting
    ws.Cells(1,9) = "Ticker"
    ws.Cells(1,10) = "Yearly Change"
    ws.Cells(1,11) = "Percent Change"
    ws.Cells(1,12) = "Total Volume"
    ws.Cells(1,16) = "Ticker"
    ws.Cells(1,17) = "Value"
    ws.Cells(2, 15) = "Greatest % Increase"
    ws.Cells(3, 15) = "Greatest % Decrease"
    ws.Cells(4, 15) = "Greatest Total Volume"
    ws.Range("O:O").ColumnWidth = 22
    ws.Range("J:L").ColumnWidth = 17
    ws.Range("Q:Q").ColumnWidth = 17

    
    For i = 2 to LastRow

        

        'Finding Opening Value
        Dim OpeningValue As Double
        If ws.Cells(i-1,1) <> ws.Cells(i,1) Then
            OpeningValue = ws.Cells(i,3)
        End If 

        If ws.Cells(i+1,1) <> ws.Cells(i,1) Then

            'Ticker Name Placement  
            SummaryTableRow = SummaryTableRow + 1
            ws.Cells(SummaryTableRow, 9) = ws.Cells(i,1)

            'Adding up TotalVolume
            TotalVolume = TotalVolume + ws.Cells(i,7)
            ws.Cells(SummaryTableRow,12) = TotalVolume
        
            'Reset TotalVolume Count
            TotalVolume = 0
        
            'Finding Closing Value
            Dim ClosingValue As Double
            ClosingValue = ws.Cells(i,6)

            'Placing Yearly Change and Percent Change
            YearlyChange = (ClosingValue - OpeningValue)
            ws.Cells(SummaryTableRow, 10) = YearlyChange
        
            If OpeningValue <> 0 Then
                ws.Cells(SummaryTableRow, 11) = (YearlyChange / OpeningValue)
                ws.Range("K:K").NumberFormat = "0.00%"
            End If 

        Else
            'Adding up TotalVolume when ticker = ticker
            TotalVolume = TotalVolume + ws.Cells(i,7)
        
        End If

        
    Next i

    'Finding Highest % Increase
    MaxChange = Application.WorksheetFunction.Max(ws.Range("K:K"))
    ws.Range("Q2") = MaxChange

    'Finding Highest % Decrease
    MinChange = Application.WorksheetFunction.Min(ws.Range("K:K"))
    ws.Range("Q3") = MinChange
    ws.Range("Q2:Q3").NumberFormat = "0.00%"

    'Finding Highest Total Volume
    MaxTotal = Application.WorksheetFunction.Max(ws.Range("L:L"))
    ws.Range("Q4") = MaxTotal

    'Finding last row in summary table
    dim TableRow 
    TableRow = ws.Cells(Rows.Count,9).End(xlUp).Row

    For i = 2 to TableRow

        'Placing Ticker name next to each value
        If ws.Cells(i, 11) = MaxChange Then
            ws.Range("P2") = ws.Cells(i, 9)
        ElseIf ws.Cells(i,11) = MinChange Then
            ws.Range("P3") = ws.Cells(i, 9)
        ElseIf ws.Cells(i,12) = MaxTotal Then
            ws.Range("P4") = ws.Cells(i, 9)
        End If 
    
        'Conditional Formatting Yearly Change
        If ws.Cells(i, 10) < 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 3
        ElseIf ws.Cells(i,10) >= 0 Then 
            ws.Cells(i, 10).Interior.ColorIndex = 4
        End If 

    Next i

Next ws

End Sub 