' Credit
    ' Instructor Eli Rosenberg tutored a group of us with starter code
    Sub Stocks()
    Dim lastrow As Long
    Dim i As Long
    Dim count As Long
    Dim opening As Double
    Dim closing As Double
    Dim volume As Double
    Dim ticker As String
    Dim percent_change As Double
    
    For Each ws In Worksheets
        lastrow = ws.Cells(Rows.count, 1).End(xlUp).Row ' end row
        count = 2 'start row
        volume = 0
        opening = ws.Cells(2, 3).Value
        ticker = ws.Cells(2, 1).Value
        
        'header row for first output columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        For i = 2 To lastrow
            'store stock volume sum as a cumulative sum
            volume = volume + ws.Cells(i, 7).Value
            
            'run the below if loop until ticker changes over
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'store the last close price before ticker changes
                closing = ws.Cells(i, 6).Value
                percent_change = (closing - opening) / opening
                
                'store the current ticker and current ticker yearly proce change before changing over to the next ticker
                ws.Cells(count, 9).Value = ticker
                ws.Cells(count, 10).Value = closing - opening
                
                ' indicate color diff for price increase vs. price decrease
                If (closing - opening > 0) Then
                    ws.Cells(count, 10).Interior.ColorIndex = 4
                ElseIf (closing - opening < 0) Then
                    ws.Cells(count, 10).Interior.ColorIndex = 3
                End If
                
                'populate the percent change of each ticker display in percentage format
                ws.Cells(count, 11).Value = percent_change
                ws.Cells(count, 11).NumberFormat = "0.00%"
                'populate the ticker volume
                ws.Cells(count, 12).Value = volume
                
                opening = ws.Cells(i + 1, 3).Value
                ticker = ws.Cells(i + 1, 1).Value
                count = count + 1
                volume = 0
            End If
        Next i
    Next ws
End Sub