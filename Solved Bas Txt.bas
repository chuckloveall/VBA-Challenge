Sub stock_ticker():

'define variables
Dim openval As Double
Dim closer As Double
Dim diff As Double
Dim counter As LongLong
Dim ticker As String
Dim ticker2 As String
Dim ticker3 As String
Dim ticker4 As String
Dim percentage As Double
Dim max As Double
Dim maxv As LongLong
Dim min As Double

'initialize counter variables at 0
 max = Cells(2, 11).Value
 min = Cells(2, 11).Value
maxv = Cells(2, 12).Value
'loop through all sheets
For Each ws In Worksheets

    'insert summary table to house results and headers
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2


    'Insert Summary Table Headers
    ws.Cells(1, 9).Value = "Ticker Symbol"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    'Extra summary headers
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    'Autofit Column Header width
    '??ActiveSheet v ws?? which to use
    ws.Range("I:L").EntireColumn.AutoFit
    'find last filled row
    RowCount = ws.Cells(Rows.Count, 1).End(xlUp).Row
    'initialize openval value
    openval = ws.Cells(2, 3).Value
    For i = 2 To RowCount
    'determine when change happens in column A (ticker)
    'happens when we reach end of the ticker
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        'grab closing value after change
        closer = ws.Cells(i, 6).Value

        'yearly change is difference between open and close price
        diff = closer - openval
        'Debug.Print (diff)
        'Debug.Print (closer)


        'edge case of percentage =0
            If openval = 0 Then
            ws.Cells(i, 11).Value = 0
            Else
            'Calculate percantage change
            percentage = diff / openval


            End If

        'Debug.Print (percentage)
        ws.Range("K" & Summary_Table_Row).Value = percentage
        'if percentage >0 then
        'color green
        'else
        'color red
        'end if

        ws.Range("J" & Summary_Table_Row).Value = diff
        'grab new openval
        openval = ws.Cells(i + 1, 3).Value
        'grab ticker value and assign to ticker symbol
        ticker = ws.Cells(i, 1).Value
        ws.Range("I" & Summary_Table_Row).Value = ticker
        'assign counter sum to total stock volume
        ws.Range("L" & Summary_Table_Row).Value = counter
        'reset counter back to zero to begin for next ticket
        counter = 0
        'add summary table row
        Summary_Table_Row = Summary_Table_Row + 1

    'sum each vol of ticker until changed
        Else
        'add total volume
        counter = counter + ws.Cells(i, 7).Value

        End If

    Next i
    'Conditional formating to color yearly change based on positive or negative value
    'find end of summary table row
    EndSummary = ws.Cells(Rows.Count, 11).End(xlUp).Row
    'format column K to percentage format
    ws.Range("K2:K" & EndSummary).NumberFormat = "0.00%"
    'iterate through values
    For j = 2 To EndSummary
        If ws.Cells(j, 10).Value > 0 Then
        'green
            ws.Cells(j, 10).Interior.ColorIndex = 4
        Else
        'red
            ws.Cells(j, 10).Interior.ColorIndex = 3

    '?? do we need to worry about edge case of value=0 ??

        End If
    Next j

        'check value against max value
        For k = 2 To EndSummary
            If Cells(k, 11).Value > max Then
            'if cell value is greater than max, then assign new value to max and move to next k
            max = Cells(k, 11).Value
            'grab text (ticker) from answer
            ticker2 = Cells(k, 9).Value

            'check value against min value
            ElseIf Cells(k, 11).Value < min Then
            min = Cells(k, 11).Value
            'grab text (ticker) from answer
            ticker3 = Cells(k, 9).Value

            End If
        Next k

        For l = 2 To EndSummary
         If Cells(l, 12).Value > maxv Then
            maxv = Cells(l, 12).Value
            ticker4 = Cells(l, 9).Value
        End If
        Next l

     'assign values to greatest increase/decrease and total volume
        ws.Range("P2").Value = max
        ws.Range("P2").NumberFormat = "0.00%"
        ws.Range("P3").Value = min
        ws.Range("P3").NumberFormat = "0.00%"
        ws.Range("P4").Value = maxv
        ws.Range("O2").Value = ticker2
        ws.Range("O3").Value = ticker3
        ws.Range("O4").Value = ticker4

Next ws

End Sub
