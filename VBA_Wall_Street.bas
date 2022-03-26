Attribute VB_Name = "Module1"
Sub stock_summary()

For Each ws In Worksheets
    'label top row in each worksheet
    ws.Range("i1").Value = "Ticker"
    ws.Range("j1").Value = "Yearly Change"
    ws.Range("k1").Value = "Percent Change"
    ws.Range("l1").Value = "Total Stock Volume"
    
    'find last row
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim ticker_first_row, ticker_last_row, total_volume, summary_row, summary_count As Long
    Dim ticker_name As String
    ticker_first_row = 2
    summary_row = 2
    total_volume = 0
    summary_count = 0
    
    For i = 2 To lastrow
        tickername = ws.Cells(i, 1).Value
        total_volume = total_volume + ws.Cells(i, 7).Value
        
        'find last row of individual ticker
        If Not (tickername = ws.Cells(i + 1, 1)) Then
            ticker_last_row = i
            'summary ticker name
            ws.Cells(summary_row, 9).Value = tickername
            
            'summary yearly change
            ws.Cells(summary_row, 10).Value = ws.Cells(ticker_last_row, 6).Value - ws.Cells(ticker_first_row, 3).Value
            
            'summary percent change
            ws.Cells(summary_row, 11).Value = Str(Round(100 * ws.Cells(summary_row, 10).Value / ws.Cells(ticker_first_row, 3), 2)) + "%"
            If (ws.Cells(summary_row, 10).Value < 0) Then
                ws.Cells(summary_row, 10).Interior.Color = RGB(255, 0, 0)
            ElseIf (ws.Cells(summary_row, 10).Value > 0) Then
                ws.Cells(summary_row, 10).Interior.Color = RGB(0, 255, 0)
            End If
            'summary stock volume
            ws.Cells(summary_row, 12).NumberFormat = "#,##0"
            ws.Cells(summary_row, 12).Value = total_volume
            
            'reset variables for next ticker
            total_volume = 0
            ticker_first_row = i + 1
            summary_row = summary_row + 1
            summary_count = summary_count + 1
            
        End If
    
    Next i
    
    'annual report labels
    ws.Range("p1").Value = "Ticker"
    ws.Range("q1").Value = "Value"
    ws.Range("o2").Value = "Greatest % Increase"
    ws.Range("o3").Value = "Greatest % Decrease"
    ws.Range("o4").Value = "Greatest Total Volume"
    
    'find last summary row
    lastsum = ws.Cells(Rows.Count, 7).End(xlUp).Row
    
    
    Dim vols As Variant
    Dim changes As Variant
    'Erase vols
    'Erase changes
    
    
    'load data into arrays for searching
    vols = ws.Range("l2:l" & lastsum).Value
    changes = ws.Range("k2:k" & lastsum).Value
    
    
    
    'report greatest % increase and matching ticker
    ws.Range("q2").Value = Application.Max(changes)
    ws.Range("p2").Value = ws.Cells(Application.Match(ws.Range("q2").Value, changes, 0), 9).Value
    ws.Range("q2:q3").NumberFormat = "#,##0.00%"

    'report greatest % decrease and matching ticker
    ws.Range("q3").Value = Application.Min(changes)
    ws.Range("p3").Value = ws.Cells(Application.Match(ws.Range("q3").Value, changes, 0), 9).Value
    
    'report greatest total volume
    ws.Range("q4").Value = Application.Max(vols)
    ws.Range("p4").Value = ws.Cells(Application.Match(ws.Range("q4").Value, vols, 0), 9).Value
    ws.Range("q4").NumberFormat = "#,##0"

    
    ws.Range("i1:q1").EntireColumn.AutoFit
Next ws

End Sub
