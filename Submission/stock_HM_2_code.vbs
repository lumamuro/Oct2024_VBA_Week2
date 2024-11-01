Sub ticker()

    'Define variables
    Dim ws As Worksheet
    Dim ticker As String
    Dim next_ticker As String
    Dim volume As LongLong
    Dim volume_total As LongLong
    Dim i As Long
    Dim sumary_row As Long
    Dim lastRow As Long
    
    'new variables
    Dim open_price As Double
    Dim closing_price As Double
    Dim change As Double
    Dim percent_change As Double
    
    
    For Each ws In ThisWorkbook.Worksheets
        ' Create title headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"

       'reset per ticker
        volume_total = 0
        open_price = ws.Cells(2, 3).Value
        sumary_row = 2
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
        'first Loop
    For i = 2 To lastRow
        ' extract values from workbook
        ticker = ws.Cells(i, 1).Value
        volume = ws.Cells(i, 7).Value
        next_ticker = ws.Cells(i + 1, 1).Value
    
        ' if statement
    If (ticker <> next_ticker) Then
        ' add total volume
       volume_total = volume_total + volume
            
        ' Calculate
        closing_price = ws.Cells(i, 6).Value
        change = closing_price - open_price
        percent_change = change / open_price
    
        ' create title header
        ws.Cells(sumary_row, 9).Value = ticker
        ws.Cells(sumary_row, 10).Value = change
        ws.Cells(sumary_row, 11).Value = FormatPercent(percent_change)
        ws.Cells(sumary_row, 12).Value = volume_total
                
       ' Conditional Formatting
       If (change > 0) Then
        ws.Cells(sumary_row, 10).Interior.ColorIndex = 4
       ElseIf (change < 0) Then
        ws.Cells(sumary_row, 10).Interior.ColorIndex = 3
       Else
      ' Do Nothing (default White)
       End If
    
      ' reset total
          volume_total = 0
         sumary_row = sumary_row + 1
         open_price = ws.Cells(i + 1, 3).Value ' the open price of the next ticker
      Else
      ' add total volume
         volume_total = volume_total + volume
      End If
        Next i
        
        ' Second Loop for Second sumary row
        Dim max_price As Double
        Dim min_price As Double
        Dim max_volume As LongLong
        Dim max_price_ticker As String
        Dim min_price_ticker As String
        Dim max_volume_ticker As String
        Dim j As Integer
        
        ' calculate to first row of the first sumary row for comparison
        max_price = ws.Cells(2, 11).Value
        min_price = ws.Cells(2, 11).Value
        max_volume = ws.Cells(2, 12).Value
        max_price_ticker = ws.Cells(2, 9).Value
        min_price_ticker = ws.Cells(2, 9).Value
        max_volume_ticker = ws.Cells(2, 9).Value
        
        For j = 2 To sumary_row
            ' Compare current row to first row
        If (ws.Cells(j, 11).Value > max_price) Then
                ' We have a new Max Percent Change!
                max_price = ws.Cells(j, 11).Value
                max_price_ticker = ws.Cells(j, 9).Value
        End If
            
        If (Cells(j, 11).Value < min_price) Then
       ' We have a new Min Percent Change!
          min_price = ws.Cells(j, 11).Value
          min_price_ticker = ws.Cells(j, 9).Value
        End If
            
       If (Cells(j, 12).Value > max_volume) Then
       ' We have a new Max Volume!
         max_volume = ws.Cells(j, 12).Value
         max_volume_ticker = ws.Cells(j, 9).Value
            End If
        Next j
        
        ' create title header
        ws.Range("O2").Value = max_price_ticker
        ws.Range("O3").Value = min_price_ticker
        ws.Range("O4").Value = max_volume_ticker
        ws.Range("P2").Value = FormatPercent(max_price)
        ws.Range("P3").Value = FormatPercent(min_price)
        ws.Range("P4").Value = max_volume
    Next ws
End Sub

