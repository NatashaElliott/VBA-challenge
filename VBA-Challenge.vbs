Sub Ticker()
    Dim TickerName As String
    Dim i As Long
    Dim TickerCounter As Integer
    Dim lastrow As Long
    Dim Openprice As Double
    Dim Closeprice As Double
    Dim Stockvolume As Double
    Dim Greatestincrease As Double
    Dim Greatestdecrease As Double
    Dim Greatestvolume As Double
    Dim Increase As String
    Dim Decrease As String
    Dim Volume As String
    
    ' Loop through all sheets
    For Each ws In Worksheets
    
    TickerName = ""
    TickerCounter = 2
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastrow
        If (ws.Cells(i, 1).Value <> TickerName) Then
            Stockvolume = ws.Cells(i, 7).Value
            ws.Cells(TickerCounter, 12).Value = Stockvolume
            Openprice = ws.Cells(i, 3).Value
            TickerName = ws.Cells(i, 1).Value
            ws.Cells(TickerCounter, 9).Value = TickerName
        ElseIf (ws.Cells(i, 1).Value = TickerName) Then
            Stockvolume = Stockvolume + ws.Cells(i, 7).Value
        End If

        If (ws.Cells(i + 1, 1).Value <> TickerName) Then
            Closeprice = ws.Cells(i, 6).Value
            ws.Cells(TickerCounter, 10).Value = Closeprice - Openprice
            ws.Cells(TickerCounter, 10).NumberFormat = "0.00"
            If ws.Cells(TickerCounter, 10).Value > 0 Then
                ws.Cells(TickerCounter, 10).Interior.ColorIndex = 4
            End If
            If ws.Cells(TickerCounter, 10).Value < 0 Then
                ws.Cells(TickerCounter, 10).Interior.ColorIndex = 3
            End If
            ws.Cells(TickerCounter, 11).Value = (Closeprice - Openprice) / Openprice
            ws.Cells(TickerCounter, 11).NumberFormat = "0.00%"
            ws.Cells(TickerCounter, 12).Value = Stockvolume
            TickerCounter = TickerCounter + 1
        End If
      
    Next i
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    ' Bonus Increase
    Greatestincrease = 0
    For i = 2 To TickerCounter
    If ws.Cells(i, 11).Value > Greatestincrease Then
    Greatestincrease = ws.Cells(i, 11).Value
    Increase = ws.Cells(i, 9).Value
    End If
    
    Next i
    
    ws.Cells(2, 17).Value = Greatestincrease
    ws.Cells(2, 16).Value = Increase
    ws.Cells(2, 17).NumberFormat = "0.00%"
    
    ' Bonus Decrease
    Greatestdecrease = 0
    For i = 2 To TickerCounter
    If ws.Cells(i, 11).Value < Greatestdecrease Then
    Greatestdecrease = ws.Cells(i, 11).Value
    Decrease = ws.Cells(i, 9).Value
    End If
    
    Next i
    
    ws.Cells(3, 17).Value = Greatestdecrease
    ws.Cells(3, 16).Value = Decrease
    ws.Cells(3, 17).NumberFormat = "0.00%"
    
    ' Bonus Volume
    Greatestvolume = 0
    For i = 2 To TickerCounter
    If ws.Cells(i, 12).Value > Greatestvolume Then
    Greatestvolume = ws.Cells(i, 12).Value
    Volume = ws.Cells(i, 9).Value
    End If
    
    Next i
    
    ws.Cells(4, 17).Value = Greatestvolume
    ws.Cells(4, 16).Value = Volume
    
    Next ws
    
End Sub
