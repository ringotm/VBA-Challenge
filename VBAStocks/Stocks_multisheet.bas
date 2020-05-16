Attribute VB_Name = "Module12"
Sub stocks()

For Each ws In Worksheets

    Dim nextRow As Long
    Dim lastRow As Long
    Dim yearOpen As Double
    Dim yearClose As Double
    Dim TotalVolume As Double

    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    nextRow = 2

    yearOpen = ws.Cells(2, 3).Value
    TotalVolume = 0

    For i = 2 To lastRow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ws.Cells(nextRow, 9).Value = ws.Cells(i, 1).Value
            yearClose = ws.Cells(i, 6).Value
            ws.Cells(nextRow, 10).Value = yearOpen - yearClose
                If ws.Cells(nextRow, 10).Value < 0 Then
                    ws.Cells(nextRow, 10).Interior.ColorIndex = 3
                ElseIf ws.Cells(nextRow, 10).Value > 0 Then
                    ws.Cells(nextRow, 10).Interior.ColorIndex = 4
                End If
            ws.Cells(nextRow, 11).Value = (yearOpen - yearClose) / yearOpen
            ws.Cells(nextRow, 11).NumberFormat = "0.00%"
            yearOpen = ws.Cells(i + 1, 3).Value
            ws.Cells(nextRow, 12).Value = TotalVolume + ws.Cells(i, 7).Value
            nextRow = nextRow + 1
            TotalVolume = 0
        Else
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        End If
    Next i
    
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"

    lastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row

    ws.Range("Q2").Value = Application.WorksheetFunction.Max(ws.Columns("K"))
    ws.Range("Q3").Value = Application.WorksheetFunction.Min(ws.Columns("K"))
    ws.Range("Q4").Value = Application.WorksheetFunction.Max(ws.Columns("L"))
    ws.Range("Q2:Q3").NumberFormat = "0.00%"

    For i = 2 To lastRow
        If ws.Cells(i, 11).Value = ws.Range("Q2").Value Then
            ws.Range("P2").Value = ws.Cells(i, 9).Value
        ElseIf ws.Cells(i, 11).Value = ws.Range("Q3").Value Then
            ws.Range("P3").Value = ws.Cells(i, 9).Value
        ElseIf ws.Cells(i, 12).Value = ws.Range("Q4").Value Then
            ws.Range("P4").Value = ws.Cells(i, 9).Value
        End If
    Next i
Next ws
'MsgBox (Cells(35, 11).Value)
'MsgBox (Range("Q2").Value)
'MsgBox
End Sub

