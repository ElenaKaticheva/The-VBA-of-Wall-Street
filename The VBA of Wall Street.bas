Attribute VB_Name = "Module1"
Sub ws()
    'Loop for all sheets
Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
Dim open_price_beginning As Double
Dim close_price_end As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim total_volume As Double
Dim row As Double
Dim column As Double
Dim ticker_name As String
Dim x As Long
    total_volume = 0
    column = 1
    row = 2
    
Cells(1, 9).Value = "ticker"
Cells(1, 10).Value = "total stock volume"
Cells(1, 11).Value = "yearly change"
Cells(1, 12).Value = "percent change"
Cells(1, 16).Value = "ticker"
Cells(1, 17).Value = "value"
Cells(2, 15).Value = "greatest % increase"
Cells(3, 15).Value = "greatest % decrease"
Cells(4, 15).Value = "greatest volume total"

open_price = Cells(2, column + 2).Value
'determine the last row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).row

For x = 2 To LastRow
    If Cells(x + 1, column).Value <> Cells(x, column).Value Then
        ticker_name = Cells(x, column).Value
            Cells(row, column + 8).Value = ticker_name
        close_price_end = Cells(x, column + 5).Value
        yearly_change = close_price_end - open_price_beginning
            Cells(row, column + 9).Value = yearly_change
        If (open_price_beginning = 0 And close_price_end = 0) Then
            percent_change = 0
        ElseIf (open_price_beginning = 0 And close_price_end > 0) Or (open_price_beginning = 0 And close_price_end < 0) Then
            percent_change = 1
        Else
                percent_change = yearly_change / open_price_beginning
                Cells(row, column + 10).Value = percent_change
                Cells(row, column + 10).NumberFormat = "0.00%"
        End If
        total_volume = total_volume + Cells(x, column + 6).Value
            Cells(row, column + 11).Value = total_volume
        open_price_beginning = Cells(x + 1, column + 2)
                total_volume = 0
                row = row + 1
        Else
        total_volume = total_volume + Cells(x, column + 6).Value
        End If
    Next x
        
        'determine the last row for yearly change
        LastRowYearlyChange = ws.Cells(Rows.Count, column + 8).End(xlUp).row
        
        'color change
            For y = 2 To LastRowYearlyChange
                If (Cells(y, column + 9).Value > 0 Or Cells(y, column + 9).Value = 0) Then
                    Cells(y, column + 9).Interior.ColorIndex = 4
                ElseIf Cells(y, column + 9).Value < 0 Then
                    Cells(y, column + 9).Interior.ColorIndex = 3
                End If
            Next y

        'determine the last row
            For i = 2 To LastRowYearlyChange
                If Cells(i, column + 10).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & LastRowYearlyChange)) Then
                    Cells(2, column + 15).Value = Cells(i, column + 8).Value
                    Cells(2, column + 16).Value = Cells(i, column + 10).Value
                    Cells(2, column + 16).NumberFormat = "0.00%"
                ElseIf Cells(i, column + 10).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & LastRowYearlyChange)) Then
                Cells(3, column + 15).Value = Cells(i, column + 8).Value
                Cells(3, column + 16).Value = Cells(i, column + 10).Value
                Cells(3, column + 16).NumberFormat = "0.00%"
                ElseIf Cells(i, column + 11).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & LastRowYearlyChange)) Then
                Cells(4, column + 15).Value = Cells(i, column + 8).Value
                Cells(4, column + 16).Value = Cells(i, column + 11).Value
                End If
        Next i
    Next ws

End Sub

    

