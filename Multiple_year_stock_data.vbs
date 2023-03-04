Sub StockLoop()

'Define Variables
Dim WS As Worksheet
Dim LastRow, Row, n, j, s, k As Long
Dim oPrice, cPrice As Double
Dim StockData As Range
Dim a As Variant

'Loop though each worksheet
For Each WS In ActiveWorkbook.Worksheets
WS.Select

'Create new column
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

'Create new row
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

'Define value to variable
j = 2
k = 1
LastRow = Cells(Rows.Count, "A").End(xlUp).Row
Row = Cells(Rows.Count, "I").End(xlUp).Row

'Set Stock Data Range
Set StockData = Range("A2")

'Loop through the Stock Data for ticker symbol and sum up the total stock volume
Do While Len(StockData) > 0
    Do
        n = n + StockData.Offset(, 6).Value
        Set StockData = StockData.Offset(1)
    Loop While StockData.Row = 2 Or StockData.Value = StockData.Offset(-1).Value
    Cells(j, "I") = StockData.Offset(-1).Value
    Cells(j, "L") = n
    n = 0
    j = j + 1
Loop

'Loop through the Stockdata and find the Opening Price and Closing Price
With Range("A2:B" & LastRow + 1)
        a = .Value
        s = 1

        For i = LBound(a) To UBound(a) - 1
            If a(i, 1) <> a(i + 1, 1) Then
                Set StockData = .Range("A" & s).Resize(i - s + 1)
                
                k = k + 1
                
                oPrice = StockData(1).Offset(, 2).Value
                cPrice = StockData(StockData.Rows.Count).Offset(, 5).Value
                
                'Calculate the Yearly Change and store in row J
                Range("J" & k).Value = cPrice - oPrice
                
                'Format the cell background colour base on positive and negative figure
                If (Range("J" & k).Value > 0) Then
                    Range("J" & k).Interior.ColorIndex = 4
                Else
                    Range("J" & k).Interior.ColorIndex = 3
                End If
                
                'Calculate the percentage change and store in row K
                Range("K" & k).Value = FormatPercent(((cPrice - oPrice) / oPrice), -1)
                
                s = i + 1
            End If
        Next i
End With

For i = 2 To Row
    'Find the max yearly change figure
    If (Range("J" & i) = Application.WorksheetFunction.Max(WS.Range("J2:J" & Row))) Then
        Range("P2") = Range("I" & i)
        Range("Q2") = Range("J" & i)
    End If

    'Find the min yearly change figure
    If (Range("J" & i) = Application.WorksheetFunction.Min(WS.Range("J2:J" & Row))) Then
        Range("P3") = Range("I" & i)
        Range("Q3") = Range("J" & i)
    End If

    'Find the max total stock volume
    If (Range("L" & i) = Application.WorksheetFunction.Max(WS.Range("L2:L" & Row))) Then
        Range("P4") = Range("I" & i)
        Range("Q4") = Range("L" & i)
    End If

Next i

'Set Column size
ActiveSheet.UsedRange.EntireColumn.AutoFit

Next

End Sub

