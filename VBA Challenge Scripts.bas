Attribute VB_Name = "Module3"
Sub StocksLoop():

Dim TickerSymbol As String
Dim YearlyChange As Double
Dim PercentYearlyChange As Double
Dim FirstOpen As Double
Dim LastClose As Double
Dim Volume As LongLong
Dim LastRow As Long
Dim Summary_Table_Row As Integer
Dim ws As Worksheet
Summary_Table_Row = 2

Cells(1, 9).Value = "Ticker Symbol"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(2, 14).Value = "Greatest % Increase"
Cells(3, 14).Value = "Greatest % Decrease"
Cells(4, 14).Value = "Greatest Total Volume"
Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Value"

Range("K2:K43398").NumberFormat = "0.00%"
FirstOpen = Cells(2, 3).Value
Volume = 0
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To LastRow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value And Cells(i + 1, 3).Value <> 0 Then
        TickerSymbol = Cells(i, 1).Value
        Cells(Summary_Table_Row, 9).Value = TickerSymbol
        LastClose = Cells(i, 6).Value
        PercentYearlyChange = ((LastClose - FirstOpen) / FirstOpen)
        YearlyChange = (LastClose - FirstOpen)
        Cells(Summary_Table_Row, 9).Value = TickerSymbol
        Cells(Summary_Table_Row, 10).Value = YearlyChange
        Cells(Summary_Table_Row, 11).Value = PercentYearlyChange
        Cells(Summary_Table_Row, 12).Value = Volume
        Summary_Table_Row = Summary_Table_Row + 1
        FirstOpen = Cells(i + 1, 3).Value
    Else
        Volume = Volume + Cells(i, 7).Value
    End If

   
    If Cells(i, 10).Value > 0 Then
        Cells(i, 10).Interior.ColorIndex = 4
    ElseIf Cells(i, 10).Value < 0 Then
        Cells(i, 10).Interior.ColorIndex = 3
    End If
    
    Next i

    
End Sub


