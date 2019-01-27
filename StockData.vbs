Sub WallStreet():

Dim Stock_Name As String
Dim Stock_Total As String
Stock_Total = 0
Dim Stock_Open As Double
Dim Stock_Close As Double
Dim Summary_Table_Row As Integer
Dim ws As Worksheet
Summary_Table_Row = 2

For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Stock_Name = Cells(i, 1).Value
        Stock_Total = Stock_Total + Cells(i, 7).Value
        Range("I" & Summary_Table_Row).Value = Stock_Name
        Range("J" & Summary_Table_Row).Value = Stock_Total
        Summary_Table_Row = Summary_Table_Row + 1
        Brand_Total = 0
  
    Else
        Brand_Total = Brand_Total + Cells(i, 7).Value

    End If
Next i

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Total Stock Volume"

End Sub

