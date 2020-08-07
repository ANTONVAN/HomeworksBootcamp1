Attribute VB_Name = "Módulo1"
Sub Ticker()

Range("J:J", "K:K").Select
Selection.EntireColumn.Hidden = False

Range("I1") = "Ticker"
Range("L1") = "Yearly Change"
Range("M1") = "Percentage Change"


Dim OpeningValue As Double
Dim ClosingValue As Double

Dim Difference As Double
Dim PercentageChange As Double
Dim StockVolume As Variant
Dim TotalStockVolume As Variant
Dim i As Long
Dim j As Long


j = 2
For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row + 1


'Operations are made when there is a change in the ticker symbol
If Cells(i, 1) <> Cells(i - 1, 1) Then

'List Ticker Symbols
Cells(j, 9) = Cells(i, 1)

'Store the opening and closing values
Cells(j, 10) = Cells(i, 3)
Cells(j, 11) = Cells(i - 1, 6)
OpeningValue = Cells(j, 10)
ClosingValue = Cells(j + 1, 11)

'Calculate the Yearly Change
Difference = ClosingValue - OpeningValue
Cells(j, 12) = Difference
    
Cells(j, 12).Font.ColorIndex = 1
If Difference > 0 Then
Cells(j, 12).Interior.ColorIndex = 4
Else
Cells(j, 12).Interior.ColorIndex = 3
End If

'Calculate Percentage Change
If OpeningValue <> 0 Then
PercentageChange = Round((Difference / OpeningValue) * 100, 2)
Cells(j, 13) = Str(PercentageChange) + "%"
Else
PercentageChange = 0
Cells(j, 13) = PercentageChange
End If

Cells(j - 1, 14) = TotalStockVolume
j = j + 1
TotalStockVolume = 0
StockVolume = 0
End If

StockVolume = Cells(i, 7)
TotalStockVolume = TotalStockVolume + StockVolume

Next i
Range("N1") = "Total Stock Volume"

Range("J:J", "K:K").Select
Selection.EntireColumn.Hidden = True

End Sub

