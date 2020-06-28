Attribute VB_Name = "Module2"
Sub RunIt():

Dim Lastrow As Long
Dim Ticker As String
Dim TickerRow As Variant
Dim OpenT As Variant
Dim CloseT As Variant
Dim Change As Variant
Dim PercentChange As Variant
Dim TickerVol As Variant

TickerVol = 0
TickerRow = 2
OpenT = Cells(2, 3).Value
Lastrow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To Lastrow

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

Ticker = Cells(i, 1).Value

TickerVol = TickerVol + Cells(i, 7).Value

Range("K" & TickerRow).Value = Ticker

Range("N" & TickerRow).Value = TickerVol

TickerVol = 0

CloseT = Cells(i, 6).Value
                   
Change = CloseT - OpenT

    If Change = 0 Then
    PercentChange = 0
    
    Else
    PercentChange = Change / OpenT
    
    End If

OpenT = Cells(i + 1, 3).Value

Range("L" & TickerRow).Value = Change

Range("M" & TickerRow).Value = PercentChange

TickerRow = TickerRow + 1


Else


TickerVol = TickerVol + Cells(i, 7).Value


End If

Next i

End Sub
