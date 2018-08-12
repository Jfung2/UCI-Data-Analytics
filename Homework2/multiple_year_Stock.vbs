Sub Multiple_Year_Stock()

Dim Ticker As String
Dim Volume As Double
Volume = 0
Dim Summary As Integer
Summary = 2

For i = 2 To 705719
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
Ticker = Cells(i, 1).Value
Volume = Volume + Cells(i, 7).Value
Range("I" & Summary).Value = Ticker
Range("J" & Summary).Value = Volume
Summary = Summary + 1
Volume = 0
Else
Volume = Volume + Cells(i, 7).Value
End If

Next i

End sub