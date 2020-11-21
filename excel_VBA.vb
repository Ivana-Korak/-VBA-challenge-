Sub stockmarket()
Dim wsk As Worksheet
Dim alphabet As Worksheet
Set aplhabet = Worksheets(1)

For Each wsk In Worksheets
Dim ticker As String
Dim openprice As Double
Dim closeprice  As Double

Dim earlychange As Double
Dim precentchange As Double
Dim totalstock As Double

'setting loop variables trough sheet

Dim counter As Long

Dim counter2 As Long

ticker = " "
totalstock = 0
precentchange = 0
earlychangde = 0

Dim lastrow As Long
lastrow = ws.Cells(Rows.Count, 1).End(xIUp).Row

For R = 2 To lastrow

If ws.Cells(R + 1, 1).Value <> ws.Cells(R, 1).Value Then

ws.Cells(1, 9).Value = "Ticker"

ws.Cells(1, 10).Value = "precentchange"

ws.Cells(1, 11).Value = "earlychange"




End Sub

