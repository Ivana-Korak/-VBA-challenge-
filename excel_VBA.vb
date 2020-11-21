Sub stockmarket()



For Each ws In Worksheets
Dim ticker As String
Dim start  As Double
Dim closeprice As Double

Dim yearlychange As Double
Dim precentchange As Double
Dim totalstock As Double

'setting loop variables trough sheet

Dim counter As Long
counter = 1
Dim counter2 As Long

ticker = " "
totalstock = 0
precentchange = 0
start = 0
closeprice = 0



Dim lastrow As LongLong
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

ws.Cells(1, 9).Value = "Ticker"

ws.Cells(1, 10).Value = "precent change"

ws.Cells(1, 12).Value = "yearlychange"

ws.Cells(1, 11).Value = "total stock"

For R = 2 To lastrow

If ws.Cells(R + 1, 1).Value <> ws.Cells(R, 1).Value Then
closeprice = ws.Cells(R, 6).Value
totalstock = totalstock + ws.Cells(R, 7).Value
ticker = ws.Cells(R, 1).Value
counter = counter + 1
ws.Cells(counter, 9).Value = ticker
ws.Cells(counter, 10).Value = (closeprice - start)
ws.Cells(counter, 11).Value = totalstock
    If start = 0 Then
      ws.Cells(counter, 12).Value = ((closeprice - start) / 1) * 100
    Else
      ws.Cells(counter, 12).Value = ((closeprice - start) / start) * 100
    End If


ElseIf ws.Cells(R + 1, 1).Value = ws.Cells(R, 1).Value Then

totalstock = totalstock + ws.Cells(R, 7).Value

start = ws.Cells(R, 3).Value
End If
Next R
Dim lastrow1 As Long
lastrow1 = ws.Cells(Rows.Count, 9).End(xlUp).Row


'MsgBox (lastrow1)
Dim B As Long

For B = 2 To lastrow1

If ws.Cells(B, 10).Value > 0 Then
ws.Cells(B, 10).Interior.ColorIndex = 4
ElseIf ws.Cells(B, 10).Value < 0 Then
ws.Cells(B, 10).Interior.ColorIndex = 3
End If
Next B
Next




End Sub


