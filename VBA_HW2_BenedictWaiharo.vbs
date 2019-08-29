

Sub VBA_of_Wall_Street()

Dim sht As Worksheet
For Each sht In ThisWorkbook.Worksheets
'Find the last Row with data in a Column
'In this example we are finding the last row of column A
Dim lastRow As Long
With ActiveSheet
lastRow = .Cells(.Rows.count, "I").End(xlUp).Row
End With
'MsgBox lastRow



'Total number of rows
Dim finalRow As Long
finalRow = Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

'Array to store the lowerbound and upperbound range value
Dim upperB(), lowerB() As Long, size As Long
ReDim upperB(finalRow), lowerB(finalRow)

'range of ticker values
Dim tickerRange As Range



Dim countTicker, lowerBound, upperBound, midBound As Long

countTicker = 0  'counts the number of occurence of a ticker
lowerBound = 0   'lowerbound value
upperBound = 0   'upperbound value

Set tickerRange = Range("$A2:$A" & finalRow)



'Ticker calculator and compilation
Dim count As Long
count = 1

For i = 2 To finalRow
If Cells(CLng(i), 1).Value = Cells(CLng(i + 1), 1).Value Then
Cells(CLng(count + 1), 9).Value = Cells(CLng(i + 1), 1).Value

Else
count = CLng(count) + 1

End If

Next i
Range("i1").Value = "Ticker"
'End of ticker


'Volume, Percent Change, Yearly change and volume calculations
For j = 2 To lastRow
'ticker range, lower , mid and upperbound captured
countTicker = WorksheetFunction.CountIf(tickerRange, Cells(j, 9).Value)

lowerBound = CLng(lowerBound) + CLng(countTicker)
upperBound = CLng(lowerBound) + CLng(countTicker)
midBound = (lowerBound + 2) - countTicker

'Cells(CLng(j), 14).Value = lowerBound + 1
'Cells(CLng(j), 12).Value = upperBound
'Cells(CLng(j), 13).Value = midBound
'values copied to the array
lowerB(j) = midBound
upperB(j) = lowerBound + 1
'calculatoin of the volume
Cells(CLng(j), 12).Value = Application.Sum(Range(Cells(lowerB(j), 7), Cells(upperB(j), 7)))
'calculation of yearly change
Cells(CLng(j), 10).Value = Cells(upperB(j), 6).Value - Cells(lowerB(j), 3).Value
'calculation of percent change
Cells(CLng(j), 11).Value = (Cells(CLng(j), 10).Value / Cells(lowerB(j), 3).Value)
'conversion to percentagevalue
Cells(CLng(j), 11).NumberFormat = "0.00%"
Next j

'conditional formatting
Set yearlyRange = Range("j2:j" & lastRow)
For Each cell In yearlyRange
If cell.Value2 < 0 Then
Range(cell.Address).Offset(0, 0).Interior.ColorIndex = 3
Else
Range(cell.Address).Offset(0, 0).Interior.ColorIndex = 4
End If
Range("j1").Interior.ColorIndex = 0
Next cell

'"Greatest % increase", "Greatest % Decrease" and "Greatest total volume"

Range("q2") = WorksheetFunction.Max(Columns(11)) 'Greatest % increase
Range("q3") = WorksheetFunction.Min(Columns(11)) 'Greatest % Decrease
Range("q4") = WorksheetFunction.Max(Columns(12)) 'Greatest total volume
Range("q2:q3").NumberFormat = "0.00%"            'Conversion to Percentage value

'Corresponding Ticker value
For i = 2 To lastRow
 'adjacent ticker value (Greatest % increase)
If Cells(4, 17).Value = Cells(i, 12).Value Then
Range("p4").Value = Cells(i, 12).Offset(0, -3)
End If
'adjacent ticker value (Greatest % Decrease)
If Cells(3, 17).Value = Cells(i, 11).Value Then
Range("p3").Value = Cells(i, 11).Offset(0, -2)
End If
'adjacent ticker value (Greatest total volume)
If Cells(2, 17).Value = Cells(i, 11).Value Then
Range("p2").Value = Cells(i, 11).Offset(0, -2)
End If

Next i






Range("l1") = "Total Stock Volume"
Range("j1") = "Yearly Change"
Range("k1") = "Percent Change"
Range("o2") = "Greatest % increase"
Range("o3") = "Greatest % Decrease"
Range("o4") = "Greatest total volume"
Range("p1") = "Ticker"
Range("q1") = "Value"

'Auto fit every column in every sheet in workbook
sht.Cells.EntireColumn.AutoFit
Next sht

End Sub




