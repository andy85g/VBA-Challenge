Attribute VB_Name = "Module1"
Sub stonks()
' set variables
Dim ws As Worksheet
For Each ws In Worksheets
ws.Activate
Dim stocks_volume As Double
Dim decrease As Double
Dim j As Double
Dim max As Double
Dim stocks_name As String
Dim greater As Double
Dim stocks_close As Double
Dim percent_change As Double
Dim stocks_change As Double
Dim increases As Double
Dim stocks_open As Double
Dim ticker_list As Long
Dim i As Double
' set initial values to variables used for loop calculations
stocks_volume = 2
max = 0
increase = 0
decrease = 0
' variables for value for last row
Dim LastRow As Long
Dim LastRow2 As Long
' find last row of columns automatically
' ensure formula finds the changing last rows from one worksheet to another
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
LastRow2 = ws.Cells(Rows.Count, 12).End(xlUp).Row
' start counter at row 2 to avoid headers
ticker_list = 2
For i = 2 To LastRow
' if the ticker symbol in column 1 does not match the row below it
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
' then take ticker symbol value
stocks_name = Cells(i, 1).Value
' and add the last stocks volume for this quarter to the total
stocks_volume = stocks_volume + Cells(i, 7).Value
' assign value of opening price for first ticker symbol since the loop wont catch it for first time through
If stocks_open = 0 Then
stocks_open = Cells(2, 3).Value
Else
End If
' set stocks close value to be value then found in column 6
stocks_close = Cells(i, 6).Value
' a bug in my code was setting some closing values to percentage, this was an easy fix
Cells(i, 6).Style = "Normal"
' calculate percentage change value and format the cell
percent_change = ((stocks_close / stocks_open) - 1)
'calculate stocks change value
stocks_change = Cells(i, 6).Value - stocks_open
' list ticker smbol in Range I
Range("I" & ticker_list).Value = stocks_name
' list the stocks's change value in Range J
Range("J" & ticker_list).Value = stocks_change
' list percentage change value and format the cell
Range("K" & ticker_list).Value = percent_change
Range("K" & ticker_list).Style = "Percent"
Range("K" & ticker_list).NumberFormat = "0.00%"
' list stocks volume data in Range L
Range("L" & ticker_list).Value = stocks_volume
' advance ticker list value to next row, the first row on a new symbol
ticker_list = ticker_list + 1
' take the value of the symbols opening day value for the quarter
stocks_open = Cells(i + 1, 3).Value
' set stocks volume variable at zero to adding up volume for new symbol
stocks_volume = 0
Cells(i, 11).Style = "Percent"
' if next ticker symbol in Rnage A matches the current row, add volume to total and repeat i loop to continue down Range A
Else
stocks_volume = stocks_volume + Cells(i, 7).Value
End If
' colour formatting Quarterly Change value
' no colour if zero
If Range("J" & ticker_list).Value = 0 Then
Range("J" & ticker_list).Interior.ColorIndex = Clear
' green if value is positive
ElseIf Range("J" & ticker_list).Value > 0 Then
Range("J" & ticker_list).Interior.ColorIndex = 4
' red if value is negative
ElseIf Range("J" & ticker_list).Value < 0 Then
Range("J" & ticker_list).Interior.ColorIndex = 3
End If
Next i
' j loop to find values for greatest percentage increase, decrease, stocks volume
For j = 2 To LastRow2
' finding greatest increase and fomatting cell
If Cells(j, 11).Value >= increase Then
increase = Cells(j, 11).Value
Cells(2, 15).Value = "Greatest % increase"
Cells(2, 17).Value = increase
Cells(2, 16).Value = Cells(j, 9).Value
Cells(2, 17).Style = "Percent"
Cells(2, 17).NumberFormat = "0.00%"
' finding greatest decrease and fomatting cell
ElseIf Cells(j, 11).Value <= decrease Then
decrease = Cells(j, 11).Value
Cells(3, 15).Value = "Greatest % decrease"
Cells(3, 17).Value = decrease
Cells(3, 16).Value = Cells(j, 9).Value
Cells(3, 17).Style = "Percent"
Cells(3, 17).NumberFormat = "0.00%"
' finding greatest total volume and formatting cell
ElseIf Cells(j, 12).Value >= max Then
max = Cells(j, 12).Value
Cells(4, 15).Value = "Greatest total volume"
Cells(4, 17).Value = max
Cells(4, 16).Value = Cells(j, 9).Value
Cells(4, 17).Style = "Normal"
Else
End If
Next j
Next ws
End Sub
