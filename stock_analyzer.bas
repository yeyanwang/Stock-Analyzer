Attribute VB_Name = "Module1"
Sub stock_analyze()

' Assign appropriate column and row headers
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

' Loop thru the <ticker> column to get unqiue Ticker value on each row
' Add the ticker to the new Ticker column if not already exists
Dim r, r2, last As Long
Dim ticker_name As String
r2 = 1
last = Cells(rows.Count, 1).End(xlUp).row
ticker_name = ""

For r = 2 To last
    If Cells(r, 1).Value <> ticker_name Then
        r2 = r2 + 1
        ticker_name = Cells(r, 1).Value
        Cells(r2, 9).Value = ticker_name
    End If
Next r

' Get the open and close prices and find the change
' Loop thru <ticker> column and Ticker column to compare ticker name
' Use counter variable to store which row the open and close price is on
' Use sum to storage the total stock volume for each ticker
' open price = first row when looking at that ticker
' close price = last row before going to another ticker
' sum = all volume add up for that ticker name
Dim op, cl, op_row, cl_row, counter, last2 As Long
Dim sum As Double
last2 = Cells(rows.Count, 9).End(xlUp).row

For r2 = 2 To last2
    counter = 0
    sum = 0
    For r = 2 To last
        If Cells(r2, 9).Value = Cells(r, 1).Value Then
            sum = sum + Cells(r, 7).Value
            If counter = 0 Then
                op_row = r
            End If
            counter = counter + 1
        cl_row = r
        End If
    op = Cells(op_row, 3).Value
    cl = Cells(cl_row, 6).Value
    Next r
    ' yearly change = close - open for that ticker
    Cells(r2, 10).Value = cl - op
    If Cells(r2, 10).Value > 0 Then
        Cells(r2, 10).Interior.Color = vbGreen
    ElseIf Cells(r2, 10).Value < 0 Then
        Cells(r2, 10).Interior.Color = vbRed
    End If
    ' percent change = % of yearly change / open price
    Cells(r2, 11).Value = FormatPercent((cl - op) / op, 2)
    Cells(r2, 12) = sum
Next r2

' Loop thru the yearly change column to find the maximum (greatest increase) and minimum values (Greatest decrease)
' Loop thru the volume column to find the greatest volume
Dim min, max, vol As Double
min = Cells(2, 11).Value
max = Cells(2, 11).Value
vol = 0

For r2 = 2 To last2
    If Cells(r2, 11).Value < min Then
        min = Cells(r2, 11).Value
    End If
    If Cells(r2, 11).Value > max Then
        max = Cells(r2, 11).Value
    End If
    If Cells(r2, 12).Value > vol Then
        vol = Cells(r2, 12).Value
    End If
Next r2
Cells(2, 17).Value = FormatPercent(max)
Cells(3, 17).Value = FormatPercent(min)
Cells(4, 17).Value = vol

' Loop thru the Value column with greatest increase, greatest decrease and greatest volume
' Loop thru the Percent Change column and Total Stock Volumn column
' Compare the values to find the ticker name for each
For r2 = 2 To last2
    If Cells(2, 17).Value = Cells(r2, 11).Value Then
        Cells(2, 16).Value = Cells(r2, 9).Value
    ElseIf Cells(3, 17).Value = Cells(r2, 11).Value Then
        Cells(3, 16).Value = Cells(r2, 9).Value
    ElseIf Cells(4, 17).Value = Cells(r2, 12).Value Then
        Cells(4, 16).Value = Cells(r2, 9).Value
    End If
Next r2
End Sub
