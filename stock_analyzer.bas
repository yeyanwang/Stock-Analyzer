Attribute VB_Name = "Stock_Analyzer"
Sub Stock_Analyzer()
    Dim sheet_count, sheet As Integer
    sheet_count = ActiveWorkbook.Worksheets.count
    ' Loop thru each worksheet in the workbook
    For sheet = 1 To sheet_count
        ' Create appropriate column and row headers
        Worksheets(sheet).Activate
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        
        Dim r, r2, start, last As Long
        Dim change, percent_change, sum As Double
        change = 0
        percent_change = 0
        sum = 0
        start = 2
        r2 = 2
        last = Cells(Rows.count, 1).End(xlUp).Row
        
        ' Loop thru the <ticker> column to get unduplicated ticker symbol
        For r = 2 To last
            ' Check ticker name
            If Cells(r, 1).Value <> Cells(r + 1, 1).Value Then
                ' Add the unduplicated ticker symbol to the Ticker column if not already exists
                Cells(r2, 9).Value = Cells(r, 1).Value
                ' Calcalte the yearly change, percent change, and total stock volume for each stock
                change = Cells(r, 6).Value - Cells(start, 3).Value
                percent_change = FormatPercent(change / Cells(start, 3).Value)
                sum = sum + Cells(r, 7).Value
                ' Assign values to appropriate cells
                Cells(r2, 10).Value = change
                Cells(r2, 11).Value = percent_change
                ' Apply conditional formatting to cells in yearly change and percent change columns
                If Cells(r2, 10).Value > 0 Then
                    Range(Cells(r2, 10), Cells(r2, 11)).Interior.Color = vbGreen
                ElseIf Cells(r2, 10).Value < 0 Then
                    Range(Cells(r2, 10), Cells(r2, 11)).Interior.Color = vbRed
                End If
                Cells(r2, 12).Value = sum
                ' Set the starting volume back to 0 for the next ticker
                sum = 0
                ' Start row for open price change to next row for next ticker
                start = r + 1
                ' Row to assign values change to the next row for next ticker
                r2 = r2 + 1
            Else
            ' Keep add to the sum when the ticker name is the same
                sum = sum + Cells(r, 7).Value
            End If
        Next r
        
        Dim min, max, vol As Double
        min = Cells(2, 11).Value
        max = Cells(2, 11).Value
        vol = 0
        ' Loop thru the percent change column to find the greatest increase and the greatest decrease
        ' Loop thru the volume column to find the greatest volume
        For r = 2 To r2 - 1
            If Cells(r, 11).Value < min Then
                min = Cells(r, 11).Value
            End If
            If Cells(r, 11).Value > max Then
                max = Cells(r, 11).Value
            End If
            If Cells(r, 12).Value > vol Then
                vol = Cells(r, 12).Value
            End If
        Next r
        ' Assign values to appropriate cells
        Cells(2, 17).Value = FormatPercent(max)
        Cells(3, 17).Value = FormatPercent(min)
        Cells(4, 17).Value = vol
        
        ' Loop thru the Value column with greatest increase, greatest decrease and greatest volume
        ' Loop thru the Percent Change column and Total Stock Volumn column
        ' Compare the values to find the ticker name for each
        For r2 = 2 To r2 - 1
            If Cells(2, 17).Value = Cells(r2, 11).Value Then
                Cells(2, 16).Value = Cells(r2, 9).Value
            ElseIf Cells(3, 17).Value = Cells(r2, 11).Value Then
                Cells(3, 16).Value = Cells(r2, 9).Value
            ElseIf Cells(4, 17).Value = Cells(r2, 12).Value Then
                Cells(4, 16).Value = Cells(r2, 9).Value
            End If
        Next r2
        ' Apply auto-fit to cells in entire worksheet
        Worksheets(sheet).Cells.EntireColumn.AutoFit
        Worksheets(sheet).Cells.EntireRow.AutoFit
    Next sheet
End Sub

