Attribute VB_Name = "VBA_challenge"
 Sub stock_data()
 
    Dim out_row As Long
    Dim r As Long
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim frst_price As Double
    Dim lst_price As Double
    Dim prcnt_price As Double
    Dim chnge_price As Double
    
'to make sure the code works on every worksheet

  For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    
    'defining all starter variables
    frst_price = ws.Cells(2, 3).Value
    chnge_price = 0
    lst_price = 0
    prcnt_change = 0
    out_row = 2
    'creating the titles/labels for each new data set
    ws.Cells(1, "I").Value = "Ticker"
    ws.Cells(1, "J").Value = "Quarterly Change"
    ws.Cells(1, "K").Value = "% Change"
    ws.Cells(1, "L").Value = "Total Volume"
    
    ws.Cells(2, "O").Value = "Greatest % Increase"
    ws.Cells(3, "O").Value = "Greatest % Decrease"
    ws.Cells(4, "O").Value = "Greatest Total Volume"
    ws.Cells(1, "P").Value = "Ticker"
    ws.Cells(1, "Q").Value = "Value"
    
    'making sure the titles are auto fit to display correctly
    ws.Range("I1:L1").Columns.AutoFit
    ws.Range("O2:O4").Columns.AutoFit
    'remember to change range to "Q:Q" when completed
    ws.Range("Q:Q").Columns.AutoFit
    
    'defining the lastRow, so the code will run through the entire sheet
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row - 1
    
    
    'formatting the percent change column as percents
    Range("K:K").NumberFormat = "0.00%"
    
        For r = 2 To LastRow
            If ws.Cells(r, 1).Value <> ws.Cells(r + 1, 1).Value Then
                'displaying the ticker and the total volume
                ws.Cells(out_row, "I").Value = ws.Cells(r, 1).Value
                ws.Cells(out_row, "L").Value = volume + ws.Cells(r, 7).Value
                'zeroing out the volume to reset it
                volume = 0
                
                'setting the closing price, change in price, and the percent change in price
                lst_price = ws.Cells(r, "F").Value
                chnge_price = lst_price - frst_price
                
                If frst_price <> 0 Then
                    prcnt_change = chnge_price / frst_price
                Else
                    prcnt_change = 0
               End If
               
                'displaying the % change / change in price + assigning the first price for the next ticker
                ws.Cells(out_row, "K").Value = prcnt_change
                frst_price = ws.Cells(r + 1, "C").Value
                ws.Cells(out_row, "J").Value = chnge_price
                
                out_row = out_row + 1
                
            Else
                volume = volume + ws.Cells(r, 7).Value

            End If
        Next r
        
    Next ws
    

    
End Sub
Sub max_min()

'dim variables
Dim LastRow As Long
Dim r As Long
Dim minValue As Double
Dim maxValue As Double
Dim maxValue2 As Double
Dim ws As Worksheet


    For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row - 1

    'establishing the max/min values greatest decrease/increase of % and the greatest volume
    maxValue = Application.WorksheetFunction.Max(Range(Cells(2, 11), Cells(LastRow, 11)))
    minValue = Application.WorksheetFunction.min(Range(Cells(2, 11), Cells(LastRow, 11)))
    maxValue2 = Application.WorksheetFunction.Max(Range(Cells(2, 12), Cells(LastRow, 12)))

    'displaying those values
    ws.Cells(3, "Q").Value = minValue
    ws.Cells(2, "Q").Value = maxValue
    ws.Cells(4, "Q").Value = maxValue2
    ws.Range("Q2:Q3").NumberFormat = "0.00%"

    'displaying respective ticker
        For r = 2 To LastRow
            If ws.Cells(r, "K").Value = maxValue Then
                ws.Cells(2, "P").Value = ws.Cells(r, "I").Value
            ElseIf ws.Cells(r, "K").Value = minValue Then
                ws.Cells(3, "P").Value = ws.Cells(r, "I").Value
            ElseIf ws.Cells(r, "L").Value = maxValue2 Then
                ws.Cells(4, "P").Value = Cells(r, "I").Value
            End If
        
        Next r
    
    Next ws

End Sub

Sub coloring()

Dim r As Long
Dim ws As Worksheet
Dim LastRow As Long
LastRow = Cells(Rows.Count, "A").End(xlUp).Row - 1

    For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    
        For r = 2 To LastRow
            'color conditional formatting for quarterly change
            If ws.Cells(r, "J").Value > 0 Then
                ws.Cells(r, "J").Interior.ColorIndex = 4 'green
            ElseIf ws.Cells(r, "J").Value < 0 Then
                ws.Cells(r, "J").Interior.ColorIndex = 3 'red
            End If
    
        Next r
        
    Next ws
    
End Sub
