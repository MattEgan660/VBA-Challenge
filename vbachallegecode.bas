Sub ticker_analysis()

'Definitions
Dim ws As Worksheet
Dim Ticker As String
Dim year_open As Double
Dim year_close As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim vol As Integer
Dim row As Double


vol = 0
row = 2


'loop through each worksheet in active workbook
    For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
    
    'find last row of data
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row
    
    'add summary headings
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    'add Max/Min summary table
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    
     'set yearly change column formating
     ws.Columns("K").NumberFormat = "0.00%"
    
    'set firt year open price
     year_open = Cells(2, 3).Value
          
        'loop through data
        For i = 2 To lastrow
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                'Ticker
                Ticker = Cells(i, 1).Value
                Cells(row, 9).Value = Ticker
                
                'Year close
                year_close = Cells(i, 6).Value
                
                'Yearly change
                Yearly_Change = year_close - year_open
                Cells(row, 10).Value = Yearly_Change
                    
                    
                'percent change
                    If year_open = 0 And year_close = 0 Then
                        Percent_Change = 0
                    ElseIf year_open = 0 And year_close <> 0 Then
                        Percent_Change = 1
                    Else
                        Percent_Change = Yearly_Change / year_open
                        Cells(row, 11).Value = Percent_Change
                    End If
                
                'Total Volume
                Volume = Volume + Cells(i, 7).Value
                Cells(row, 12).Value = Volume
                
                'iterate to next row
                row = row + 1
                
                'Reset year_open
                year_open = Cells(i + 1, 3)
                
                'reset volume
                Volume = 0
            
            Else
                Volume = Volume + Cells(i, 7).Value
            End If
        
        Next i
 
newdata_lastrow = ws.Cells(Rows.Count, 9).End(xlUp).row

        For j = 2 To newdata_lastrow
            If Cells(j, 10).Value >= 0 Then
            Cells(j, 10).Interior.Color = vbGreen
            Else
            Cells(j, 10).Interior.Color = vbRed
            End If
        Next j
    
    For k = 2 To newdata_lastrow
        If Cells(k, 11).Value = Application.WorksheetFunction.Max(ws.Range("k2:k" & newdata_lastrow)) Then
            Cells(2, 15).Value = Cells(k, 9).Value
            Cells(2, 16).Value = Cells(k, 11).Value
            Cells(2, 16).NumberFormat = "0.00%"
        ElseIf Cells(k, 11).Value = Application.WorksheetFunction.Min(ws.Range("k2:k" & newdata_lastrow)) Then
            Cells(3, 15).Value = Cells(k, 9).Value
            Cells(3, 16).Value = Cells(k, 11).Value
            Cells(3, 16).NumberFormat = "0.00%"
        ElseIf Cells(k, 12).Value = Application.WorksheetFunction.Max(ws.Range("l2:l" & newdata_lastrow)) Then
            Cells(4, 15).Value = Cells(k, 9).Value
            Cells(4, 16).Value = Cells(k, 12).Value
        End If
    
    Next k

    Next ws

End Sub


