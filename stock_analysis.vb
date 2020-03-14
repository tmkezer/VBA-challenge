Sub Button1_Click()
    Dim ticker As String
    Dim close_px As Double
    Dim open_px As Double
    
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    ws = ActiveWorkbook.Worksheets.Count
    
    For x = 1 To ws
        
        
        lastrow = ActiveWorkbook.Worksheets(x).Cells(Rows.Count, 1).End(xlUp).Row
        
        total_vol = 0
        
        ActiveWorkbook.Worksheets(x).Cells(1, 10) = "Ticker"
        ActiveWorkbook.Worksheets(x).Cells(1, 11) = "Open Price"
        ActiveWorkbook.Worksheets(x).Cells(1, 12) = "Close Price"
        ActiveWorkbook.Worksheets(x).Cells(1, 13) = "Yearly Change"
        ActiveWorkbook.Worksheets(x).Cells(1, 14) = "Percent Change"
        ActiveWorkbook.Worksheets(x).Cells(1, 15) = "Total Volume"
        ActiveWorkbook.Worksheets(x).Cells(1, 18) = "Ticker"
        ActiveWorkbook.Worksheets(x).Cells(1, 19) = "Value"
        
        
        For i = 2 To lastrow
            If ActiveWorkbook.Worksheets(x).Cells(i, 1) <> ActiveWorkbook.Worksheets(x).Cells(i + 1, 1) Then
                ' GRAB TICKER
                ticker = ActiveWorkbook.Worksheets(x).Cells(i, 1)
                ActiveWorkbook.Worksheets(x).Range("J" & Summary_Table_Row).Value = ticker
                
                ' GRAB 20161230 Close Px
                close_px = ActiveWorkbook.Worksheets(x).Cells(i, 6)
                ActiveWorkbook.Worksheets(x).Range("L" & Summary_Table_Row).Value = close_px
                
                'Change over year
                ActiveWorkbook.Worksheets(x).Range("M" & Summary_Table_Row).Value = close_px - open_px
                
                'Calc % Change, skip if close px is zero
                If (close_px <> 0) And (open_px <> 0) Then
                
                    ActiveWorkbook.Worksheets(x).Range("N" & Summary_Table_Row).Value = (close_px - open_px) / open_px
                
                End If
                
                'Reset Open and Close Prices
                close_px = 0
                open_px = 0
                
                
                ' Calc total vol
                total_vol = total_vol + ActiveWorkbook.Worksheets(x).Cells(i, 7)
                ActiveWorkbook.Worksheets(x).Range("O" & Summary_Table_Row).Value = total_vol
                
                Summary_Table_Row = Summary_Table_Row + 1
                
                total_vol = 0
            ElseIf ActiveWorkbook.Worksheets(x).Cells(i, 1) <> ActiveWorkbook.Worksheets(x).Cells(i - 1, 1) Then
                total_vol = total_vol + ActiveWorkbook.Worksheets(x).Cells(i, 7)
                
                open_px = ActiveWorkbook.Worksheets(x).Cells(i, 3)
                ActiveWorkbook.Worksheets(x).Range("K" & Summary_Table_Row).Value = open_px
            Else
                total_vol = total_vol + ActiveWorkbook.Worksheets(x).Cells(i, 7)
            
            End If
    
        Next i
        
        last_sum_row = ActiveWorkbook.Worksheets(x).Cells(Rows.Count, 10).End(xlUp).Row
        
        
        For i = 2 To last_sum_row
            
            ActiveWorkbook.Worksheets(x).Cells(i, 14).Style = "Percent"
            If ActiveWorkbook.Worksheets(x).Cells(i, 13) > 0 Then
                ActiveWorkbook.Worksheets(x).Cells(i, 13).Interior.ColorIndex = 4
            Else
                ActiveWorkbook.Worksheets(x).Cells(i, 13).Interior.ColorIndex = 3
            End If
            
            ActiveWorkbook.Worksheets(x).Cells(2, 19).Value = WorksheetFunction.Max(Worksheets(x).Range("N" & 2 & ":" & "N" & i))
            ActiveWorkbook.Worksheets(x).Cells(3, 19).Value = WorksheetFunction.Min(Worksheets(x).Range("N" & 2 & ":" & "N" & i))
            ActiveWorkbook.Worksheets(x).Cells(4, 19).Value = WorksheetFunction.Max(Worksheets(x).Range("O" & 2 & ":" & "O" & i))
        Next i
        
        'For i = 2 To 3
            'Cells(i, 17).Style = "Percent"
        'Next i
        
        For i = 2 To last_sum_row
            If ActiveWorkbook.Worksheets(x).Cells(i, 14) = ActiveWorkbook.Worksheets(x).Cells(2, 19) Then
                ActiveWorkbook.Worksheets(x).Cells(2, 18) = ActiveWorkbook.Worksheets(x).Cells(i, 10)
            ElseIf ActiveWorkbook.Worksheets(x).Cells(i, 14) = ActiveWorkbook.Worksheets(x).Cells(3, 19) Then
                ActiveWorkbook.Worksheets(x).Cells(3, 18) = ActiveWorkbook.Worksheets(x).Cells(i, 10)
            ElseIf ActiveWorkbook.Worksheets(x).Cells(i, 15) = ActiveWorkbook.Worksheets(x).Cells(4, 19) Then
                ActiveWorkbook.Worksheets(x).Cells(4, 18) = ActiveWorkbook.Worksheets(x).Cells(i, 10)
            End If
        Next i
        
        For i = 2 To 4
            If i = 2 Then
                ActiveWorkbook.Worksheets(x).Cells(i, 17) = "Greatest % Increase"
                ActiveWorkbook.Worksheets(x).Cells(i, 19).Style = "Percent"
            ElseIf i = 3 Then
                ActiveWorkbook.Worksheets(x).Cells(i, 17) = "Greatest % Decrease"
                ActiveWorkbook.Worksheets(x).Cells(i, 19).Style = "Percent"
            Else
                ActiveWorkbook.Worksheets(x).Cells(i, 17) = "Greatest Total Volume"
            End If
        Next i
                
        Summary_Table_Row = 2
    Next x
End Sub