Attribute VB_Name = "Module1"
Sub Stock():

    ' Define dims
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_vol As Double

    For Each ws In Worksheets

        ' VBA to get last row and last column index
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        LastCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        
        ' Last row for Bonus
        LastRowBonus = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
        ' Headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Bonus Headers
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"

        ' AutoFit the columns to Headers and Bonus Headers
        ws.Range(ws.Cells(1, 9), ws.Cells(1, 12)).EntireColumn.AutoFit
        ws.Range(ws.Cells(1, 16), ws.Cells(1, 17)).EntireColumn.AutoFit
        ws.Range(ws.Cells(2, 15), ws.Cells(4, 15)).EntireColumn.AutoFit

        ' Check the ticker changes
        row_tracker = 2
    
        ' Grabs the initial open_price of the first ticker section
        open_price = ws.Cells(2, 3).Value
    
        ' Initial total_vol value
        total_vol = 0
    
        ' For loop that goes through the entire data to:
         ' - Find and list out all the unique tickers
         ' - Calculate yearly_change, percent_change, and total_vol
        For i = 2 To LastRow
        
            ' Calculates the total_vol
            total_vol = total_vol + ws.Cells(i, 7).Value
        
            ' Compares current cell with next cell and if not equal, it will do following
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
                ' Assign value of current cell to a cell
                ws.Cells(row_tracker, 9).Value = ws.Cells(i, 1).Value
            
                ' Grabs the close_price
                close_price = ws.Cells(i, 6).Value
            
                ' Calculates yearly_change and assign to a cell
                yearly_change = close_price - open_price
                ws.Cells(row_tracker, 10).Value = yearly_change
            
                ' Conditional formatting for highlighting positive and negative changes
                If ws.Cells(row_tracker, 10).Value > 0 Then
                    ws.Cells(row_tracker, 10).Interior.ColorIndex = 4
                ElseIf ws.Cells(row_tracker, 10).Value < 0 Then
                    ws.Cells(row_tracker, 10).Interior.ColorIndex = 3
                End If

                ' Calculate the percent_change and assign to a cell with formatting
                If open_price <> 0 And Not IsNull(open_price) Then
                    percent_change = yearly_change / open_price
                Else
                    percent_change = open_price
                End If
                
                ws.Cells(row_tracker, 11).Value = percent_change
                ws.Cells(row_tracker, 11).NumberFormat = "0.00%"
                
                ' Assign total_vol to a cell and zero out total_vol
                ws.Cells(row_tracker, 12).Value = total_vol
                total_vol = 0
            
                ' Increment row_tracker
                row_tracker = row_tracker + 1
            
                ' Set open_price to new open price value
                open_price = ws.Cells(i + 1, 3).Value
            
            End If
        Next i
        
        Dim percent_max As Double
        Dim percent_min As Double
        
        percent_max = 0
        percent_min = 0
        
        'For i = 2 To LastRowBonus
            'MaxValue = ws.WorksheetFunction.max(ws.Range(ws.Cells(2, 11), ws.Cells(LastRowBonus, 11)))
            'If Cells(i, 11).Value > percent_max Then
            '    percent_max = Cells(i, 11).Value
            'ElseIf Cells(i, 11).Value <= percent_min Then
            '    percent_min = Cells(i, 11).Value
            'MsgBox (percent_max)
            'MsgBox (percent_min)
            'End If
        'Next i
        
        Dim rng As Range: Set rng = Application.Range(Cells(2, 11), LastRowBonus)
        'Dim i As Integer
        
        For i = 2 To rng
            If Cells(i, 11).Value > percent_max Then
                percent_max = Cells(i, 11).Value
            MsgBox (percent_max)
            End If
        Next i
        
    Next ws
    MsgBox ("Completed")
End Sub
