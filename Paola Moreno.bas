Attribute VB_Name = "Module1"

Sub TickerCalculation()

For Each ws In Worksheets

    'Add headers to the Summary Table columns
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    
    'Initialize variables for Summary Table columns
    Dim ticker As String, open_price As Double, close_price As Double, yearly_change As Double, percent_change As Double, total_stock As Double: total_stock = 0

    
    'Keep track of the row number in the Summary Table where each ticker's data will be entered
   Dim summary_table_row As Integer: summary_table_row = 2
    
    'Determine the row number of the last data entry in the dataset.
    Dim last_row As Long
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Cycle through every row of data.
    For i = 2 To last_row
    
        'Verify if the current row represents the initial entry in the dataset.
        If i = 2 Then
            
            'Initialize the opening price for the first ticker.
            open_price = ws.Cells(i, 3).Value
            
            'Increase the total stock volume.
            total_stock = total_stock + ws.Cells(i, 7).Value
        
        ' Ensure that we are analyzing data for the same stock ticker; if not, perform specific actions.
        
        ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'Define the ticker.
            ticker = ws.Cells(i, 1).Value
            
            'Increase the total volume.
            total_stock = total_stock + ws.Cells(i, 7).Value
            
            'Establish the closing price for the current ticker.
            close_price = ws.Cells(i, 6).Value
            
            ' Determine the yearly change.
            yearly_change = close_price - open_price
            
            'Determine the percentage change.
            percent_change = (close_price - open_price) / open_price
            
            'Display the ticker name on the Summary Table.
            ws.Range("I" & summary_table_row).Value = ticker
            
            'Display yearly change in the Summary Table.
            ws.Range("J" & summary_table_row).Value = yearly_change
            
            'Display percent change in the Summary Table and format it as a percentage.
            ws.Range("K" & summary_table_row).Value = FormatPercent(percent_change)
            
            'Apply conditional formatting to the yearly and percent change.
            Dim colorIndex As Long
If yearly_change >= 0 Then
    colorIndex = 4 ' Green color index
Else
    colorIndex = 3 ' Red color index
End If
ws.Range("J" & summary_table_row & ":K" & summary_table_row).Interior.colorIndex = colorIndex

            
            'Proceed to the next row in the Summary Table.
            summary_table_row = summary_table_row + 1
            
            'Reset total stock volume amount
            total_stock = 0
            
           'Establish the opening price for the next ticker.
            open_price = ws.Cells(i + 1, 3).Value
            
        Else
        
            'Increase the total stock volume.
            total_stock = total_stock + ws.Cells(i, 7).Value
            
        End If
            
    Next i
    
    'Setting Bonus Summary Table Values
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    
    'Initialize variables for the bonus summary table
    Dim increase_ticker As String, decrease_ticker As String, stock_ticker As String, increase As Double
increase = 0

    Dim decrease As Double
    decrease = 0
    Dim stock As Double
    stock = 0
    
    'Go through every row of the Summary Table
    For i = 2 To summary_table_row
    
        'If the current ticker has a bigger percent change than the previous ticker saved in the "increase" variable, then update the greatest % increase value
        If ws.Cells(i, 11).Value > increase Then
            increase = ws.Cells(i, 11).Value
            increase_ticker = ws.Cells(i, 9).Value
        
        'If it's less than the previous ticker's decrease percent change, update the greatest % decrease value
        ElseIf ws.Cells(i, 11).Value < decrease Then
            decrease = ws.Cells(i, 11).Value
            decrease_ticker = ws.Cells(i, 9).Value
            
        End If
        
        'If the current ticker's total volume is greater than the previous ticker's, update the greatest total volume value
        If ws.Cells(i, 12).Value > stock Then
            stock_ticker = ws.Cells(i, 9).Value
            stock = ws.Cells(i, 12).Value
        
        End If
        
    Next i
    
    'Print bonus summary table findings
    ws.Range("O2").Value = increase_ticker
    ws.Range("P2").Value = FormatPercent(increase)
    ws.Range("O3").Value = decrease_ticker
    ws.Range("P3").Value = FormatPercent(decrease)
    ws.Range("O4").Value = stock_ticker
    ws.Range("P4").Value = stock

Next ws

End Sub


