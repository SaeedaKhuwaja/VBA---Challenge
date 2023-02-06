
Sub stock()
    ' track changes in column A
    ' add up totals in column G based on changes in Column A
    
    ' each time the ticker symbol changes in column A,
    ' populate the name of the ticker in column J
    ' display the total in column M
    ' reset the total and start tracking for the next ticker symbol
    
    'to run the same code in all the sheets
    For Each ws In Worksheets
    
        'set the header in the sheet
        ws.Cells(1, 10).Value = "Ticker Symbol"
        ws.Cells(1, 11).Value = "Yearly Change"
        ws.Cells(1, 12).Value = "Percentage Change"
        ws.Cells(1, 13).Value = "Total Stock Volume"
        ws.Cells(1, 15).Value = "Greatest % Change"
        ws.Cells(2, 15).Value = "Lowest % Change"
        ws.Cells(3, 15).Value = "Greatest Total Volume"
        
        'create a variable to hold the last row
        Dim lastRow As Long
        
        ' first find the last row in the sheet
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).row
        
        'create the variables
        Dim tsymbol As String
        
        Dim greatest_percentage_change As Double
        
        Dim lowest_percentage_change As Double
        
        Dim greatest_total_volume As Double
        
        Dim percentage_change As Double
        
        Dim greatest_symbol As String
        
        Dim lowest_symbol As String
        
        Dim total_symbol As String
        
        'declare variable for open ticker
        Dim open_ticker As Double
                
        'declare variable for close ticker
        Dim close_ticker As Double
            
        'declare a variable to store yearly change
        Dim yearly_change As Double
        
        ' variable to hold the totals for the tickers volume
        Dim Total_volume As Double
        
         ' variable to hold the rows in open column
        Dim skipRow As Boolean
    
            ' Initialize the total volume to 0
            Total_volume = 0
            
            ' Initialize the variable skipRow to 0
            skipRow = False
            
            'Initialize the greatest percentage change to 0.0
            greatest_percentage_change = 0#
            
            'Initialize the lowest percentage chanhe to highest value
            lowest_percentage_change = 100#
            
            'initialize the greatest total volume to 0
            greatest_total_volume = 0
        
        ' variable to hold the rows in the total colums (columns J and M)
        Dim tRows As Long
        ' first row to populate in column J and 2nd in M
        tRows = 2
        
        ' declare variable to hold the row
        Dim row As Long
        
         ' loop through the rows and check the changes in the ticker symbols
        For row = 2 To lastRow
            ' check the changes in the ticker symbols
            If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
        
                ' if the ticker symbol changes, then display the change
                'MsgBox (Cells(row, 1).Value + " -> " + Cells(row + 1, 1).Value)
                
                'take the last value of the close year
                close_ticker = ws.Cells(row, 6).Value
                
                'calculate the yearly change by subtracting the open ticker from the close ticker
                yearly_change = close_ticker - open_ticker
                
                'calculate the percentage change by dividing the yearly change by open ticker
                percentage_change = yearly_change / open_ticker
                
                'run if loop to calculate the greatest perentage change
                'if percentage change > 0 then set the value to greatest percentage change
                If percentage_change > greatest_percentage_change Then
                     greatest_percentage_change = percentage_change
                     
                     'take the following symbol of the greatest percenatge change
                     greatest_symbol = ws.Cells(row, 1).Value
                     
                End If
                
                'run if loop to calculate the lowest perentage change
                'if percentage change < 100 then set the value to greatest percentage change
                If percentage_change < lowest_percentage_change Then
                    lowest_percentage_change = percentage_change
                    
                    'take the following symbol of the lowest percenatge change
                    lowest_symbol = ws.Cells(row, 1).Value
                    
                End If
                
                'run if loop to calculate the greatest total volume
                'if total volume > 0 then set the value to greatest total volume
                If Total_volume > greatest_total_volume Then
                     greatest_total_volume = Total_volume
                     
                     total_symbol = ws.Cells(row, 1).Value
                End If
                   
                    'set the headers
                    ws.Cells(1, 17).Value = greatest_percentage_change
                    ws.Cells(1, 16).Value = greatest_symbol
                    ws.Cells(2, 17).Value = lowest_percentage_change
                    ws.Cells(2, 16).Value = lowest_symbol
                    ws.Cells(3, 17).Value = greatest_total_volume
                    ws.Cells(3, 16).Value = total_symbol
                
                ' set the ticker name
                tsymbol = ws.Cells(row, 1).Value ' grabs the value from column A BEFORE the change
                
                'display the ticker symbols in column J
                ws.Cells(tRows, 10).Value = tsymbol
                 
                'diaply the total volume in Column M
                ws.Cells(tRows, 13).Value = Total_volume
                
                'display the percentage change in Column L
                ws.Cells(tRows, 12).Value = percentage_change
                
                If percentage_change > 0 Then
                        
                    ws.Cells(tRows, 12).Interior.ColorIndex = 4
                    
                    Else
                    
                    ws.Cells(tRows, 12).Interior.ColorIndex = 3
                    
                End If
                
                skipRow = False
                
                'set the conditional formating in yearly change
                If yearly_change > 0 Then
                
                    ws.Cells(tRows, 11).Value = yearly_change
                    ws.Cells(tRows, 11).Interior.ColorIndex = 4 'set the color to GREEN if its positive
                    
                Else
                    
                    ws.Cells(tRows, 11).Value = yearly_change
                    ws.Cells(tRows, 11).Interior.ColorIndex = 3 'set the color to RED if its negative
                    
                End If
                
                ' add 1 to the ticker row to go to the next row
                tRows = tRows + 1
                
                ' add to the total volume
                Total_volume = Total_volume + ws.Cells(row, 7).Value  'grab the total value before change
            
                'reset the total to 0 before adding new symbol
                Total_volume = 0
            
            Else
                
                'if there is no change in the ticker symbol keep adding the total
                Total_volume = Total_volume + ws.Cells(row, 7).Value
                
                    'If there is no change in the ticker symbol keep skipping the rows.
                    If skipRow = False Then
                        
                        'if the ticker symbol changes take the open value
                        open_ticker = ws.Cells(row, 3).Value
                        
                        'after taking one value skip the open value if the ticker symbol is same
                        skipRow = True
                    End If
                
    
            End If
                
        Next row
        
    Next ws
    
End Sub
