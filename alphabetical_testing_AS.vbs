'Author: Arpita Sharma
'Date: 8/20/2023
'Purpose: The goal of this VBA is to summarize stock changes between Jan and Dec 2020 .

''''''''''''''''''''''
''''''Setting Up '''''
''''''''''''''''''''''

Sub AlphabeticalTesting()

'''''''''''''''''''''''''''''''''''''
''''''Setting Up for All Sheets '''''
'''''''''''''''''''''''''''''''''''''

' setting up the basics for each sheet
For Each ws In Worksheets

'''''''''''''''''''''''''''''''''
''''''Setting Variable Type '''''
'''''''''''''''''''''''''''''''''

        'setting variable types
        Dim ticker As String
        Dim open_price As Double
        Dim close_price As Double
        Dim yearly_change As Double
        Dim percent_change As Double
        Dim total_stock_volume As Double
        Dim max As Double
        Dim min As Double
        Dim max_stock As Double
        Dim lastRow As Double
        Dim lastSumRow As Double


'''''''''''''''''''''''''''''''''
''''''SET SUMMARY TABLE '''''''''
'''''''''''''''''''''''''''''''''

        ' Keep track of the location for each credit card brand in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        '*************************
        
        'setting column names here for Summary Table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 13).Value = "Open Price"
        ws.Cells(1, 14).Value = "Close Price"
        
        'setting cell names here for Greatest % Increase, Greatest % Decrease, and Greatest Total Volume
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'setting cell names for last two columns
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"


'*************************

'''''''''''''''''''''''''''''''''
''''''TABULATE SUMMARY TABLE '''''''''
'''''''''''''''''''''''''''''''''

        'Let's start really basic here and just try to figure out how to calculate everything for one row, and then work up from there
        
        'Calculating variables
        
        'Since each sheet I will be working on ends on a different row, I am going to create a last row variable
        lastRow = ws.Range("A" & Rows.Count).End(xlUp).Row
        
        
        
        For Row = 2 To lastRow
        
            If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value And ws.Cells(Row, 2).Value = 20201231 Then
            
            'set close price
               close_price = ws.Cells(Row, 6).Value
            
                  'Print close price to summary column
                   ws.Range("N" & Summary_Table_Row).Value = close_price
                    
        
            'Set the Ticker name
                ticker = ws.Cells(Row, 1).Value
        
                'Print Ticker to Summary Column
                 ws.Range("I" & Summary_Table_Row).Value = ticker
        
                'Add one to the summary table row
                 Summary_Table_Row = Summary_Table_Row + 1
                
            ElseIf ws.Cells(Row, 2).Value = 20200102 Then
              
              'set open price
               open_price = ws.Cells(Row, 3).Value
        
                 'Print open price to summary column
                  ws.Range("M" & Summary_Table_Row).Value = open_price
                    
                
             ElseIf ws.Cells(Row + 1, 1).Value = ws.Cells(Row, 1).Value Then
                
                'set total stock volume
                total_stock_volume = total_stock_volume + ws.Cells(Row, 7).Value
                
                    'Print total stock volumne
                    ws.Range("L" & Summary_Table_Row) = total_stock_volume
         
            End If
            
        Next Row

'************************************
        'Now working on calculating change variables and stock total volumne
        'starting with creating a last row for the summary table
        
        
        lastSumRow = ws.Range("I" & Rows.Count).End(xlUp).Row
        
        
        For Summary_Row = 2 To lastSumRow
            
            'set yearly change variable
            yearly_change = ws.Range("N" & Summary_Row) - ws.Range("M" & Summary_Row)
            
                'Print yearly_change to summary column
                 ws.Range("J" & Summary_Row) = yearly_change
                 
                'Format cells for yearly change, if value is less than 0, then color set to red, otherwise set to green
                If ws.Range("J" & Summary_Row).Value < 0 Then
                ws.Range("J" & Summary_Row).Interior.ColorIndex = 3
                
                ElseIf ws.Range("J" & Summary_Row).Value >= 0 Then
                ws.Range("J" & Summary_Row).Interior.ColorIndex = 4
                
                End If
                
            'set percent_change variable
            percent_change = (ws.Range("N" & Summary_Row) - ws.Range("M" & Summary_Row)) / ws.Range("M" & Summary_Row)
            
                'Print percent change to summary column
                ws.Range("K" & Summary_Row) = percent_change
            
                'Format percent change as percentage
                ws.Range("K" & Summary_Row).NumberFormat = "0.00%"
                
           
            
        Next Summary_Row

'************************************

        'Delete columns, no longer needed
        ws.Range("M1:M" & lastSumRow).ClearContents
        ws.Range("N1:N" & lastSumRow).ClearContents
    
        'now, let's figure out the greatest increase in percent change
        max = Application.WorksheetFunction.max(ws.Range("K2:K" & lastSumRow))
        
        'Print max change in cells
        ws.Range("Q2") = max
        
        'Format percent change as percentage
        ws.Range("Q2").NumberFormat = "0.00%"
        
        'let's also figure out the greatest decrease i percent change
        min = Application.WorksheetFunction.min(ws.Range("K2:K" & lastSumRow))
        
        'Print min change in cells
        ws.Range("Q3") = min
        
        'Format percent change as percentage
        ws.Range("Q3").NumberFormat = "0.00%"
    
        'now, let's figure out the greatest total stock volume
        max_stock = Application.WorksheetFunction.max(ws.Range("L2:L" & lastSumRow))
        
        'Print max stock total volume change in cells
        ws.Range("Q4") = max_stock
    
        'now, let's do a vlookup to figure out the ticker for each of these
        'we will start by copying the Ticker column I to column M
         ws.Range("M1:M" & lastSumRow).Value = ws.Range("I1:I" & lastSumRow).Value
     
        'Okay, nowe we will do a vlookup for Max and Min Percent Change
         Ticker_1 = Application.WorksheetFunction.VLookup(ws.Range("Q2").Value, ws.Range("K2:M" & lastSumRow), 3, False)
         Ticker_2 = Application.WorksheetFunction.VLookup(ws.Range("Q3").Value, ws.Range("K2:M" & lastSumRow), 3, False)
         Ticker_3 = Application.WorksheetFunction.VLookup(ws.Range("Q4").Value, ws.Range("L2:M" & lastSumRow), 2, False)
     
        'Now, printing tickers
         ws.Range("P2") = Ticker_1
         ws.Range("P3") = Ticker_2
         ws.Range("P4") = Ticker_3
     
        'Formatting cells so that I can see everything clearly
         ws.Range("O1").EntireColumn.AutoFit
         ws.Range("P1").EntireColumn.AutoFit
         ws.Range("Q1").EntireColumn.AutoFit
    
        'Okay, great, see everything, now clearing content of column, no longer needed
        ws.Range("M1:M" & lastSumRow).ClearContents
Next ws
        
End Sub


