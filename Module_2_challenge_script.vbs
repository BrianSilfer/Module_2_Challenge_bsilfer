Attribute VB_Name = "Module1"
Sub Module_2_Challenge_script()


'loop through each ws
For Each ws In Worksheets
         'Insert new columns
        
         'Insert new columns
    ws.Range("H1").EntireColumn.Insert
    ws.Range("I1").EntireColumn.Insert
    ws.Range("J1").EntireColumn.Insert
    ws.Range("K1").EntireColumn.Insert

'Set up for summary table

    ws.Range("N1").Value = "Ticker"
    ws.Range("O1").Value = "Value"
    ws.Range("M2").Value = "Greatest % Increase"
    ws.Range("O2").Style = "Percent"
    ws.Range("M3").Value = "Greatest % Decrease"
    ws.Range("O3").Style = "Percent"
    ws.Range("M4").Value = "Greatest Total Volume"
    
        'Add headers for each additional column
    ws.Cells(1, 8).Value = "Ticker"
    ws.Cells(1, 9).Value = "Yearly Change"
    ws.Cells(1, 10).Value = "Percent Change"
    ws.Cells(1, 11).Value = "Total Stock Volume"
        
        'define last row variable for looping through spreadsheet
    Dim row_count As Integer
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    row_count = 0
    last_row_sum = ws.Cells(Rows.Count, 8).End(xlUp).Row
    
        'create variable for ticker symbol
    Dim tick_symb As String
            
        ' create variable to store opening price
    Dim open_price As Double
        
        'create variable to store closing price
    Dim close_price As Double
        
        'Set variable for yearly change
    Dim yearly_change As Double
    
        
        'Set Variable for percent change
    Dim percent_change As Double
    

        
        'create variable to store total stock volume
    Dim total_stock_volume As Double
    total_stock_volume = 0
        'Create variable for summary table row
    Dim Sum_table_row As Double
    Sum_table_row = 2
    open_price = 0
    close_price = 0
    yearly_change = 0
    percent_change = 0
        'Loop through ticker names and output info
    For i = 2 To last_row
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Set ticker symbol to tickername
        tick_symb = ws.Cells(i, 1).Value
                
                'Set open price
        open_price = ws.Cells(i - row_count, 3).Value
                
                'Set closing price
        close_price = ws.Cells(i, 6).Value
                
                'Set yearly change value
        yearly_change = close_price - open_price
                
                'Set percent change value
                
        percent_change = (close_price - open_price) / open_price
                
                'Set Stock Volume
        total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
        
                'Print ticker symbol to summary table column H
                
        ws.Range("H" & Sum_table_row).Value = tick_symb
                
                'Print Yearly change values into column I
                
        ws.Range("I" & Sum_table_row).Value = yearly_change
                
                'Print Percent change values into column J
                
        ws.Range("J" & Sum_table_row).Value = percent_change
               
          'Print total stock volume calues in Column K
                
        'Print total stock volume
        ws.Range("K" & Sum_table_row).Value = total_stock_volume
        
        'Autofit Columns
        ws.Columns("H:M").AutoFit
        
        'Reset Variables
        Sum_table_row = Sum_table_row + 1
        open_price = 0
        close_price = 0
        total_stock_volume = 0
        yearly_change = 0
        percent_change = 0
        row_count = 0
       
        Else
            total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
            row_count = row_count + 1
        End If
       
    Next i


    'Color and percent Formatting Loop
    For i = 2 To 3001
        If ws.Cells(i, 9).Value < 0 Then
            ws.Cells(i, 9).Interior.ColorIndex = 3
        Else
            ws.Cells(i, 9).Interior.ColorIndex = 4
        End If
        ws.Cells(i, 10).Style = "Percent"
    
    Next i


    'Max Percent Loop
    Dim sum_row_count As Integer
    sum_row_count = 2
    For i = 2 To 3001
        If ws.Cells(i, 10).Value = WorksheetFunction.Max(ws.Range("J2:J3001").Value) Then
        
            max_percent = WorksheetFunction.Max(ws.Cells(i, 10).Value)
            ws.Range("O2").Value = max_percent
            ws.Range("N2").Value = ws.Cells(i, 8).Value
        Else
            sum_row_count = sum_row_count + 1
            
            
        End If

    Next i
'   Min Percent Loop
    For i = 2 To 3001
        If ws.Cells(i, 10).Value = WorksheetFunction.Min(ws.Range("J2:J3001").Value) Then
     
            min_percent = WorksheetFunction.Min(ws.Cells(i, 10).Value)
            ws.Range("O3").Value = min_percent
            ws.Range("N3").Value = ws.Cells(i, 8).Value
        Else
            sum_row_count = sum_row_count + 1
            
            
        End If

    Next i
    
'   Max total Volume loop

    For i = 2 To 3001
        If ws.Cells(i, 11).Value = WorksheetFunction.Max(ws.Range("K2:K3001").Value) Then
        
            max_volume = WorksheetFunction.Max(ws.Cells(i, 11).Value)
            ws.Range("O4").Value = max_volume
            ws.Range("N4").Value = Cells(i, 8).Value
        Else
            sum_row_count = sum_row_count + 1
            
            
        End If

    Next i

Next ws


                  
                     


End Sub
