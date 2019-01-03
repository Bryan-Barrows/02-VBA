Sub alphabetical_testing()

    'define variables
    Dim ticker As String
    Dim stockVolume As Double
    Dim summary_table_row As Double
    Dim year_open As Double
    Dim year_close As Double
    Dim year_open_date As Integer
    Dim max_increase As Variant
    Dim max_decrease As Variant
    Dim max_volume As Variant
   
          
           
    'set values for variables
    stockVolume = 0
    summary_table_row = 2
    
    
    'column / table headings
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Yearly Percentage"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    
       
    'loop thru all tickers
    For i = 2 To 70926
    
        'check if first line of stock
        If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
        
            'if so, set as opening value
            year_open = Cells(i, 3).Value
        
        End If
        
        'check if last row of stock
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            'if so, set as closing value
            year_close = Cells(i, 6).Value
            
            'calculate yearly change & percentage
            yearly_change = year_close - year_open
            year_percent = yearly_change / year_open
            
            ticker = Cells(i, 1).Value
            
            stockVolume = stockVolume + Cells(i, 7).Value
            
            Range("j" & summary_table_row).Value = yearly_change
            
            Range("i" & summary_table_row).Value = ticker
            
            Range("k" & summary_table_row).Value = year_percent
            
            Range("l" & summary_table_row).Value = stockVolume
            
            summary_table_row = summary_table_row + 1
            
            stockVolume = 0
             
            
        Else
        
        'if not start of new ticker, add volume to total
        stockVolume = stockVolume + Cells(i, 7).Value
        
        End If
    
    Next i
     
    
    'find max values from summary table
    max_increase = WorksheetFunction.Max(Range("K:K"))
    max_decrease = WorksheetFunction.Min(Range("K:K"))
    max_volume = WorksheetFunction.Max(Range("L:L"))
    
    
    'populate cells with max values
    Cells(2, 16).Value = max_increase
    Cells(3, 16).Value = max_decrease
    Cells(4, 16).Value = max_volume
    
    For j = 2 To 70926
    
        If Cells(j, 11).Value = Cells(2, 16).Value Then
            Cells(2, 15).Value = Cells(j, 9).Value
            
        End If
        
        If Cells(j, 11).Value = Cells(3, 16).Value Then
            Cells(3, 15).Value = Cells(j, 9).Value
        
        End If
        
        If Cells(j, 12).Value = Cells(4, 16).Value Then
            Cells(4, 15).Value = Cells(j, 9).Value
            
        End If
        
            
    Next j
    

End Sub
