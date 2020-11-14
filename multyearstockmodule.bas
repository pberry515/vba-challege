Attribute VB_Name = "Module1"
Sub Ticker()
'Loop to update all worksheets in the workbook

Dim Ws As Worksheet

    For Each Ws In Worksheets
    
    'Defining variables needed for loop
    
    'Create Variable for TickerSymbol
    Dim TickerSymbol As String
    
    'Create Variable for Ticker_Total_Volume and Count
    Dim Ticker_Total_Volume As Double
    Ticker_Total_Volume = 0
    
    'Set location for each ticker symbol in summary_table
    Dim Summary_Table_row As Integer
    Summary_Table_row = 2
    
    'Create variables for Yearly_Change and Percent_Change
    Dim Open_Price As Double
    Open_Price = Ws.Cells(2, 3).Value
    Dim Close_Price As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    
    'Create Summary Table Headers
    Ws.Cells(1, 9).Value = "Ticker"
    Ws.Cells(1, 10).Value = "Yearly_Change"
    Ws.Cells(1, 11).Value = "Percent_Change"
    Ws.Cells(1, 12).Value = "Total Stock Volume"
    
    'Count the number of rows in the first column
    lastrow = Ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Establish loop through each of the rows by the ticker symbol
    
    For i = 2 To lastrow
    
        If Ws.Cells(i + 1, 1).Value <> Ws.Cells(i, 1).Value Then
        
        TickerSymbol = Ws.Cells(i, 1).Value
        Ticker_Total_Volume = Ticker_Total_Volume + Ws.Cells(i, 7).Value
        
        'Add to summary_table_row
        Ws.Range("I" & Summary_Table_row).Value = TickerSymbol
        
        'Add the volume to each ticker in summary table
        Ws.Range("L" & Summary_Table_row).Value = Ticker_Total_Volume
        Ws.Range("L" & Summary_Table_row).NumberFormat = "#,###"
        
        
        'Define Closing Price
        Close_Price = Ws.Cells(i, 6).Value
        
        'Calculate Yearly_Change
        Yearly_Change = (Close_Price - Open_Price)
        
        'Add Yearly Change to Summary Table
        Ws.Range("J" & Summary_Table_row).Value = Yearly_Change
        
                                
        'Ensure result is not 0 for Percent_Change
        If Open_Price = 0 Then
            Percent_Change = 0
        Else
            Percent_Change = Yearly_Change / Open_Price
            
        End If
        
        'Add Yearly_Change fore each TickerSymbol in the Summary Table
        Ws.Range("K" & Summary_Table_row).Value = Percent_Change
        Ws.Range("K" & Summary_Table_row).NumberFormat = "0.00%"
        
        'Reset the row counter and add one to Summary_Table_Row
        Summary_Table_row = Summary_Table_row + 1
        
        'Reset the trade volume to zero
        Ticker_Total_Volume = 0
                
        'Reset the Opening Price
        Open_Price = Ws.Cells(i + 1, 3)
        
    Else
    
        'Add Volume to trade
        Ticker_Total_Volume = Ticker_Total_Volume + Ws.Cells(i, 7).Value
                
        
    End If
    
    Next i
    
    'Format positive and negative growth
    'Identify last row
    lastrow_summary_table = Ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'Color coding for positive green negative red
    For i = 2 To lastrow_summary_table
    
        If Ws.Cells(i, 10).Value > 0 Then
        
            Ws.Cells(i, 10).Interior.ColorIndex = 10
            
        Else
            Ws.Cells(i, 10).Interior.ColorIndex = 3
            
        End If
        
    Next i
    
    'Review the stock price changes
    
    Ws.Cells(2, 15).Value = "Greatest % Increase"
    Ws.Cells(3, 15).Value = "Greates % Decrease"
    Ws.Cells(4, 15).Value = "Greatest Total Volume"
    Ws.Cells(1, 16).Value = "TickerSymbol"
    Ws.Cells(1, 17).Value = "Value"
    Ws.Cells(i, 15).ColumnWidth = 19
    Ws.Cells(i, 16).ColumnWidth = 13
    Ws.Cells(i, 17).ColumnWidth = 13
    
    
    'Identifying the max Percent_Change and adding to table
    
    For i = 2 To lastrow_summary_table
    
        If Ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(Ws.Range("K2:K" & lastrow_summary_table)) Then
            Ws.Cells(2, 16).Value = Ws.Cells(i, 9).Value
            Ws.Cells(2, 17).Value = Ws.Cells(i, 11).Value
            Ws.Cells(2, 17).NumberFormat = "0.00%"
        
        'Identifying the Min Percent_Change
        
        ElseIf Ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(Ws.Range("K2:K" & lastrow_summary_table)) Then
            Ws.Cells(3, 16).Value = Ws.Cells(i, 9).Value
            Ws.Cells(3, 17).Value = Ws.Cells(i, 11).Value
            Ws.Cells(3, 17).NumberFormat = "0.00%"
            
        'Identifying the Max Vol of the trade
        
        ElseIf Ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(Ws.Range("L2:L" & lastrow_summary_table)) Then
            Ws.Cells(4, 16).Value = Ws.Cells(i, 9).Value
            Ws.Cells(4, 17).Value = Ws.Cells(i, 12).Value
            Ws.Cells(4, 17).NumberFormat = "#,###"
            
            
        End If
       
    Next i
        
Next Ws
      
     
End Sub

