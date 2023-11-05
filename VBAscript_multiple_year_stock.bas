Attribute VB_Name = "VBAscript_multiple_year_stock"
Sub Multiple_year_stock_data()

For Each ws In Worksheets

    ' Set an initial variable for holding the ticker
    Dim ticker As String
  
    ' Set an initial variable for holding the volume
    Dim volume As Double
    volume = 0
    
    ' Set an initial variable for holding the open price
    Dim open_price As Double
    open_price = 0
    
    ' Set an initial variable for holding the close price
    Dim close_price As Double
    close_price = 0
    
    ' Keep track of the location for each ticker in the table
    Dim Table_Row As Integer
    Table_Row = 2
          
    ' Set up variable for Greatest % increase as Double
    Dim Greatest_Per_Inc As Double
    
    ' Set up variable for Greatest % Decrease as Double
    Dim Greatest_Per_Dec As Double
    
    ' Set up variable for Greatest Total Volume as Double
    Dim Greatest_Tot_Vol As Double
        
    ' Setup initial variable for last row
    Dim lastrow As Long
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Setup variable for match function for greatest percent increase
    Dim match_gpi As Integer
    
    ' Setup variable for match function for greatest percent decrease
    Dim match_gpd As Integer
    
    ' Setup variable for match function for greatest total volume
    Dim match_gtv As Integer
    
    ' Cell M1 is named Ticker
    ws.Cells(1, 13) = "Ticker"
   
    ' Cell N1 is named Yearly Change
    ws.Cells(1, 14) = "Yearly Change"
    
    ' Cell O1 is named Percent Change
    ws.Cells(1, 15) = "Percent Change"
    
    ' Cell P1 is named Total Stock Volume
    ws.Cells(1, 16) = "Total Stock Volume"
    
    ' Cell S2 is named Greatest % increase
    ws.Cells(2, 19) = "Greatest % Increase"
    
    ' Cell S3 is named Greatest % Decrease
    ws.Cells(3, 19) = "Greatest % Decrease"
    
    ' Cell S4 is named Greatest Total Volume
    ws.Cells(4, 19) = "Greatest Total Volume"
    
    ' Cell T1 is named Ticker
    ws.Cells(1, 20) = "Ticker"
    
    ' Cell U2 is named Value
    ws.Cells(1, 21) = "Value"
    
    ' Loop through all stock tickers
    For t = 2 To lastrow
  
        ' Check if we are still within the same stock ticker, if it is not...
        If ws.Cells(t - 1, 1).Value <> ws.Cells(t, 1).Value Then
            
            ' Set the Stock Ticker
            ticker = ws.Cells(t, 1).Value
            
            ' Add to the Volume
            volume = volume + ws.Cells(t, 7).Value
                       
            ' Set value for row new ticker begins
            TickerStart = t
                       
            ' Set the Open Price
            open_price = ws.Cells(TickerStart, 3).Value
                        
        ElseIf ws.Cells(t + 1, 1).Value <> ws.Cells(t, 1).Value Then
        
            ' Set the Close Price
            close_price = ws.Cells(t, 6).Value
            
            ' Calculate Yearly Change close price - open price
            yearly_change = close_price - open_price
           
            
            ' Calculate Percent Change if open price is zero by looping through open prices for ticker until next non zero value
                If ws.Cells(TickerStart, 3).Value = 0 Then
                        
                   ' Loops through row where ticker starts to row when ticker changes
                   For findrow = TickerStart To t
                    
                    ' Compares values in open price cells starting from ticker start to where ticker ends
                    If ws.Cells(findrow, 3).Value <> 0 Then
                    
                        ' Updates the tickerstart variable to the row that satifies condition and calculates percent change using value
                        TickerStart = findrow
                                              
                        Percent_Change = ((close_price - ws.Cells(TickerStart, 3).Value) / ws.Cells(TickerStart, 3).Value)
                        
                        Exit For
                    End If
                    
                   Next findrow
                    
                ' If the open price is not zero it calculates percent change
                Else
                                     
                    Percent_Change = ((yearly_change) / open_price)
                    
                End If
                
            ' Add to the Volume
            volume = volume + ws.Cells(t, 7).Value
    
            ' Print the ticker in the Summary Table
            ws.Range("M" & Table_Row).Value = ticker
            
            ' Print the Yearly change in Summary Table
            ws.Range("N" & Table_Row).Value = yearly_change
            
            ' Check if yearly change is positive
                If yearly_change > 0 Then
                    
                    ' Format cell containing yearly change to green if value is positive
                    ws.Cells(Table_Row, 14).Interior.ColorIndex = 4
                    
                Else
                    
                    ' Format cell containing yearly change to red if value if negative
                    ws.Cells(Table_Row, 14).Interior.ColorIndex = 3
                   
                ' End if then loop
                End If
                
            
            ' Print the Yearly chnage in Summary Table
            ws.Range("O" & Table_Row).Value = FormatPercent(Percent_Change)
    
    
            ' Print the Total Stock Volume in Summary Table
            ws.Range("P" & Table_Row).Value = volume
    
            ' Add one to the table row
            Table_Row = Table_Row + 1
    
            ' Reset Volume Total
            volume = 0
    
        ' If the cell immediately following a row is the same brand...
        Else
    
            ' Add to the Volume
            volume = volume + ws.Cells(t, 7).Value
        
              
        End If
    
    Next t
    
    ' Setup initial variable for last row
    Dim summ_lastrow As Double
    summ_lastrow = ws.Cells(Rows.Count, 13).End(xlUp).Row
    
    ' Finds Greatest percent increase using max function and assigns it to variable
    Greatest_Per_Inc = WorksheetFunction.Max(ws.Range("O2:O" & summ_lastrow))
        
    ' Prints the Greatest % Increase and formats as a percent
    ws.Cells(2, 21).Value = FormatPercent(Greatest_Per_Inc)
    
    ' Finds the Greatest % Decrease and assigns it to variable
    Greatest_Per_Dec = WorksheetFunction.Min(ws.Range("O2:O" & summ_lastrow))
    
    ' Prints the Greatest % Decrease and formats as a percent
    ws.Cells(3, 21).Value = FormatPercent(Greatest_Per_Dec)
    
    ' Finds the Greatest Total Volume and assigns it to a variable
    Greatest_Tot_Vol = WorksheetFunction.Max(ws.Range("P2:P" & summ_lastrow))
    
    ' Prints the Greatest Total Volume
    ws.Cells(4, 21).Value = Greatest_Tot_Vol
    
    ' Uses Match function to find the row in Summary table with greatest percent increase
    match_gpi = WorksheetFunction.Match(Greatest_Per_Inc, ws.Range("O2:O" & summ_lastrow), 0)
    
    ' Enter Match Ticker into cell
    ws.Cells(2, 20).Value = ws.Range("M" & match_gpi + 1).Value
    
    ' Uses Match function to find the row in Summary table with greatest percent decrease
    match_gpd = WorksheetFunction.Match(Greatest_Per_Dec, ws.Range("O2:O" & summ_lastrow), 0)
    
    ' Enter Match Ticker into cell
    ws.Cells(3, 20).Value = ws.Range("M" & match_gpd + 1).Value
    
    ' Use match Funtion to find the roe in Summay table with the Greatest Total Volume
    match_gtv = WorksheetFunction.Match(Greatest_Tot_Vol, ws.Range("P2:P" & summ_lastrow), 0)
    
    ' Enter Match Ticker into cell
    ws.Cells(4, 20).Value = ws.Range("M" & match_gtv + 1).Value
    
    ' Autofits the data in the columns
    ws.Range("A:U").Columns.AutoFit
   
Next ws

End Sub
