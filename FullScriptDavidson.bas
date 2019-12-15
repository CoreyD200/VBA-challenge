Attribute VB_Name = "FullScriptDavidson"
Sub StockSummary()

'start loop to move through all of the worksheets in the workbook
For Each ws In Worksheets
 
    'Create variable to count the number of shareholders per Ticker
    Dim ShareHolderCount As Double

    'Create variable to roll the total of each shareholder number of shares
    Dim TickerTotal As Double
    Dim Ticker As String

    'create variables for 1st stock price of the year and final stock price of the year
    'encountered stackoverflow with "double" and had to change to Long
    Dim Startprice As Long
    Dim EndPrice As Double

    'keep track of the location for each ticker in the summary table
    Dim Summary_Table_Row As Integer

    'Assign starting values
    ShareHolderCount = 0
    TickerTotal = 0
    Summary_Table_Row = 2

    'get row count and last row and summary row count
    Dim RowCount As Double
    Dim lastrow As Double
    Dim lastSummaryRow As Double
    
     'figure out the row number for the last row in each sheet
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
 
   'figure out the row number for the last row in each sheet
    lastSummaryRow = Cells(Rows.Count, 10).End(xlUp).Row
   
    
    'set row count to '2' to accomodate headers
    RowCount = 2


    'assign a column headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Amount"
    
    
    'assign column formatting
    ws.Range("K1").NumberFormat = "0.00%"
    
    
    'loop through all stocks to gather all of the information needed for the summary table
    For i = 2 To lastrow
        
        'comapre value in the "ticker" column to the one below it to find when one ends and the next one starts
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'Set the ticker name
            Ticker = ws.Cells(i, 1).Value
        
            'Add the number of shares
            TickerTotal = TickerTotal + ws.Cells(i, 7).Value
        
            'figure the final stockprice of the year
            EndPrice = ws.Cells(i, 6).Value
            
            'figue the initial start price of the year (both start and end date assume the dates are sorted with 1st date on the top and end dates are on the bottom)
            Startprice = ws.Cells(i - ShareHolderCount, 3).Value
              
            'put the ticker symbol in the next row on column I
            ws.Range("I" & Summary_Table_Row).Value = Ticker
                    
            'put the Total Number of shares in the next row on column L
            ws.Range("L" & Summary_Table_Row).Value = TickerTotal
                                   
            'put $amount changed from beginning of year to end of year in the next row on column J
            ws.Range("J" & Summary_Table_Row).Value = EndPrice - Startprice
            
            'convert the $ amount changed to a percentage gained/lost in column K
                'Add next line to eliminate division by zero error
                If Startprice <> 0 Then
                ws.Range("K" & Summary_Table_Row).Value = 100 * ((EndPrice - Startprice) / Abs(Startprice))
               
                
                End If
            
            'reset tickertotal
            TickerTotal = 0
            
            'reset ShareholderCount
            ShareHolderCount = 0
            
            'go to next line in Summary table
            Summary_Table_Row = Summary_Table_Row + 1
        
            'Debug print commands to see the start and end values for each ticker, used to double check the correct values were being generated
            'Debug.Print Ticker
            'Debug.Print Startprice
            'Debug.Print EndPrice
        
        Else
        
            'Set the ticker total to the value in the first row of of the ticker
            TickerTotal = TickerTotal + ws.Cells(i, 7).Value
            'incriment shareholderCount
            ShareHolderCount = ShareHolderCount + 1
            
        'the End If makes sense to me but the one above does not. I dont understand how there are 2 End If for a single IF
        End If
            
   'loops to next i
   Next i


    ' Start loop for formatting of the summary table
    For j = 2 To lastSummaryRow

        If ws.Cells(j, 10).Value >= 0 Then
        
           ws.Cells(j, 10).Interior.ColorIndex = 4
           ws.Cells(j, 11).Font.ColorIndex = 10
        
        Else
            ws.Cells(j, 10).Interior.ColorIndex = 3
            ws.Cells(j, 11).Font.ColorIndex = 3
        End If

    Next j
    
    ws.Range("J1").EntireColumn.AutoFit
    ws.Range("K1").EntireColumn.AutoFit
    ws.Range("L1").EntireColumn.AutoFit

    ' The "0.00% formatting not working for some reason. If you go in and type over the cell the formatting will take place.
    ' attempted to move it up in the program but that did not work either
    ws.Range("K1").NumberFormat = "0.00%"
    
    'insert text for greatest and least changes and greatest total volume
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"


    ws.Cells(2, 16).Value = Application.WorksheetFunction.Max(ws.Range("K:K"))
    ws.Cells(3, 16).Value = Application.WorksheetFunction.Min(ws.Range("K:K"))
    ws.Cells(4, 16).Value = Application.WorksheetFunction.Max(ws.Range("L:L"))
    
    Set tickerMaxx = ws.Range("K:K").Find(ws.Cells(2, 16).Value).Offset(0, -2)
    ws.Cells(2, 15) = tickerMaxx

    Set tickerMin = ws.Range("K:K").Find(ws.Cells(3, 16).Value).Offset(0, -2)
    ws.Cells(3, 15) = tickerMin
    
    Set tickerVol = ws.Range("L:L").Find(ws.Cells(4, 16).Value).Offset(0, -3)
    ws.Cells(4, 15) = tickerVol


    'format last few columns
    ws.Range("N1").EntireColumn.AutoFit
    ws.Range("O1").EntireColumn.AutoFit
    ws.Range("P1").EntireColumn.AutoFit

'loops to next worksheet
Next ws
       
End Sub


