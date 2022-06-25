Attribute VB_Name = "Module1"
'this code will track the total stock volume of each ticker
Sub tickerVolume():

    For Each ws In Worksheets
    
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim ticker As String
    Dim summaryTable As Integer
    Dim priceRow As Long
    
    'BONUS
    'set variables for greatest increase and decrease for total volume
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestTotal As Double
    Dim greatestTickerIncrease As String
    Dim greatestTickerDecrease As String
    Dim greatestTickerTotal As String
    
    'title of each column
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    'variable for the ticker
    ticker = ""
    
    'variable for total stock volume
    totalStockVolume = 0
    
    'variable to keep track of location of different open prices
    priceRow = 2
    
    'variable for summary table row
    summaryTable = 2
    
    'use function to find last row in the sheet
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'loop from row 2 in column A out to the last row
    For Row = 2 To lastRow
    
        'check to see if ticker changes
        If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then
            
            'if ticker changes and setting the ticker type
            ticker = ws.Cells(Row, 1).Value
            
            'add the last volume from the row
            totalStockVolume = totalStockVolume + ws.Cells(Row, 7).Value
            
            'add ticker to I column in summary table row
            ws.Cells(summaryTable, 9).Value = ticker
            
            'add total volume to L column in summary table row
            ws.Cells(summaryTable, 12).Value = totalStockVolume
            
            'Calculate yearly change and percent change
            openPrice = ws.Range("C" & priceRow).Value
            closePrice = ws.Range("F" & Row).Value
            yearlyChange = closePrice - openPrice
            
         If openPrice = 0 Then
                percentChange = 0
            
            Else
                
                percentChange = yearlyChange / openPrice
                    
            'print values of yearly change and percent change
            ws.Range("J" & summaryTable).Value = Format(yearlyChange, "0.00")
            ws.Range("K" & summaryTable).Value = Format(percentChange, "0.00%")
            
                'conditional formatting (green positive, red negative)
                If ws.Range("J" & summaryTable).Value > 0 Then
                   ws.Range("J" & summaryTable).Interior.ColorIndex = 4
                
                Else
                
                   ws.Range("J" & summaryTable).Interior.ColorIndex = 3
                
                End If
            
            'go to next summary table row (add 1 on to the value of the summary table)
            summaryTable = summaryTable + 1
            priceRow = Row + 1
            
            'reset totals to 0
            totalStockVolume = 0
            openPrice = 0
            closePrice = 0
        
        End If
        
        Else
        
                'if ticker stays the same add to the total volume from the G column
                totalStockVolume = totalStockVolume + ws.Cells(Row, 7).Value
            
        End If

    Next Row
    
    'BONUS
    'set first greatest increase/decrease/total ticker percent change and total stock volume
   greatestIncrease = ws.Range("K2").Value
   greatestDecrease = ws.Range("K3").Value
   greatestTotal = ws.Range("K4").Value
   
   'define last row of ticker
   lastRow = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    'loop through each row of column I to find greatest ticker
    For Row = 2 To lastRow
    
        If ws.Range("K" & Row + 1).Value > greatestIncrease Then
        
            greatestIncrease = ws.Range("K" & Row + 1).Value
            greatestTickerIncrease = ws.Range("I" & Row + 1).Value
            
        ElseIf ws.Range("K" & Row + 1).Value < greatestDecrease Then
            
            greatestDecrease = ws.Range("K" & Row + 1).Value
            greatestTickerDecrease = ws.Range("I" & Row + 1).Value
        
        ElseIf ws.Range("L" & Row + 1).Value > greatestTotal Then
            greatestTotal = ws.Range("L" & Row + 1).Value
            greatestTickerTotal = ws.Range("I" & Row + 1).Value
        
        End If
        
            Next Row
    
            'print all data in respective rows and columns
            ws.Range("P2").Value = greatestTickerIncrease
            ws.Range("P3").Value = greatestTickerDecrease
            ws.Range("P4").Value = greatestTickerTotal
            ws.Range("Q2").Value = Format(greatestIncrease, "0.00%")
            ws.Range("Q3").Value = Format(greatestDecrease, "0.00%")
            ws.Range("Q4").Value = greatestTotal
    
    Next ws
    
End Sub
