Attribute VB_Name = "Module1"
Sub Multiple_Year_Stock_Analysis()

'Instructions
 '*****************************************************************************************
    'Create a script that loops through all the stocks for one year and outputs the following information:


'Part 1
'--------------------------------------
    'The ticker symbol

    'Caculate yearly change from the opening price at the beginning of a given year to the closing price at the end of that year

    'Calculate the percentage change from the opening price at the beginning of a given year to the closing price at the end of that year

    'Calculate the total stock volume of the stock

'Part 2
'--------------------------------------
    'Conditional Formatting

'Part 3
'--------------------------------------
    'Summarize Greatest % Increase, Decrease and Total Volume
    

 '*****************************************************************************************

'Part 1


'Apply macro to all worksheets

  For Each ws In Worksheets

'Create variables

    'Set intitial variable for holding ticker symbol
    Dim tickersymbol As String
    
    'Set intial variable for ticker volume
    Dim ticker_volume As Double
    tickervolume = 0
    
    'Set intial variable for holding total stock volume by ticker symbol
    Dim totalstockvolume As Double
        
    'Keep Track of Ticker Symbols in summary table
    Dim Summary_Ticker_Row As Integer
    Summary_Ticker_Row = 2
    
    'Set intitial variable for opening price
    Dim openingprice As Double
    
    'Set intitial variable for closing price
    Dim closingprice As Double
    
    'Set initial variable for yearly change
    Dim yearlychange As Double

    'Set initial variable for holding percentage change by year
    Dim percentagechange As Double
  

    'Set Opening Price
    openingprice = Cells(2, 3).Value
    
'Print Summary Table Headers
    Cells(1, 9).Value = "Ticker Symbol"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percentage Change"
    Cells(1, 12).Value = "Total Stock Volume"


'Create a loop to search tickers symbols
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    For I = 2 To lastrow

  ' Check if we are still within the same ticker
    If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
    
       
       ' Set the Ticker Symbol and calculate total volume
         tickersymbol = Cells(I, 1).Value
         tickervolume = tickervolume + Cells(I, 7).Value
         
       ' Print the Ticker Symbol in the Summary Table
         Range("I" & Summary_Ticker_Row).Value = tickersymbol
         
       ' Print the Total Stock Volume to the Summary Table
         Range("L" & Summary_Ticker_Row).Value = tickervolume

'yearly change

    'Get Closing Price
    closingprice = Cells(I, 6).Value
    
    'Calculate Yearly Change
    yearlychange = (closingprice - openingprice)

    ' Print the Year Change to the Summary Table
      Range("J" & Summary_Ticker_Row).Value = yearlychange

'percentage change
    If (openingprice = 0) Then
    
    percentagechange = 0
    
    Else
    
    percentagechange = yearlychange / openingprice
    
    End If
         
      
    ' Print the Percent Change to the Summary Table
      Range("K" & Summary_Ticker_Row).Value = percentagechange
      Range("K" & Summary_Ticker_Row).NumberFormat = "0.00%"
   
 'Reset row counter
 Summary_Ticker_Row = Summary_Ticker_Row + 1
 
 'Rest Volume
 tickervolume = 0
 
 'Reset opening price
 openingprice = Cells(I + 1, 3)
        
Else


 ' If the cell immediately following a row is the same ticker
    tickervolume = tickervolume + Cells(I, 7).Value

   End If

  Next I
  
 '******************************************************************************************
  
'Part 2
  

lastrow_summary_table = ws.Cells(Rows.Count, 9).End(xlUp).Row
  
'Conditional formatting showing green for positive change and red for negative change
For I = 2 To lastrow_summary_table

    'If the value of the cell is equal to or greater than 90
    If Cells(I, 10).Value >= 0 Then

    'Color code Green
    Cells(I, 10).Interior.ColorIndex = 4

    Else
    'Color code red
    Cells(I, 10).Interior.ColorIndex = 3

    End If

        Next I

 '*****************************************************************************************

'Part 3


'Print Summary Table Headers
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker Symbol"
    ws.Cells(1, 17).Value = "Value"
    
For I = 2 To lastrow_summary_table


'Return Greatest % Increase
If ws.Cells(I, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrow)) Then
ws.Cells(2, 16).Value = ws.Cells(I, 9).Value
ws.Cells(2, 17).Value = ws.Cells(I, 11).Value
ws.Cells(2, 17).NumberFormat = "0.00%"

'Return Greatest % Decrease
ElseIf ws.Cells(I, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrow)) Then
ws.Cells(3, 16).Value = ws.Cells(I, 9).Value
ws.Cells(3, 17).Value = ws.Cells(I, 11).Value
ws.Cells(3, 17).NumberFormat = "0.00%"


'Return Greatest Total Volume
ElseIf ws.Cells(I, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrow)) Then
ws.Cells(4, 16).Value = ws.Cells(I, 9).Value
ws.Cells(4, 17).Value = ws.Cells(I, 12).Value

        End If
    
    Next I

Next ws

End Sub



 

    





