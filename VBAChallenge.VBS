Sub StockData()

Dim LR, tableRow, countTicker As Integer
Dim Total, yrChange, perChange As Double
Dim TickerRange As Range
Dim ws As Worksheet



For Each ws In Worksheets

'Counting the number of rows with data in the first column in each sheet
 LR = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Defining a Range for Column A, that will be used later on for Count If Function
 Set TickerRange = ws.Range("A2: A" & LR)

'Labeling the column headers in the summary table we are going to create
 ws.Cells(1, 9).Value = "Ticker"
 ws.Cells(1, 10).Value = "Yearly Change"
 ws.Cells(1, 11).Value = "Percentage Change"
 ws.Cells(1, 12).Value = "Total Stock Volume"

 'tablerow variable is counting the number of rows in the summary table we are creating; it will start from row 2 as that's where we're going to start writing.
 tableRow = 2
 'Total stock volume variable is also set equal to the first value in row 2 & column 7. This will later be updated for each ticker symbol.
 Total = ws.Cells(2, 7).Value

'Getting the Ticker symbol, Total Stock Volume and write it in the summary table
'Next will be counting the number of each specific tickers and then find the open and close locations; yearly change would then be close minus open value.

 For i = 2 To LR
'Identifying unique ticker symbols
  If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
  
   ws.Cells(tableRow, 9).Value = ws.Cells(i, 1).Value
   ws.Cells(tableRow, 12).Value = Total
   Total = ws.Cells(i + 1, 7).Value
  
  'Yearly change is calculated from the opening price at the beginning of a given year to the closing price at the end of that year.
  'Using "count if" function to count the number of ticker symbols which is going to be used to find the location of the first opening day.
  'For Count If function we need the range as the first argument which is column A, and the criteria as the second argument which is the ticker symbols; here I am getting the unique ticker symbols from the summary table.
   countTicker = Application.WorksheetFunction.CountIf(TickerRange, ws.Cells(tableRow, 9).Value)
   yrChange = ws.Cells(i, 6).Value - ws.Cells(i - countTicker + 1, 3).Value
   ws.Cells(tableRow, 10).Value = yrChange
   
  
  'Color formatting positive and negative cells in column 10 (yearly Change)
   If ws.Cells(tableRow, 10).Value >= 0 Then
    ws.Cells(tableRow, 10).Interior.ColorIndex = 4
    
    
   Else
    ws.Cells(tableRow, 10).Interior.ColorIndex = 3
   End If
  
  
  'The percent change from opening price at the beginning of a given year to the closing price at the end of that year
  '**An importnat note here is that since there are some incidences where the initial open value is zero, the division should take palce as long as the denominator, i.e. open value, is not zero.
    If ws.Cells(i - countTicker + 1, 3).Value <> 0 Then
     perChange = yrChange / ws.Cells(i - countTicker + 1, 3).Value
     ws.Cells(tableRow, 11).Value = perChange
     ws.Cells(tableRow, 11).Style = "Percent"
    Else
     ws.Cells(tableRow, 11).Value = 0
    End If
     
   
  
  
   'Updating the ticker count value
   countTicker = 0
  
   tableRow = tableRow + 1
  
  
  
'Calculating the total stock volume of the stock for each ticker symbol
  Else
   Total = Total + ws.Cells(i + 1, 7).Value
 
  End If
 Next i
 
 
Next ws

End Sub



