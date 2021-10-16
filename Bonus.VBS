Sub GreatestValues()
Dim Max, Min, MaxStock As Double
Dim nameTicker As String



For Each ws In Worksheets

'Counting the number of rows in the summary table produced from module 1
 rowTable = ws.Cells(Rows.Count, 9).End(xlUp).Row


'Finfing the greatest % increase and its corresponding ticker name from the summary table
 Max = ws.Cells(2, 11).Value
 nameTicker = ws.Cells(2, 9).Value

 For i = 2 To rowTable
  If ws.Cells(i + 1, 11).Value > Max Then
   Max = ws.Cells(i + 1, 11).Value
   nameTicker = ws.Cells(i + 1, 9).Value
  End If
 
 Next i


'Printing the results in the appropriate cells in the spreadsheet
 ws.Cells(2, 17).Value = Max
 ws.Cells(2, 16).Value = nameTicker
 ws.Cells(2, 17).Style = "Percent"
 
 
 

'Fiding the greatest % decrease and its corresponding ticker name from the summary table
 Min = ws.Cells(2, 11).Value
 nameTicker = ws.Cells(2, 9).Value
 For i = 2 To rowTable
  If ws.Cells(i + 1, 11).Value < Min Then
   Min = ws.Cells(i + 1, 11).Value
   nameTicker = ws.Cells(i + 1, 9).Value
  End If
 
 Next i


'Printing the results in the appropriate cells in the spreadsheet
 ws.Cells(3, 17).Value = Min
 ws.Cells(3, 16).Value = nameTicker
 ws.Cells(3, 17).Style = "Percent"




'Fiding the greatest total stock volume and its corresponding ticker name from the summary table
 MaxStock = ws.Cells(2, 12).Value
 nameTicker = ws.Cells(2, 9).Value
 For i = 2 To rowTable
  If ws.Cells(i + 1, 12).Value > MaxStock Then
   MaxStock = ws.Cells(i + 1, 12).Value
   nameTicker = ws.Cells(i + 1, 9).Value
  End If
 
 Next i


'Printing the results in the appropriate cells in the spreadsheet
 ws.Cells(4, 17).Value = MaxStock
 ws.Cells(4, 16).Value = nameTicker




'Row & column Labels for the results
 ws.Cells(2, 15).Value = "Greatest % Increase"
 ws.Cells(3, 15).Value = "Greatest % Decrease"
 ws.Cells(4, 15).Value = "Greatest Total Volume"
 ws.Cells(1, 16).Value = "Ticker"
 ws.Cells(1, 17).Value = "Value"


 ws.Columns("O:Q").AutoFit

Next ws

End Sub


