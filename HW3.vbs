Sub StockAnalysis()

  ' Set a variable for specifying the column of interest
  
  Dim column As Integer
  column = 1


'For Each ws in Worksheets

'Open the WS.'

Cells(1, 8).Value = "ticker"  'Column H
Cells(1, 9).Value = "total stock volume" 'Column I
CeLLS(1, 10).Value = "Yearly stock price change" 'Column J
Cells(1, 11).Value = "Yearly stock price change (%)"  'Column K

'Easy and Medium
Startstockprice = Cell(2, column+2).Value
j = 2

  SUM = 0

  LastRow = Cells(Rows.Count, 1).End(xlUp).Row
 

  For i = 2 To LastRow

         'If ws.Cells(i+1, column).Value = ws.Cells(i, column).Value Then

         If Cells(i+1, column).Value = Cells(i, column).Value Then

            SUM = Cells(i, column).Value + SUM  'computing total stock volume

         else If Cells(i + 1, column).Value <> Cells(i, column).Value Then

            Cells(j, 8).Value = Cells (i, column).Value 'ticker
            
            Cells(j, 9).Value = SUM 'total stock volume

            EndStockprice = Cells(i, column+5).Value

            Cells(j, 10).Value = StartStockprice - EndStockprice 'Yearly Change
   
            Cells(j, 11).Value = ((StartStockprice-EndStockprice)/StartStockprice)*100 'Percent Change

            StartStockprice =Cells(i+1, column+2).Value

            j = j + 1

         End if

  Next i


'Hard
LastRow = j

j=2,

MaxStock = 0
YearMax = 0
PercentChangeMax = 0

while j < LastRow+1

if Cells(j, 9).Value > MaxStock

MaxStock = Cells(j, 9).Value
m=j


if Cells(j, 10).Value > YearMax

YearMax = Cells(j, 10).Value 
mm=j

if Cells(j, 11).Value > PercentChangeMax

PercentChangeMax = Cells(j, 11).Value        'Maximum percent "increase" in a Year
mmm=j


j = j + 1

End while

Cells(2, 12).Value = "MaxStock" 'Column L
Cells(3, 12).Value = "YearMax"  
Cells(4, 12).Value = "PercentIncreaseMax"

Cells(2, 13).Value = Cells(m, 9).Value    'Column M
Cells(3, 13).Value = Cells(mm, 10).Value  
Cells(4, 13).Value = Cell(mmm, 11).Value

Cells(2, 14).Value = MaxStock   'Column N
Cells(3, 14).Value = YearMax
Cells(4, 14).Value = PercentChangeMax

'Next ws

End Sub 