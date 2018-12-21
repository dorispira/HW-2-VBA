Sub StockAnalysis():

  ' Set an initial variable for holding the brand name
  Dim stockName As String
  Dim i As Long
  Dim s1, s2, s3 As Worksheet
  
  Set s1 = Sheets("2014")

  
  ' Set an initial variable for holding the total
  Dim stockTotal As Double
  stockTotal = 0

  ' Keep track of the location for each stock
    Dim StockTableRow As Integer
    StockTableRow = 2

    s1.Activate
    Range("J1").Value = "Ticker"
    Range("K1").Value = "Total Stock Volume"

    lastRow = Cells(Rows.Count, 1).End(xlUp).Row

  ' Loop through all stocks

  For i = 2 To lastRow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Brand name
      stockName = Cells(i, 1).Value

      ' Add to the Brand Total
      stockTotal = stockTotal + Cells(i, 7).Value

      ' Print the CC Brand in the Summary Table
      Range("J" & StockTableRow).Value = stockName

      ' Print the Brand Amount to the Summary Table
      Range("K" & StockTableRow).Value = stockTotal

      ' Add one to the summary table row
     StockTableRow = StockTableRow + 1
      
      ' Reset the Brand Total
      stockTotal = 0

    ' If the cell immediately following a row is the same brand
    Else

      ' Add to the Brand Total
      stockTotal = stockTotal + Cells(i, 7).Value

    End If

  Next i

End Sub