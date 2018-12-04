Sub stockTickerTotals()

    
Dim lRow As Long
Dim StoockTotalPerSymbol As Double

StoockTotalPerSymbol = 0
' find the row count
lRow = Cells(Rows.Count, 1).End(xlUp).Row

'print the rows to check
MsgBox (lRow)

'variable to store the value of stock ticker
Dim stockTicker As String

'variable to store the value of stock total
Dim stockTotal As Double

For Each ws In Worksheets
   
        ws.Columns("H:H").Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        ws.Range("H1").Value = "Stock Ticker"
        ws.Columns("I:I").Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        ws.Range("I1").Value = "Total"
        
        
        ' Loop through all stock detail rows
        For i = 2 To 101

            ' Check if we are still within the same credit card brand, if it is not...
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

                ' read the stock ticker name
                 stockTicker = Cells(i, 1).Value

                ' Add to the stock ticker total
                stockTotal = stockTotal + Cells(i, 7).Value

                ' Print the stock ticker name
                Range("H" & i).Value = stockTicker

                ' Print the stock total for all the stocks with the same sticker type
                Range("I" & i).Value = stockTotal

                 ' Add one to the row
                    i = i + 1
      
                    ' Reset the stock total
                    stockTotal = 0

            ' If the cell immediately following a row is the same stock...
            Else

            ' Add to the stock Total
            stockTotal = stockTotal + Cells(i, 7).Value

            End If

        Next i
        
  
Next ws


End Sub

                    ' Reset the stock total
                    stockTotal = 0

            ' If the cell immediately following a row is the same stock...
            Else

            ' Add to the stock Total
            stockTotal = stockTotal + Cells(i, 7).Value

            End If

        Next i
        
  
Next ws


End Sub
