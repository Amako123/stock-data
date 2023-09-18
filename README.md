# stock-dataSub StockData()
'Define

Dim a As Integer
Dim ws_num As Integer
Dim starting_ws As Worksheet
Dim ticker As String
Dim volume As Double
Dim yearopen As Double
Dim yearclose As Double
Dim yearlychange As Double
Dim percentchange As Double
Dim SummaryTable_Row As Integer
Dim Row As Integer
    

ws_num = ThisWorkbook.Worksheets.Count

'looping through worksheet
For a = 1 To ws_num

'has it work only on the current workbook
ThisWorkbook.Worksheets(a).Activate

' headers
 Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

'setup integers
Summary_Table_Row = 2
'Row = ActiveSheet.UsedRange.Rows.Count


'loop
For i = 2 To ActiveSheet.UsedRange.Rows.Count

     If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
           
         ticker = Cells(i, 1).Value
           vol = Cells(i, 7).Value
        
         year_open = Cells(i, 3).Value
            year_close = Cells(i, 6).Value

          yearly_change = year_close - year_open
          percent_change = year_close / year_open
        
        'insert values into summary
          Cells(Summary_Table_Row, 9).Value = ticker
           Cells(Summary_Table_Row, 10).Value = yearly_change
            Cells(Summary_Table_Row, 11).Value = percent_change
            Cells(Summary_Table_Row, 12).Value = vol
         Summary_Table_Row = Summary_Table_Row + 1

         vol = 0

     End If



'finish loop
    Next i
'format K columns to percent
Columns("K").NumberFormat = "0.00%"

' next sheet
    Next a

End Sub
