Attribute VB_Name = "HW_module1"
Sub vba()
'* Create a script that will loop through all the stocks for one year and output the following information.
        '  * The ticker symbol.
        '  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
        '  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
        '  * The total stock volume of the stock.
'* You should also have conditional formatting that will highlight positive change in green and negative change in red.
' aknowledments to CP & EB for all the help in both recovering my lost data; loop for price reference.
    
    For Each ws In ActiveWorkbook.Worksheets ' loopfor ws.
                  ws.Range("I1").Value = "ticker"
                  ws.Range("J1").Value = "yearly_change"
                  ws.Range("K1").Value = "yearly_percentage"
                  ws.Range("L1").Value = "total_stock_volume"
                  
                  'variables
                  Dim ticker As String
                  Dim yopen As Double
                  Dim yclose As Double
                  Dim yearly_percentage As Double
                  Dim total_stock_volume As LongLong 'variable for Total Volume/ set counter to zero.
                  total_stock_volume = 0
                  firstTickerRow = 2 'set ticker row open.
                  Dim summary_table_row As Long 'tanbe row and starting point.
                  summary_table_row = 2
                  'last row for A.
                   lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
                   ' loop trough tickers.
                   For i = 2 To lastrow
                              ' check if next cell is different.
                              If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                                      ticker = ws.Cells(i, 1).Value ' set ticker name.
                                      total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value ' add total to each ticker.
                                      ' First open price non-zero.
                                      If total_stock_volume = 0 Then
                                           ws.Range("J" & summary_table_row).Value = 0
                                           ws.Range("K" & summary_table_row).Value = 0
                                      Else
                                            If ws.Cells(firstTickerRow, 3).Value = 0 Then
                                                For x = firstTickerRow To i
                                                    If ws.Cells(x, 3).Value <> 0 Then
                                                        firstTickerRow = x
                                                        Exit For
                                                    End If
                                                Next x
                                            End If
                                            
                                            'open close value reference.
                                            yopen = ws.Cells(firstTickerRow, 3).Value
                                            yclose = ws.Cells(i, 6).Value
                                            ' changes.
                                            yearly_change = yclose - yopen
                                            yearly_percentage = yearly_change / yopen
                                            ' Display Yearly Change (Price) & (Percent).
                                            ws.Range("J" & summary_table_row).Value = yearly_change
                                            ws.Range("K" & summary_table_row).Value = yearly_percentage
                                            firstTickerRow = i + 1
                                      End If
                                    ' Display new ticker name.
                                    ws.Range("I" & summary_table_row).Value = ticker
                                    ' Display Total Stock Volume.
                                    ws.Range("L" & summary_table_row).Value = total_stock_volume
                                    ' increment table.
                                    summary_table_row = summary_table_row + 1
                                    ' reset totals.
                                    total_stock_volume = 0
                                    icounter = 0
                              Else
                                ' total volume stock added & counter.
                                total_volume_stock = total_volume_stock + ws.Cells(i, 7).Value
                                icounter = icounter + 1
                              End If
                   Next i
                   
                     ' Autofit new table.
                     ws.Columns("I:L").Columns.AutoFit
                     ' formatting new table.
                     ws.Columns("K").NumberFormat = "0.00%"
                     ws.Columns("L").NumberFormat = "#,###,##0"
                    ' conditional formatting.
                    ' color change.
                    pchange = 4
                    nchange = 3
                    'last row fro summary table.
                    lastRow2 = ws.Cells(Rows.Count, 11).End(xlUp).Row
                    ' iteration for color change depending on value.
                    For m = 2 To lastRow2
                            If ws.Cells(m, 10).Value >= 0 Then
                                    ws.Cells(m, 10).Interior.ColorIndex = pchange
                    
                    'in case o negative change.
                    
                            ElseIf ws.Cells(m, 10).Value < 0 Then
                                    ws.Cells(m, 10).Interior.ColorIndex = nchange
                                    
                    
                            End If
                    Next m
    Next ws
            
End Sub
