Attribute VB_Name = "extra"
Sub extra()
        For Each ws In ActiveWorkbook.Worksheets
                    ' variables.
                    Dim incpertick  As String
                    Dim incperval As Double
                    Dim decpertick As String
                    Dim decperval As Double
                    Dim inctotick As String
                    Dim inctoval As LongLong
                    
                    ' set table with headers info.
                    ws.Range("P1").Value = "Ticker"
                    ws.Range("Q1").Value = "Value"
                    ws.Range("O2").Value = "Greatest % increase"
                    ws.Range("O3").Value = "Greatest % decrease"
                    ws.Range("O4").Value = "Greatest total volume"
                    
            'look for info.
            lastrowTable2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
            maxinc = Application.WorksheetFunction.Max(ws.Range("K:K"))
            maxdec = Application.WorksheetFunction.Min(ws.Range("K:K"))
            maxvol = Application.WorksheetFunction.Max(ws.Range("L:L"))
            'loop trough info.
            For k = 2 To lastrowTable2
                ' Greatest % increase
                  If ws.Cells(k, 11).Value = maxinc Then
                    incpertick = ws.Cells(k, 9).Value
                    incperval = ws.Cells(k, 11).Value
                    ' info to cells
                    ws.Cells(2, 16).Value = incpertick
                    ws.Cells(2, 17).Value = incperval
                  End If
                  
                  ' Greatest % decrease
                  If ws.Cells(k, 11).Value = maxdec Then
                    decpertick = ws.Cells(k, 9).Value
                    decperval = ws.Cells(k, 11).Value
                    ' info to cells
                    ws.Cells(3, 16).Value = decpertick
                    ws.Cells(3, 17).Value = decperval

                End If

                   ' Greatest total volume
                  If ws.Cells(k, 12).Value = maxvol Then
                    inctotick = ws.Cells(k, 9).Value
                    inctoval = ws.Cells(k, 12).Value
                    ' info to cells
                    ws.Cells(4, 16).Value = inctotick
                    ws.Cells(4, 17).Value = inctoval
                  End If
            Next k
            ' Cell formatting.
            ws.Columns("O:Q").Columns.AutoFit
            ws.Range("Q2:Q3").NumberFormat = "0.00%"
            ws.Range("Q4").NumberFormat = "#,###,##0"
        Next ws
End Sub
