Attribute VB_Name = "Module1"
Sub info():
    Dim sheet As Worksheet
    For Each sheet In ThisWorkbook.Worksheets
        openStock = sheet.Cells(2, 3).Value
        closeStock = 0
        tickTable = 2
        totalStockVol = 0
        lastRow = sheet.Cells(sheet.Rows.Count, 1).End(xlUp).Row
    
        sheet.Range("I1").Value = "Ticker"
        sheet.Range("J1").Value = "Yearly Change"
        sheet.Range("K1").Value = "Percent Change"
        sheet.Range("L1").Value = "Total Stock Volume"
    
        For i = 2 To lastRow
            If sheet.Cells(i + 1, 1).Value <> sheet.Cells(i, 1).Value Then
                closeStock = sheet.Cells(i, 6).Value
            
                sheet.Cells(tickTable, 9).Value = sheet.Cells(i, 1).Value
                sheet.Range("K" & tickTable).NumberFormat = "0.00%"
                
                sheet.Cells(tickTable, 10).Value = closeStock - openStock
                sheet.Range("J" & tickTable).NumberFormat = "0.00"
                    If sheet.Cells(tickTable, 10).Value < 0 Then
                        sheet.Cells(tickTable, 10).Interior.ColorIndex = 3
                    Else
                        sheet.Cells(tickTable, 10).Interior.ColorIndex = 4
                    End If
                    
                sheet.Cells(tickTable, 11).Value = (closeStock - openStock) / openStock
                sheet.Cells(tickTable, 12).Value = sheet.Cells(i, 7).Value + totalStockVol
            
                openStock = sheet.Cells(i + 1, 3).Value
                tickTable = 1 + tickTable
                totalStockVol = 0
            Else
                totalStockVol = sheet.Cells(i, 7).Value + totalStockVol
         End If
        Next i
      Next sheet
End Sub
