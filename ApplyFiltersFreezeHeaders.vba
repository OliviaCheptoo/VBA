Sub ApplyFiltersAndFreezeHeaders()
    Dim ws As Worksheet
    Dim lastCol As Long

    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Sheets
        ' Find the last column in the sheet (based on the first row)
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        
        ' Check if AutoFilter is already applied
        If Not ws.AutoFilterMode Then
            ' Apply AutoFilter to the header row
            ws.Rows(1).AutoFilter
        End If
        
        ' Freeze the top row (header row)
        ws.Activate ' Activate the worksheet
        ws.Range("A2").Select ' Select cell A2 to freeze the top row
        ActiveWindow.FreezePanes = True
    Next ws
    
    MsgBox "Filters applied and top rows frozen for all sheets.", vbInformation
End Sub
