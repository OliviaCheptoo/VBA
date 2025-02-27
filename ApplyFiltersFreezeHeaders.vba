Sub ApplyFiltersAndFreezeHeaders()
    Dim ws As Worksheet
    Dim lastCol As Long

    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Sheets
        ' Find the last column in the sheet (based on the first row)
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        
        ' Apply AutoFilter to the header row
        ws.Rows(1).AutoFilter
        
        ' Freeze the top row
        ws.Activate
        ws.Rows("2:2").Select
        ActiveWindow.FreezePanes = True
    Next ws
    
    MsgBox "Filters applied and top rows frozen for all sheets.", vbInformation
End Sub

