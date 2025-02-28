Sub ConvertAndFilterBySpecificDateRange()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim startDate As Date
    Dim endDate As Date
    Dim cell As Range

    ' Define the start and end dates for the filter
    startDate = DateSerial(2024, 8, 1) ' August 1, 2024
    endDate = DateSerial(2024, 12, 31) ' December 31, 2024

    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Sheets
        ' Find the last row in the sheet (based on column J)
        lastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
        ' Skip if only headers are present
        If lastRow < 2 Then GoTo NextSheet
        
        ' Find the last column in the sheet (based on the first row)
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

        ' Ensure Column J is recognized as a date format
        ws.Columns("J").NumberFormat = "mm/dd/yyyy"

        ' Convert text dates to actual date values
        For Each cell In ws.Range("J2:J" & lastRow)
            If Not IsError(cell.Value) And IsDate(cell.Value) Then
                cell.Value = CDate(cell.Value)
            End If
        Next cell

        ' Remove existing filters if present
        If ws.FilterMode Then ws.ShowAllData
        If ws.AutoFilterMode Then ws.AutoFilterMode = False

        ' Apply AutoFilter to show only the specified date range
        ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).AutoFilter Field:=10, _
            Criteria1:=">=" & Format(startDate, "mm/dd/yyyy"), _
            Operator:=xlAnd, _
            Criteria2:="<=" & Format(endDate, "mm/dd/yyyy")

NextSheet:
    Next ws
    
    MsgBox "Converted custom dates and filtered all sheets to show values from August 2024 to December 2024.", vbInformation
End Sub


