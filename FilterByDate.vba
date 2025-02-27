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
        
        ' Find the last column in the sheet (based on the first row)
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

        ' Convert custom date format in column J to actual date values
        For Each cell In ws.Range("J2:J" & lastRow)
            If Not IsError(cell.Value) And IsDate(cell.Value) Then
                cell.Value = CDate(cell.Value)
            End If
        Next cell

        ' Ensure AutoFilter is applied to the first row
        If ws.AutoFilterMode Then
            ws.AutoFilterMode = False ' Turn off autofilter if already applied
        End If
        
        ' Clear existing filters if present
        If ws.FilterMode Then
            ws.ShowAllData
        End If

        ' Apply AutoFilter to show only the specified date range
        ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).AutoFilter Field:=10, _
            Criteria1:=">=" & Format(startDate, "mm/dd/yyyy"), _
            Operator:=xlAnd, _
            Criteria2:="<=" & Format(endDate, "mm/dd/yyyy")
    Next ws
    
    MsgBox "Converted custom dates and filtered all sheets to show values from August 2024 to December 2024.", vbInformation
End Sub
