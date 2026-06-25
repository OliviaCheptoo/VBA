Sub UpdateSalesAndStocks()

    Dim wsNew As Worksheet
    Dim wsWork As Worksheet
    Dim lastRowNew As Long
    Dim lastRowWork As Long
    Dim newRow As Long
    Dim workRow As Long
    Dim itemCodeWork As String
    Dim itemModelNew As String
    Dim extractedCode As String
    Dim dashPos As Integer
    Dim i As Integer
    Dim srcCell As Range
    Dim dstCell As Range

    Set wsNew = ThisWorkbook.Sheets("New")
    Set wsWork = ThisWorkbook.Sheets("Workings")

    wsNew.Cells.UnMerge

    lastRowNew = wsNew.Cells(wsNew.Rows.Count, 1).End(xlUp).Row
    lastRowWork = wsWork.Cells(wsWork.Rows.Count, 4).End(xlUp).Row

    ' --- COLUMN MAPPING: (New col, Workings col) ---
    ' Jun 2026 Qty Sales    = New col 74 (BV)  ? Workings col 47 (AU)
    ' Sales Total           = New col 80 (CB)  ? Workings col 48 (AV)
    ' Physical Stock        = New col 93 (CO)  ? Workings col 56 (BD)
    ' Sea                   = New col 94 (CP)  ? Workings col 57 (BE)
    ' GRL                   = New col 95 (CQ)  ? Workings col 58 (BF)
    ' Unshipped             = New col 96 (CR)  ? Workings col 59 (BG)

    Dim colMap(0 To 5, 0 To 1) As Integer

    colMap(0, 0) = 74: colMap(0, 1) = 47   ' Jun Sales
    colMap(1, 0) = 80: colMap(1, 1) = 48   ' Sales Total
    colMap(2, 0) = 93: colMap(2, 1) = 56   ' Physical Stock
    colMap(3, 0) = 94: colMap(3, 1) = 57   ' Sea
    colMap(4, 0) = 95: colMap(4, 1) = 58   ' GRL
    colMap(5, 0) = 96: colMap(5, 1) = 59   ' Unshipped

    Application.ScreenUpdating = False

    For workRow = 3 To lastRowWork

        itemCodeWork = Trim(CStr(wsWork.Cells(workRow, 4).Value))

        If itemCodeWork = "" Then GoTo NextWorkRow
        If Not IsNumeric(itemCodeWork) Then GoTo NextWorkRow

        For newRow = 2 To lastRowNew

            itemModelNew = Trim(CStr(wsNew.Cells(newRow, 5).Value))

            dashPos = InStr(itemModelNew, "-")
            If dashPos > 1 Then
                extractedCode = Trim(Left(itemModelNew, dashPos - 1))
            Else
                extractedCode = itemModelNew
            End If

            If extractedCode = itemCodeWork Then

                For i = 0 To UBound(colMap, 1)
                    Set srcCell = wsNew.Cells(newRow, colMap(i, 0))
                    Set dstCell = wsWork.Cells(workRow, colMap(i, 1))

                    ' Value only
                    dstCell.Value = srcCell.Value

                    ' Fill colour only
                    dstCell.Interior.Color = srcCell.Interior.Color
                    dstCell.Interior.Pattern = srcCell.Interior.Pattern

                Next i

                Exit For

            End If

        Next newRow

NextWorkRow:
    Next workRow

    Application.ScreenUpdating = True
    MsgBox "Done! Sales and Stocks updated in Workings.", vbInformation

End Sub
