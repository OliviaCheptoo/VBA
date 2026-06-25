Sub PullFromNew()

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

    ' Unmerge entire New sheet first
    wsNew.Cells.UnMerge

    lastRowNew = wsNew.Cells(wsNew.Rows.Count, 1).End(xlUp).Row
    lastRowWork = wsWork.Cells(wsWork.Rows.Count, 4).End(xlUp).Row

    Dim colMap(0 To 22, 0 To 1) As Integer

    ' Dec 2025
    colMap(0, 0) = 7: colMap(0, 1) = 7
    colMap(1, 0) = 8: colMap(1, 1) = 8
    colMap(2, 0) = 9: colMap(2, 1) = 9

    ' Jan 2026
    colMap(3, 0) = 10: colMap(3, 1) = 10
    colMap(4, 0) = 11: colMap(4, 1) = 11
    colMap(5, 0) = 12: colMap(5, 1) = 12

    ' Feb 2026
    colMap(6, 0) = 13: colMap(6, 1) = 13
    colMap(7, 0) = 14: colMap(7, 1) = 14
    colMap(8, 0) = 15: colMap(8, 1) = 15

    ' Mar 2026
    colMap(9, 0) = 16: colMap(9, 1) = 16
    colMap(10, 0) = 17: colMap(10, 1) = 17
    colMap(11, 0) = 18: colMap(11, 1) = 18

    ' Apr 2026
    colMap(12, 0) = 19: colMap(12, 1) = 19
    colMap(13, 0) = 20: colMap(13, 1) = 20
    colMap(14, 0) = 21: colMap(14, 1) = 21

    ' May 2026
    colMap(15, 0) = 22: colMap(15, 1) = 22
    colMap(16, 0) = 23: colMap(16, 1) = 23
    colMap(17, 0) = 24: colMap(17, 1) = 24

    ' Jun 2026 - Purchase only
    colMap(18, 0) = 25: colMap(18, 1) = 25

    ' Jul 2026 - Purchase and GRL only
    colMap(19, 0) = 28: colMap(19, 1) = 26
    colMap(20, 0) = 30: colMap(20, 1) = 27

    ' Aug 2026 - Purchase only
    colMap(21, 0) = 31: colMap(21, 1) = 28

    ' Sep 2026 - Purchase only
    colMap(22, 0) = 34: colMap(22, 1) = 29

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

                    ' Copy value only
                    dstCell.Value = srcCell.Value

                    ' Copy fill colour only — leave font, borders, number format intact
                    dstCell.Interior.Color = srcCell.Interior.Color
                    dstCell.Interior.Pattern = srcCell.Interior.Pattern

                Next i

                Exit For

            End If

        Next newRow

NextWorkRow:
    Next workRow

    Application.ScreenUpdating = True
    MsgBox "Done! Data pulled from New to Workings.", vbInformation

End Sub
