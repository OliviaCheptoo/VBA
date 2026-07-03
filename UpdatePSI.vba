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
    Dim dashPos As Long
    Dim i As Long
    Dim newCol As Long
    Dim workCol As Long
    Dim srcCell As Range
    Dim dstCell As Range
    Dim modelArr As Variant
    Dim codeDict As Object
    Dim mergeState As Variant
    Dim colItemCodeWork As Long
    Dim colItemModelNew As Long
    Dim warnings As String
    Dim colMap() As Long

    ' === HEADER ROW LOCATIONS ===
    Const NEW_HEADER_ROW As Long = 1
    Const WORK_MONTH_ROW As Long = 1   ' e.g. "Dec", "Jan", "Feb"...
    Const WORK_SUB_ROW As Long = 2     ' "P" / "S" / "G"
    ' =============================

    ' === MONTHLY MAPPING - edit this table to add, remove, or change a month ===
    ' Each row: New header text (row 1) | Workings month text (row 1) |
    '           Workings sub text (row 2) | which occurrence to match.
    ' Occurrence only matters for "Dec": it appears twice on Workings - once
    ' for the old Dec block we no longer touch (1st match), and once for the
    ' new Aug-Dec Purchase-only block (2nd match). Everything else is unique,
    ' so leave occurrence at 1.
    ' Each line is its own statement (no line-continuations) - VBA caps a
    ' single continued statement at 24 lines, and this table is longer than
    ' that, so it's built one row at a time instead of one big Array(...).
    Dim mapDef() As Variant
    ReDim mapDef(0 To 25)
    mapDef(0) = Array("Jan 2026 Qty Purchase", "Jan", "P", 1)
    mapDef(1) = Array("Jan 2026 Qty Shipment", "Jan", "S", 1)
    mapDef(2) = Array("Jan 2026 GRL", "Jan", "G", 1)
    mapDef(3) = Array("Feb 2026 Qty Purchase", "Feb", "P", 1)
    mapDef(4) = Array("Feb 2026 Qty Shipment", "Feb", "S", 1)
    mapDef(5) = Array("Feb 2026 GRL", "Feb", "G", 1)
    mapDef(6) = Array("Mar 2026 Qty Purchase", "Mar", "P", 1)
    mapDef(7) = Array("Mar 2026 Qty Shipment", "Mar", "S", 1)
    mapDef(8) = Array("Mar 2026 GRL", "Mar", "G", 1)
    mapDef(9) = Array("Apr 2026 Qty Purchase", "Apr", "P", 1)
    mapDef(10) = Array("Apr 2026 Qty Shipment", "Apr", "S", 1)
    mapDef(11) = Array("Apr 2026 GRL", "Apr", "G", 1)
    mapDef(12) = Array("May 2026 Qty Purchase", "May", "P", 1)
    mapDef(13) = Array("May 2026 Qty Shipment", "May", "S", 1)
    mapDef(14) = Array("May 2026 GRL", "May", "G", 1)
    mapDef(15) = Array("Jun 2026 Qty Purchase", "Jun", "P", 1)
    mapDef(16) = Array("Jun 2026 Qty Shipment", "Jun", "S", 1)
    mapDef(17) = Array("Jun 2026 GRL", "Jun", "G", 1)
    mapDef(18) = Array("Jul 2026 Qty Purchase", "Jul", "P", 1)
    mapDef(19) = Array("Jul 2026 Qty Shipment", "Jul", "S", 1)
    mapDef(20) = Array("Jul 2026 GRL", "Jul", "G", 1)
    mapDef(21) = Array("Aug 2026 Qty Purchase", "Aug", "P", 1)
    mapDef(22) = Array("Sep 2026 Qty Purchase", "Sep", "P", 1)
    mapDef(23) = Array("Oct 2026 Qty Purchase", "Oct", "P", 1)
    mapDef(24) = Array("Nov 2026 Qty Purchase", "Nov", "P", 1)
    mapDef(25) = Array("Dec 2026 Qty Purchase", "Dec", "P", 2)
    ' =============================================================================
    ' To add a month: bump ReDim mapDef(0 To 25) up by however many new rows
    ' you're adding, then add the new mapDef(n) = Array(...) lines.
    ' To remove a month: delete its line(s) and renumber the remaining
    ' indices (0, 1, 2...) with no gaps, then shrink the ReDim bound to match.

    ReDim colMap(0 To UBound(mapDef), 0 To 1)

    On Error GoTo CleanFail

    Set wsNew = ThisWorkbook.Sheets("New")
    Set wsWork = ThisWorkbook.Sheets("Workings")

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' Only unmerge if the sheet actually has merged cells - skip the
    ' (expensive, whole-sheet) UnMerge call otherwise.
    mergeState = wsNew.UsedRange.MergeCells
    If IsNull(mergeState) Or mergeState = True Then
        wsNew.Cells.UnMerge
    End If

    ' --- Locate the item code / item model columns by header text ---
    colItemModelNew = FindColByHeader1Row(wsNew, NEW_HEADER_ROW, "ITEM MODEL")
    colItemCodeWork = FindColByHeader2Row(wsWork, WORK_MONTH_ROW, WORK_SUB_ROW, "", "ITEM CODE")
    If colItemCodeWork = 0 Then
        colItemCodeWork = FindColByHeader1Row(wsWork, WORK_MONTH_ROW, "ITEM CODE")
    End If
    If colItemModelNew = 0 Or colItemCodeWork = 0 Then
        MsgBox "Could not find 'ITEM MODEL' on New or 'ITEM CODE' on Workings." & vbCrLf & _
               "Check the header text/row hasn't changed, then re-run.", vbCritical
        GoTo CleanFail
    End If

    ' --- Resolve each monthly column pair by header text ---
    warnings = ""
    For i = 0 To UBound(mapDef)
        newCol = FindColByHeader1Row(wsNew, NEW_HEADER_ROW, CStr(mapDef(i)(0)))
        workCol = FindColByHeader2Row(wsWork, WORK_MONTH_ROW, WORK_SUB_ROW, CStr(mapDef(i)(1)), CStr(mapDef(i)(2)), CLng(mapDef(i)(3)))
        colMap(i, 0) = newCol
        colMap(i, 1) = workCol
        If newCol = 0 Or workCol = 0 Then
            warnings = warnings & "- " & CStr(mapDef(i)(0)) & " (New col=" & newCol & ", Workings col=" & workCol & ")" & vbCrLf
        End If
    Next i

    If warnings <> "" Then
        If MsgBox("Some columns weren't found and will be skipped:" & vbCrLf & warnings & vbCrLf & "Continue anyway?", vbYesNo + vbExclamation) = vbNo Then
            GoTo CleanFail
        End If
    End If

    lastRowNew = wsNew.Cells(wsNew.Rows.Count, 1).End(xlUp).Row
    lastRowWork = wsWork.Cells(wsWork.Rows.Count, colItemCodeWork).End(xlUp).Row

    ' --- Build a one-time lookup: extracted item code -> row on "New" ---
    Set codeDict = CreateObject("Scripting.Dictionary")

    If lastRowNew >= 2 Then
        modelArr = wsNew.Range(wsNew.Cells(2, colItemModelNew), wsNew.Cells(lastRowNew, colItemModelNew)).Value
        For i = 1 To UBound(modelArr, 1)
            itemModelNew = Trim(CStr(modelArr(i, 1)))
            dashPos = InStr(itemModelNew, "-")
            If dashPos > 1 Then
                extractedCode = Trim(Left(itemModelNew, dashPos - 1))
            Else
                extractedCode = itemModelNew
            End If
            If extractedCode <> "" Then
                ' first match wins, same as Exit For did in the original
                If Not codeDict.Exists(extractedCode) Then
                    codeDict.Add extractedCode, i + 1  ' array is offset by header row
                End If
            End If
        Next i
    End If

    For workRow = 3 To lastRowWork
        itemCodeWork = Trim(CStr(wsWork.Cells(workRow, colItemCodeWork).Value))
        If itemCodeWork <> "" Then
            If IsNumeric(itemCodeWork) Then
                If codeDict.Exists(itemCodeWork) Then
                    newRow = codeDict(itemCodeWork)
                    For i = 0 To UBound(colMap, 1)
                        If colMap(i, 0) > 0 And colMap(i, 1) > 0 Then
                            Set srcCell = wsNew.Cells(newRow, colMap(i, 0))
                            Set dstCell = wsWork.Cells(workRow, colMap(i, 1))
                            ' Copy value only
                            dstCell.Value = srcCell.Value
                            ' Copy fill colour only - leave font, borders, number format intact
                            dstCell.Interior.Color = srcCell.Interior.Color
                            dstCell.Interior.Pattern = srcCell.Interior.Pattern
                        End If
                    Next i
                End If
            End If
        End If
    Next workRow

CleanFail:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
    Else
        MsgBox "Done! Data pulled from New to Workings.", vbInformation
    End If
End Sub

' Searches a single header row for targetText (case-insensitive, trimmed).
' Returns the column number of the occurrence-th match, or 0 if not found.
Function FindColByHeader1Row(ws As Worksheet, headerRow As Long, targetText As String, Optional occurrence As Long = 1) As Long
    Dim lastCol As Long
    Dim c As Long
    Dim found As Long
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    found = 0
    For c = 1 To lastCol
        If UCase(Trim(CStr(ws.Cells(headerRow, c).Value))) = UCase(targetText) Then
            found = found + 1
            If found = occurrence Then
                FindColByHeader1Row = c
                Exit Function
            End If
        End If
    Next c
    FindColByHeader1Row = 0
End Function

' Searches two stacked header rows (e.g. month row + P/S/G sub-row) for a
' column where BOTH rows match. Pass "" for either text to treat it as a
' wildcard (match any value in that row). Returns the column number of the
' occurrence-th match, or 0 if not found.
Function FindColByHeader2Row(ws As Worksheet, row1 As Long, row2 As Long, text1 As String, text2 As String, Optional occurrence As Long = 1) As Long
    Dim lastCol As Long
    Dim lastCol2 As Long
    Dim c As Long
    Dim found As Long
    Dim v1 As String
    Dim v2 As String
    lastCol = ws.Cells(row1, ws.Columns.Count).End(xlToLeft).Column
    lastCol2 = ws.Cells(row2, ws.Columns.Count).End(xlToLeft).Column
    If lastCol2 > lastCol Then lastCol = lastCol2
    found = 0
    For c = 1 To lastCol
        v1 = Trim(CStr(ws.Cells(row1, c).Value))
        v2 = Trim(CStr(ws.Cells(row2, c).Value))
        If (text1 = "" Or UCase(v1) = UCase(text1)) And (text2 = "" Or UCase(v2) = UCase(text2)) Then
            found = found + 1
            If found = occurrence Then
                FindColByHeader2Row = c
                Exit Function
            End If
        End If
    Next c
    FindColByHeader2Row = 0
End Function
