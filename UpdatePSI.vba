Option Explicit

Sub PullFromNew()

    Dim wsNew As Worksheet
    Dim wsWork As Worksheet

    Dim lastRowNew As Long
    Dim lastRowWork As Long

    Dim workRow As Long
    Dim newRow As Long

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

    ' ============================================================
    ' HEADER ROW LOCATIONS
    ' ============================================================

    Const NEW_HEADER_ROW As Long = 1
    Const WORK_MONTH_ROW As Long = 2
    Const WORK_SUB_ROW As Long = 3

    ' Year used in New headers.
    ' Change this once when the reporting year changes.
    Const MAP_YEAR As Long = 2026


    ' ============================================================
    ' MONTHLY UPDATE CONTROL
    '
    ' TRUE  = update that type
    ' FALSE = do not update that type
    '
    ' Example below:
    '
    ' Feb-Jul = Purchase + Shipment + GRL
    ' Aug     = Purchase only
    '
    ' To change what gets updated next month, edit ONLY this
    ' section.
    ' ============================================================

    Dim mapDef() As Variant
    Dim mapCount As Long

    mapCount = 0

    ' ---------------- FEBRUARY ----------------
    AddMonthMap mapDef, mapCount, "Feb", MAP_YEAR, True, True, True

    ' ---------------- MARCH -------------------
    AddMonthMap mapDef, mapCount, "Mar", MAP_YEAR, True, True, True

    ' ---------------- APRIL -------------------
    AddMonthMap mapDef, mapCount, "Apr", MAP_YEAR, True, True, True

    ' ---------------- MAY ---------------------
    AddMonthMap mapDef, mapCount, "May", MAP_YEAR, True, True, True

    ' ---------------- JUNE --------------------
    AddMonthMap mapDef, mapCount, "Jun", MAP_YEAR, True, True, True

    ' ---------------- JULY --------------------
    AddMonthMap mapDef, mapCount, "Jul", MAP_YEAR, True, True, True

    ' ---------------- AUGUST ------------------
    ' Purchase only
    AddMonthMap mapDef, mapCount, "Aug", MAP_YEAR, True, False, False


    ' ============================================================
    ' EXAMPLE:
    '
    ' If next month you want:
    '
    ' Mar-Jul = P/S/G
    ' Aug     = P only
    ' Sep     = P/S/G
    '
    ' simply change the section above to:
    '
    ' AddMonthMap mapDef, mapCount, "Mar", MAP_YEAR, True, True, True
    ' AddMonthMap mapDef, mapCount, "Apr", MAP_YEAR, True, True, True
    ' AddMonthMap mapDef, mapCount, "May", MAP_YEAR, True, True, True
    ' AddMonthMap mapDef, mapCount, "Jun", MAP_YEAR, True, True, True
    ' AddMonthMap mapDef, mapCount, "Jul", MAP_YEAR, True, True, True
    ' AddMonthMap mapDef, mapCount, "Aug", MAP_YEAR, True, False, False
    ' AddMonthMap mapDef, mapCount, "Sep", MAP_YEAR, True, True, True
    '
    ' ============================================================


    ' ============================================================
    ' COLUMN MAP
    '
    ' colMap(i, 0) = column on New
    ' colMap(i, 1) = column on Workings
    ' ============================================================

    Dim colMap() As Long

    ReDim colMap(1 To mapCount, 0 To 1)


    On Error GoTo CleanFail


    ' ============================================================
    ' WORKSHEETS
    ' ============================================================

    Set wsNew = ThisWorkbook.Sheets("New")
    Set wsWork = ThisWorkbook.Sheets("Workings")


    ' ============================================================
    ' PERFORMANCE SETTINGS
    ' ============================================================

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual


    ' ============================================================
    ' UNMERGE NEW SHEET IF NECESSARY
    ' ============================================================

    mergeState = wsNew.UsedRange.MergeCells

    If IsNull(mergeState) Or mergeState = True Then
        wsNew.Cells.UnMerge
    End If


    ' ============================================================
    ' FIND ITEM MODEL COLUMN ON NEW
    ' ============================================================

    colItemModelNew = FindColByHeader1Row( _
                        wsNew, _
                        NEW_HEADER_ROW, _
                        "ITEM MODEL")


    ' ============================================================
    ' FIND ITEM CODE COLUMN ON WORKINGS
    ' ============================================================

    colItemCodeWork = FindColByHeader2Row( _
                        wsWork, _
                        WORK_MONTH_ROW, _
                        WORK_SUB_ROW, _
                        "", _
                        "ITEM CODE")


    If colItemCodeWork = 0 Then

        colItemCodeWork = FindColByHeader1Row( _
                            wsWork, _
                            WORK_MONTH_ROW, _
                            "ITEM CODE")

    End If


    ' ============================================================
    ' VALIDATE REQUIRED COLUMNS
    ' ============================================================

    If colItemModelNew = 0 Or colItemCodeWork = 0 Then

        MsgBox _
            "Could not find 'ITEM MODEL' on New or 'ITEM CODE' on Workings." & _
            vbCrLf & vbCrLf & _
            "Check that the header text has not changed.", _
            vbCritical

        GoTo CleanExit

    End If


    ' ============================================================
    ' RESOLVE THE SELECTED MONTHLY COLUMNS
    '
    ' For each selected mapping:
    '
    ' New:
    '     "Feb 2026 Qty Purchase"
    '
    ' Workings:
    '     Row 2 = Feb
    '     Row 3 = P
    '
    ' ============================================================

    Dim mapIndex As Long

    warnings = ""


    For mapIndex = LBound(mapDef) To UBound(mapDef)

        ' --------------------------------------------------------
        ' mapDef(index)(0) = New header
        ' mapDef(index)(1) = Workings month
        ' mapDef(index)(2) = Workings P/S/G
        ' mapDef(index)(3) = occurrence
        ' --------------------------------------------------------

        newCol = FindColByHeader1Row( _
                    wsNew, _
                    NEW_HEADER_ROW, _
                    CStr(mapDef(mapIndex)(0)))


        workCol = FindColByHeader2Row( _
                    wsWork, _
                    WORK_MONTH_ROW, _
                    WORK_SUB_ROW, _
                    CStr(mapDef(mapIndex)(1)), _
                    CStr(mapDef(mapIndex)(2)), _
                    CLng(mapDef(mapIndex)(3)))


        colMap(mapIndex, 0) = newCol
        colMap(mapIndex, 1) = workCol


        ' --------------------------------------------------------
        ' Record missing columns.
        ' --------------------------------------------------------

        If newCol = 0 Or workCol = 0 Then

            warnings = warnings & _
                "- " & CStr(mapDef(mapIndex)(0)) & _
                " (New col=" & newCol & _
                ", Workings col=" & workCol & ")" & _
                vbCrLf

        End If

    Next mapIndex


    ' ============================================================
    ' WARN ABOUT MISSING COLUMNS
    ' ============================================================

    If warnings <> "" Then

        If MsgBox( _
            "Some selected columns were not found:" & _
            vbCrLf & vbCrLf & _
            warnings & vbCrLf & _
            "Those columns will be skipped." & _
            vbCrLf & vbCrLf & _
            "Continue anyway?", _
            vbYesNo + vbExclamation) = vbNo Then

            GoTo CleanExit

        End If

    End If


    ' ============================================================
    ' FIND LAST ROWS
    ' ============================================================

    lastRowNew = wsNew.Cells( _
                    wsNew.Rows.Count, _
                    colItemModelNew).End(xlUp).Row


    lastRowWork = wsWork.Cells( _
                    wsWork.Rows.Count, _
                    colItemCodeWork).End(xlUp).Row


    ' ============================================================
    ' BUILD ITEM CODE LOOKUP DICTIONARY
    '
    ' New:
    '
    ' 12345 - Bosch Oven
    '
    ' becomes:
    '
    ' 12345 -> New row
    '
    ' This means we don't repeatedly scan New for every
    ' Workings row.
    ' ============================================================

    Set codeDict = CreateObject("Scripting.Dictionary")


    If lastRowNew >= 2 Then

        modelArr = wsNew.Range( _
                    wsNew.Cells(2, colItemModelNew), _
                    wsNew.Cells(lastRowNew, colItemModelNew)).Value


        For i = 1 To UBound(modelArr, 1)

            itemModelNew = Trim(CStr(modelArr(i, 1)))

            dashPos = InStr(itemModelNew, "-")


            If dashPos > 1 Then

                extractedCode = Trim( _
                                    Left(itemModelNew, dashPos - 1))

            Else

                extractedCode = itemModelNew

            End If


            If extractedCode <> "" Then

                ' First matching code wins.
                If Not codeDict.Exists(extractedCode) Then

                    ' modelArr starts at New row 2,
                    ' so array position 1 = worksheet row 2.
                    codeDict.Add extractedCode, i + 1

                End If

            End If

        Next i

    End If


    ' ============================================================
    ' MATCH WORKINGS ITEM CODES TO NEW
    ' ============================================================

    For workRow = 3 To lastRowWork

        itemCodeWork = Trim( _
                        CStr(wsWork.Cells( _
                            workRow, _
                            colItemCodeWork).Value))


        If itemCodeWork <> "" Then

            If IsNumeric(itemCodeWork) Then

                If codeDict.Exists(itemCodeWork) Then

                    ' Find corresponding row in New.
                    newRow = codeDict(itemCodeWork)


                    ' ====================================================
                    ' COPY SELECTED MONTHLY COLUMNS
                    ' ====================================================

                    For i = LBound(colMap, 1) To UBound(colMap, 1)

                        newCol = colMap(i, 0)
                        workCol = colMap(i, 1)


                        ' Only copy if both columns exist.
                        If newCol > 0 And workCol > 0 Then

                            Set srcCell = wsNew.Cells( _
                                            newRow, _
                                            newCol)


                            Set dstCell = wsWork.Cells( _
                                            workRow, _
                                            workCol)


                            ' --------------------------------------------
                            ' COPY VALUE
                            ' --------------------------------------------

                            dstCell.Value = srcCell.Value


                            ' --------------------------------------------
                            ' COPY FILL COLOUR
                            ' --------------------------------------------

                            dstCell.Interior.Color = _
                                srcCell.Interior.Color

                            dstCell.Interior.Pattern = _
                                srcCell.Interior.Pattern

                        End If

                    Next i

                End If

            End If

        End If

    Next workRow


CleanExit:

    ' ============================================================
    ' RESTORE EXCEL SETTINGS
    ' ============================================================

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True


    If Err.Number = 0 Then

        MsgBox _
            "Done! Data pulled from New to Workings.", _
            vbInformation

    End If

    Exit Sub


CleanFail:

    ' ============================================================
    ' ERROR HANDLER
    ' ============================================================

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True


    MsgBox _
        "Error " & Err.Number & ": " & Err.Description, _
        vbExclamation

End Sub


' =====================================================================
' ADD MONTH MAP
'
' This creates the individual mappings automatically.
'
' Example:
'
' AddMonthMap mapDef, mapCount, "Aug", 2026, True, False, False
'
' produces:
'
' Aug 2026 Qty Purchase -> Aug + P
'
' It does NOT create Shipment or GRL because those are False.
'
' =====================================================================

Private Sub AddMonthMap( _
    ByRef mapDef() As Variant, _
    ByRef mapCount As Long, _
    ByVal monthText As String, _
    ByVal yearNumber As Long, _
    ByVal includePurchase As Boolean, _
    ByVal includeShipment As Boolean, _
    ByVal includeGRL As Boolean)


    ' ------------------------------------------------------------
    ' PURCHASE
    ' ------------------------------------------------------------

    If includePurchase Then

        mapCount = mapCount + 1

        ReDim Preserve mapDef(1 To mapCount)

        mapDef(mapCount) = Array( _
            monthText & " " & yearNumber & " Qty Purchase", _
            monthText, _
            "P", _
            1)

    End If


    ' ------------------------------------------------------------
    ' SHIPMENT
    ' ------------------------------------------------------------

    If includeShipment Then

        mapCount = mapCount + 1

        ReDim Preserve mapDef(1 To mapCount)

        mapDef(mapCount) = Array( _
            monthText & " " & yearNumber & " Qty Shipment", _
            monthText, _
            "S", _
            1)

    End If


    ' ------------------------------------------------------------
    ' GRL
    ' ------------------------------------------------------------

    If includeGRL Then

        mapCount = mapCount + 1

        ReDim Preserve mapDef(1 To mapCount)

        mapDef(mapCount) = Array( _
            monthText & " " & yearNumber & " GRL", _
            monthText, _
            "G", _
            1)

    End If

End Sub


' =====================================================================
' FIND COLUMN BY HEADER IN ONE ROW
'
' Searches one row for targetText.
'
' Returns:
'     Column number if found
'     0 if not found
'
' =====================================================================

Function FindColByHeader1Row( _
    ws As Worksheet, _
    headerRow As Long, _
    targetText As String, _
    Optional occurrence As Long = 1) As Long

    Dim lastCol As Long
    Dim c As Long
    Dim found As Long


    lastCol = ws.Cells( _
                headerRow, _
                ws.Columns.Count).End(xlToLeft).Column


    found = 0


    For c = 1 To lastCol

        If UCase(Trim(CStr( _
            ws.Cells(headerRow, c).Value))) = _
           UCase(targetText) Then

            found = found + 1


            If found = occurrence Then

                FindColByHeader1Row = c
                Exit Function

            End If

        End If

    Next c


    FindColByHeader1Row = 0

End Function


' =====================================================================
' FIND COLUMN BY TWO STACKED HEADERS
'
' Example:
'
'             P       Q       R
' Row 2       Feb     Feb     Feb
' Row 3       P       S       G
'
' Searching for:
'
'     Feb + S
'
' returns column Q.
'
' Returns:
'     Column number if found
'     0 if not found
'
' =====================================================================

Function FindColByHeader2Row( _
    ws As Worksheet, _
    row1 As Long, _
    row2 As Long, _
    text1 As String, _
    text2 As String, _
    Optional occurrence As Long = 1) As Long

    Dim lastCol As Long
    Dim lastCol2 As Long

    Dim c As Long
    Dim found As Long

    Dim v1 As String
    Dim v2 As String


    lastCol = ws.Cells( _
                row1, _
                ws.Columns.Count).End(xlToLeft).Column


    lastCol2 = ws.Cells( _
                row2, _
                ws.Columns.Count).End(xlToLeft).Column


    If lastCol2 > lastCol Then
        lastCol = lastCol2
    End If


    found = 0


    For c = 1 To lastCol

        v1 = Trim(CStr( _
                    ws.Cells(row1, c).Value))

        v2 = Trim(CStr( _
                    ws.Cells(row2, c).Value))


        If _
            (text1 = "" Or UCase(v1) = UCase(text1)) And _
            (text2 = "" Or UCase(v2) = UCase(text2)) Then


            found = found + 1


            If found = occurrence Then

                FindColByHeader2Row = c
                Exit Function

            End If

        End If

    Next c


    FindColByHeader2Row = 0

End Function
