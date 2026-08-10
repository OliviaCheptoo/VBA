Option Explicit

Sub PullFromNew()

    Const NEW_HEADER_ROW As Long = 1
    Const WORK_MONTH_ROW As Long = 2
    Const WORK_SUB_ROW As Long = 3

    Dim wsNew As Worksheet
    Dim wsWork As Worksheet

    Dim maps As Collection
    Dim colMap As Collection

    Dim codeDict As Object

    Dim itemModelCol As Long
    Dim itemCodeCol As Long

    Dim lastNew As Long
    Dim lastWork As Long

    Dim itemArr As Variant

    Dim itemCode As String
    Dim model As String

    Dim newRow As Long
    Dim workRow As Long

    Dim dashPos As Long
    Dim i As Long

    Dim m As Variant
    Dim cm As Variant

    Dim src As Range
    Dim dst As Range

    Dim warnings As String


    '============================================================
    ' MONTHLY UPDATE CONTROL
    '
    ' PSG = Purchase + Shipment + GRL
    ' P   = Purchase only
    ' PS  = Purchase + Shipment
    ' G   = GRL only
    '
    ' Change ONLY this section each month.
    '============================================================

    Set maps = New Collection

    AddMap maps, "Feb", "PSG"
    AddMap maps, "Mar", "PSG"
    AddMap maps, "Apr", "PSG"
    AddMap maps, "May", "PSG"
    AddMap maps, "Jun", "PSG"
    AddMap maps, "Jul", "PSG"
    AddMap maps, "Aug", "P"


    On Error GoTo Fail


    '============================================================
    ' WORKSHEETS
    '============================================================

    Set wsNew = ThisWorkbook.Sheets("New")
    Set wsWork = ThisWorkbook.Sheets("Workings")


    '============================================================
    ' PERFORMANCE SETTINGS
    '============================================================

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual


    '============================================================
    ' UNMERGE NEW SHEET IF NECESSARY
    '============================================================

    Dim mergeState As Variant

    mergeState = wsNew.UsedRange.MergeCells

    If IsNull(mergeState) Or mergeState = True Then
        wsNew.Cells.UnMerge
    End If


    '============================================================
    ' FIND ITEM MODEL ON NEW
    '============================================================

    itemModelCol = FindHeader( _
                        wsNew, _
                        NEW_HEADER_ROW, _
                        "ITEM MODEL")


    '============================================================
    ' FIND ITEM CODE ON WORKINGS
    '
    ' First try:
    '     Row 2 = anything
    '     Row 3 = ITEM CODE
    '
    ' If not found, try ITEM CODE directly in row 2.
    ' This preserves the logic from your original macro.
    '============================================================

    itemCodeCol = FindTwoHeaders( _
                        wsWork, _
                        WORK_MONTH_ROW, _
                        WORK_SUB_ROW, _
                        "", _
                        "ITEM CODE")


    If itemCodeCol = 0 Then

        itemCodeCol = FindHeader( _
                            wsWork, _
                            WORK_MONTH_ROW, _
                            "ITEM CODE")

    End If


    '============================================================
    ' CHECK REQUIRED COLUMNS
    '============================================================

    If itemModelCol = 0 Or itemCodeCol = 0 Then

        Err.Raise vbObjectError + 1, , _
            "Could not find ITEM MODEL or ITEM CODE." & vbCrLf & _
            "ITEM MODEL column = " & itemModelCol & vbCrLf & _
            "ITEM CODE column = " & itemCodeCol

    End If


    '============================================================
    ' BUILD COLUMN MAP
    '
    ' Each entry contains:
    '
    ' cm(0) = New column
    ' cm(1) = Workings column
    '
    '============================================================

    Set colMap = New Collection

    For Each m In maps

        '--------------------------------------------------------
        ' PURCHASE
        '--------------------------------------------------------

        If InStr(1, m(1), "P", vbTextCompare) > 0 Then

            AddColumnMap _
                colMap, _
                wsNew, _
                wsWork, _
                m(0), _
                "P", _
                "Qty Purchase", _
                NEW_HEADER_ROW, _
                WORK_MONTH_ROW, _
                WORK_SUB_ROW, _
                warnings

        End If


        '--------------------------------------------------------
        ' SHIPMENT
        '--------------------------------------------------------

        If InStr(1, m(1), "S", vbTextCompare) > 0 Then

            AddColumnMap _
                colMap, _
                wsNew, _
                wsWork, _
                m(0), _
                "S", _
                "Qty Shipment", _
                NEW_HEADER_ROW, _
                WORK_MONTH_ROW, _
                WORK_SUB_ROW, _
                warnings

        End If


        '--------------------------------------------------------
        ' GRL
        '--------------------------------------------------------

        If InStr(1, m(1), "G", vbTextCompare) > 0 Then

            AddColumnMap _
                colMap, _
                wsNew, _
                wsWork, _
                m(0), _
                "G", _
                "GRL", _
                NEW_HEADER_ROW, _
                WORK_MONTH_ROW, _
                WORK_SUB_ROW, _
                warnings

        End If

    Next m


    '============================================================
    ' REPORT MISSING SELECTED COLUMNS
    '============================================================

    If warnings <> "" Then

        If MsgBox( _
            "Some selected columns were not found:" & _
            vbCrLf & vbCrLf & _
            warnings & _
            vbCrLf & _
            "These columns will be skipped." & _
            vbCrLf & vbCrLf & _
            "Continue?", _
            vbYesNo + vbExclamation) = vbNo Then

            GoTo CleanExit

        End If

    End If


    '============================================================
    ' FIND LAST ROWS
    '============================================================

    lastNew = wsNew.Cells( _
                    wsNew.Rows.Count, _
                    itemModelCol).End(xlUp).Row

    lastWork = wsWork.Cells( _
                    wsWork.Rows.Count, _
                    itemCodeCol).End(xlUp).Row


    '============================================================
    ' BUILD ITEM CODE DICTIONARY
    '
    ' Example:
    '
    ' New ITEM MODEL:
    '
    ' 12345 - Bosch Oven
    '
    ' becomes:
    '
    ' 12345 -> New row
    '
    ' First match wins.
    '============================================================

    Set codeDict = CreateObject("Scripting.Dictionary")


    If lastNew >= 2 Then

        itemArr = wsNew.Range( _
                    wsNew.Cells(2, itemModelCol), _
                    wsNew.Cells(lastNew, itemModelCol)).Value


        For i = 1 To UBound(itemArr, 1)

            model = Trim(CStr(itemArr(i, 1)))

            dashPos = InStr(model, "-")


            If dashPos > 1 Then

                itemCode = Trim( _
                                Left(model, dashPos - 1))

            Else

                itemCode = model

            End If


            If itemCode <> "" Then

                If Not codeDict.Exists(itemCode) Then

                    ' Array starts at New row 2.
                    ' Therefore array position 1 = row 2.
                    codeDict.Add itemCode, i + 1

                End If

            End If

        Next i

    End If


    '============================================================
    ' MATCH WORKINGS ITEMS TO NEW
    '============================================================

    For workRow = 3 To lastWork

        itemCode = Trim( _
                    CStr(wsWork.Cells( _
                        workRow, _
                        itemCodeCol).Value))


        If itemCode <> "" Then

            If codeDict.Exists(itemCode) Then

                ' Get matching row from New.
                newRow = codeDict(itemCode)


                '================================================
                ' COPY SELECTED COLUMNS
                '================================================

                For Each cm In colMap

                    ' cm(0) = New column
                    ' cm(1) = Workings column

                    Set src = wsNew.Cells( _
                                newRow, _
                                cm(0))

                    Set dst = wsWork.Cells( _
                                workRow, _
                                cm(1))


                    '------------------------------------------------
                    ' COPY VALUE
                    '------------------------------------------------

                    dst.Value = src.Value


                    '------------------------------------------------
                    ' COPY FILL ONLY
                    '------------------------------------------------

                    dst.Interior.Color = src.Interior.Color

                    dst.Interior.Pattern = src.Interior.Pattern

                Next cm

            End If

        End If

    Next workRow


CleanExit:

    '============================================================
    ' RESTORE EXCEL
    '============================================================

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True


    If Err.Number = 0 Then

        MsgBox _
            "Done! Data pulled from New to Workings.", _
            vbInformation

    End If

    Exit Sub


Fail:

    '============================================================
    ' ERROR HANDLER
    '============================================================

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True


    MsgBox _
        "Error " & Err.Number & ": " & Err.Description, _
        vbExclamation

End Sub


'====================================================================
' ADD MONTH TO CONTROL TABLE
'
' Example:
'
' AddMap maps, "Aug", "P"
'
' means:
'
' August Purchase only.
'
'====================================================================

Private Sub AddMap( _
    ByRef maps As Collection, _
    ByVal monthName As String, _
    ByVal types As String)

    maps.Add Array(monthName, types)

End Sub


'====================================================================
' FIND AND ADD A NEW/WORKINGS COLUMN PAIR
'
' New example:
'
'     Aug 2026 Qty Purchase
'
' Workings:
'
'     Row 2 = Aug
'     Row 3 = P
'
'====================================================================

Private Sub AddColumnMap( _
    ByRef colMap As Collection, _
    ByVal wsNew As Worksheet, _
    ByVal wsWork As Worksheet, _
    ByVal monthName As String, _
    ByVal subType As String, _
    ByVal newSuffix As String, _
    ByVal newHeaderRow As Long, _
    ByVal workMonthRow As Long, _
    ByVal workSubRow As Long, _
    ByRef warnings As String)

    Dim newCol As Long
    Dim workCol As Long

    Dim c As Long
    Dim lastCol As Long

    Dim header As String


    '============================================================
    ' FIND COLUMN ON NEW
    '
    ' We do NOT specify the year.
    '
    ' So:
    '
    ' Aug 2026 Qty Purchase
    '
    ' and eventually:
    '
    ' Aug 2027 Qty Purchase
    '
    ' can both be found.
    '============================================================

    lastCol = wsNew.Cells( _
                newHeaderRow, _
                wsNew.Columns.Count).End(xlToLeft).Column


    For c = 1 To lastCol

        header = Trim(CStr( _
                    wsNew.Cells(newHeaderRow, c).Value))


        If LCase(header) Like _
            LCase(monthName & " * " & newSuffix) Then

            newCol = c
            Exit For

        End If

    Next c


    '============================================================
    ' FIND COLUMN ON WORKINGS
    '============================================================

    workCol = FindTwoHeaders( _
                wsWork, _
                workMonthRow, _
                workSubRow, _
                monthName, _
                subType)


    '============================================================
    ' STORE IF BOTH EXIST
    '============================================================

    If newCol > 0 And workCol > 0 Then

        colMap.Add Array(newCol, workCol)

    Else

        warnings = warnings & _
            "- " & monthName & _
            " / " & subType & _
            " | New column=" & newCol & _
            " | Workings column=" & workCol & _
            vbCrLf

    End If

End Sub


'====================================================================
' FIND HEADER IN ONE ROW
'
' Returns:
'
'     column number if found
'     0 if not found
'
' Case-insensitive and trims spaces.
'====================================================================

Private Function FindHeader( _
    ByVal ws As Worksheet, _
    ByVal rowNum As Long, _
    ByVal target As String) As Long

    Dim c As Long
    Dim lastCol As Long


    lastCol = ws.Cells( _
                rowNum, _
                ws.Columns.Count).End(xlToLeft).Column


    For c = 1 To lastCol

        If UCase(Trim(CStr( _
            ws.Cells(rowNum, c).Value))) = _
           UCase(target) Then

            FindHeader = c
            Exit Function

        End If

    Next c


    FindHeader = 0

End Function


'====================================================================
' FIND COLUMN USING TWO HEADER ROWS
'
' Example:
'
'          P       Q       R
' Row 2    Aug     Aug     Aug
' Row 3    P       S       G
'
' FindTwoHeaders(..., "Aug", "S")
'
' returns Q.
'
' If text1 = "", row 1 is treated as a wildcard.
'====================================================================

Private Function FindTwoHeaders( _
    ByVal ws As Worksheet, _
    ByVal row1 As Long, _
    ByVal row2 As Long, _
    ByVal text1 As String, _
    ByVal text2 As String) As Long

    Dim c As Long
    Dim lastCol As Long

    Dim v1 As String
    Dim v2 As String


    lastCol = Application.Max( _
                ws.Cells(row1, ws.Columns.Count).End(xlToLeft).Column, _
                ws.Cells(row2, ws.Columns.Count).End(xlToLeft).Column)


    For c = 1 To lastCol

        v1 = Trim(CStr( _
                    ws.Cells(row1, c).Value))

        v2 = Trim(CStr( _
                    ws.Cells(row2, c).Value))


        If _
            (text1 = "" Or UCase(v1) = UCase(text1)) And _
            (text2 = "" Or UCase(v2) = UCase(text2)) Then

            FindTwoHeaders = c
            Exit Function

        End If

    Next c


    FindTwoHeaders = 0

End Function
