Option Explicit

Sub PullFromNew()

    Const NEW_HEADER_ROW As Long = 1
    Const WORK_MONTH_ROW As Long = 2
    Const WORK_SUB_ROW As Long = 3

    Dim wsNew As Worksheet, wsWork As Worksheet
    Dim codeDict As Object
    Dim maps As Collection
    Dim itemArr As Variant
    Dim itemCode As String, model As String
    Dim newRow As Long, workRow As Long
    Dim newCol As Long, workCol As Long
    Dim itemModelCol As Long, itemCodeCol As Long
    Dim dashPos As Long, i As Long
    Dim m As Variant
    Dim src As Range, dst As Range
    Dim warnings As String

    '============================================================
    ' WHAT TO UPDATE
    '
    ' PSG = Purchase + Shipment + GRL
    ' P   = Purchase only
    ' PS  = Purchase + Shipment
    ' G   = GRL only
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

    Set wsNew = ThisWorkbook.Sheets("New")
    Set wsWork = ThisWorkbook.Sheets("Workings")

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    '============================================================
    ' FIND KEY COLUMNS
    '============================================================

    itemModelCol = FindHeader(wsNew, NEW_HEADER_ROW, "ITEM MODEL")

    itemCodeCol = FindHeader(wsWork, WORK_MONTH_ROW, "ITEM CODE")

    If itemModelCol = 0 Or itemCodeCol = 0 Then
        Err.Raise vbObjectError + 1, , _
            "Could not find ITEM MODEL or ITEM CODE."
    End If

    '============================================================
    ' RESOLVE SELECTED MONTH/CATEGORY COLUMNS
    '============================================================

    Dim colMap As Collection
    Set colMap = New Collection

    For Each m In maps

        ' m(1) = month
        ' m(2) = PSG selection

        If InStr(m(2), "P") > 0 Then
            AddColumnMap colMap, wsNew, wsWork, _
                m(1), "P", "Qty Purchase", _
                NEW_HEADER_ROW, WORK_MONTH_ROW, WORK_SUB_ROW
        End If

        If InStr(m(2), "S") > 0 Then
            AddColumnMap colMap, wsNew, wsWork, _
                m(1), "S", "Qty Shipment", _
                NEW_HEADER_ROW, WORK_MONTH_ROW, WORK_SUB_ROW
        End If

        If InStr(m(2), "G") > 0 Then
            AddColumnMap colMap, wsNew, wsWork, _
                m(1), "G", "GRL", _
                NEW_HEADER_ROW, WORK_MONTH_ROW, WORK_SUB_ROW
        End If

    Next m

    '============================================================
    ' BUILD ITEM CODE LOOKUP
    '============================================================

    Dim lastNew As Long, lastWork As Long

    lastNew = wsNew.Cells(wsNew.Rows.Count, itemModelCol).End(xlUp).Row
    lastWork = wsWork.Cells(wsWork.Rows.Count, itemCodeCol).End(xlUp).Row

    Set codeDict = CreateObject("Scripting.Dictionary")

    If lastNew >= 2 Then

        itemArr = wsNew.Range( _
            wsNew.Cells(2, itemModelCol), _
            wsNew.Cells(lastNew, itemModelCol)).Value

        For i = 1 To UBound(itemArr, 1)

            model = Trim(CStr(itemArr(i, 1)))
            dashPos = InStr(model, "-")

            If dashPos > 1 Then
                itemCode = Trim(Left(model, dashPos - 1))
            Else
                itemCode = model
            End If

            If itemCode <> "" Then
                If Not codeDict.Exists(itemCode) Then
                    codeDict.Add itemCode, i + 1
                End If
            End If

        Next i

    End If

    '============================================================
    ' MATCH ITEMS AND COPY SELECTED COLUMNS
    '============================================================

    For workRow = 3 To lastWork

        itemCode = Trim(CStr(wsWork.Cells(workRow, itemCodeCol).Value))

        If itemCode <> "" And codeDict.Exists(itemCode) Then

            newRow = codeDict(itemCode)

            For Each m In colMap

                newCol = m(0)
                workCol = m(1)

                Set src = wsNew.Cells(newRow, newCol)
                Set dst = wsWork.Cells(workRow, workCol)

                dst.Value = src.Value
                dst.Interior.Color = src.Interior.Color
                dst.Interior.Pattern = src.Interior.Pattern

            Next m

        End If

    Next workRow

CleanExit:

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If Err.Number = 0 Then
        MsgBox "Done! Data pulled from New to Workings.", vbInformation
    End If

    Exit Sub

Fail:

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation

End Sub


'============================================================
' ADD MONTH TO CONTROL TABLE
'============================================================

Private Sub AddMap(ByRef maps As Collection, _
                   ByVal month As String, _
                   ByVal types As String)

    maps.Add Array(month, types)

End Sub


'============================================================
' FIND AND STORE A COLUMN PAIR
'============================================================

Private Sub AddColumnMap(ByRef colMap As Collection, _
                         ByVal wsNew As Worksheet, _
                         ByVal wsWork As Worksheet, _
                         ByVal month As String, _
                         ByVal subType As String, _
                         ByVal newSuffix As String, _
                         ByVal newHeaderRow As Long, _
                         ByVal workMonthRow As Long, _
                         ByVal workSubRow As Long)

    Dim newCol As Long
    Dim workCol As Long
    Dim c As Long
    Dim lastCol As Long
    Dim header As String

    ' Find New column.
    lastCol = wsNew.Cells( _
        newHeaderRow, wsNew.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol

        header = Trim(CStr(wsNew.Cells(newHeaderRow, c).Value))

        If LCase(header) Like LCase(month & " * " & newSuffix) Then
            newCol = c
            Exit For
        End If

    Next c

    ' Find Workings column.
    workCol = FindTwoHeaders( _
        wsWork, workMonthRow, workSubRow, _
        month, subType)

    If newCol > 0 And workCol > 0 Then
        colMap.Add Array(newCol, workCol)
    End If

End Sub


'============================================================
' FIND HEADER IN ONE ROW
'============================================================

Private Function FindHeader(ByVal ws As Worksheet, _
                            ByVal rowNum As Long, _
                            ByVal text As String) As Long

    Dim c As Long
    Dim lastCol As Long

    lastCol = ws.Cells( _
        rowNum, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol

        If UCase(Trim(CStr(ws.Cells(rowNum, c).Value))) = _
           UCase(text) Then

            FindHeader = c
            Exit Function

        End If

    Next c

End Function


'============================================================
' FIND TWO-ROW HEADER
'============================================================

Private Function FindTwoHeaders(ByVal ws As Worksheet, _
                                ByVal row1 As Long, _
                                ByVal row2 As Long, _
                                ByVal text1 As String, _
                                ByVal text2 As String) As Long

    Dim c As Long
    Dim lastCol As Long

    lastCol = Application.Max( _
        ws.Cells(row1, ws.Columns.Count).End(xlToLeft).Column, _
        ws.Cells(row2, ws.Columns.Count).End(xlToLeft).Column)

    For c = 1 To lastCol

        If UCase(Trim(CStr(ws.Cells(row1, c).Value))) = UCase(text1) _
           And _
           UCase(Trim(CStr(ws.Cells(row2, c).Value))) = UCase(text2) Then

            FindTwoHeaders = c
            Exit Function

        End If

    Next c

End Function
