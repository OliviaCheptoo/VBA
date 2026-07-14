Sub UpdateMetrics()

    Dim wsSource As Worksheet
    Dim wsWork As Worksheet

    Dim dataDict As Object
    Dim arr As Variant
    Dim totals As Variant

    Dim lastSource As Long
    Dim lastWork As Long

    Dim i As Long
    Dim workRow As Long

    Dim itemCode As String
    Dim workCode As String

'==================================================================
'                    CHANGE THESE SETTINGS ONLY
'==================================================================

    Const SOURCE_SHEET As String = "Sales"

    Const COL_CODE As String = "R"
    Const COL_QTY As String = "T"
    Const COL_REV As String = "AA"
    Const COL_GP As String = "AC"

    Const DEST_QTY As String = "BU"
    Const DEST_REV As String = "BV"
    Const DEST_GP As String = "BW"

    Const CONVERT_CODE_TO_NUMBER As Boolean = False

'==================================================================

    On Error GoTo CleanFail

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set wsSource = ThisWorkbook.Sheets(SOURCE_SHEET)
    Set wsWork = ThisWorkbook.Sheets("Workings")

    Set dataDict = CreateObject("Scripting.Dictionary")

    lastSource = wsSource.Cells(wsSource.Rows.Count, ColLetterToNum(COL_CODE)).End(xlUp).Row

    arr = wsSource.Range( _
            wsSource.Cells(2, ColLetterToNum(COL_CODE)), _
            wsSource.Cells(lastSource, ColLetterToNum(COL_GP))).Value

    '==================================================
    ' Build Dictionary
    '==================================================

    For i = 1 To UBound(arr, 1)

        If CONVERT_CODE_TO_NUMBER Then
            itemCode = Trim(CStr(Val(arr(i, 1))))
        Else
            itemCode = Trim(CStr(arr(i, 1)))
        End If

        If itemCode <> "" Then

            If dataDict.Exists(itemCode) Then

                totals = dataDict(itemCode)

                If IsNumeric(arr(i, ColLetterToNum(COL_QTY) - ColLetterToNum(COL_CODE) + 1)) Then
                    totals(0) = totals(0) + CDbl(arr(i, ColLetterToNum(COL_QTY) - ColLetterToNum(COL_CODE) + 1))
                End If

                If IsNumeric(arr(i, ColLetterToNum(COL_REV) - ColLetterToNum(COL_CODE) + 1)) Then
                    totals(1) = totals(1) + CDbl(arr(i, ColLetterToNum(COL_REV) - ColLetterToNum(COL_CODE) + 1))
                End If

                If IsNumeric(arr(i, ColLetterToNum(COL_GP) - ColLetterToNum(COL_CODE) + 1)) Then
                    totals(2) = totals(2) + CDbl(arr(i, ColLetterToNum(COL_GP) - ColLetterToNum(COL_CODE) + 1))
                End If

                dataDict(itemCode) = totals

            Else

                dataDict.Add itemCode, Array( _
                    Val(arr(i, ColLetterToNum(COL_QTY) - ColLetterToNum(COL_CODE) + 1)), _
                    Val(arr(i, ColLetterToNum(COL_REV) - ColLetterToNum(COL_CODE) + 1)), _
                    Val(arr(i, ColLetterToNum(COL_GP) - ColLetterToNum(COL_CODE) + 1)))

            End If

        End If

    Next i

    '==================================================
    ' Update Workings
    '==================================================

    lastWork = wsWork.Cells(wsWork.Rows.Count, 4).End(xlUp).Row

    For workRow = 3 To lastWork

        workCode = Trim(CStr(wsWork.Cells(workRow, 4).Value))

        If workCode <> "" Then

            If dataDict.Exists(workCode) Then

                totals = dataDict(workCode)

                wsWork.Cells(workRow, ColLetterToNum(DEST_QTY)).Value = totals(0)
                wsWork.Cells(workRow, ColLetterToNum(DEST_REV)).Value = totals(1)
                wsWork.Cells(workRow, ColLetterToNum(DEST_GP)).Value = totals(2)

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
        MsgBox SOURCE_SHEET & " updated successfully.", vbInformation
    End If

End Sub

Function ColLetterToNum(ByVal colLetter As String) As Long

    Dim i As Long
    Dim c As Long

    ColLetterToNum = 0

    For i = 1 To Len(colLetter)
        c = Asc(UCase(Mid(colLetter, i, 1))) - 64
        ColLetterToNum = ColLetterToNum * 26 + c
    Next i

End Function

