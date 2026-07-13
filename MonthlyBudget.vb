Sub UpdateBudget()

    Dim wsNew As Worksheet
    Dim wsWork As Worksheet
    Dim lastRowNew As Long
    Dim lastRowWork As Long
    Dim newRow As Long
    Dim workRow As Long
    Dim itemCodeWork As String
    Dim itemCodeNew As String
    Dim i As Long
    Dim srcCell As Range
    Dim dstCell As Range
    Dim colMap(0 To 2, 0 To 1) As Long
    Dim codeArr As Variant
    Dim codeDict As Object
    Dim mergeState As Variant

'====================================================================
'                    CHANGE THESE SETTINGS ONLY
'====================================================================

    Const SOURCE_SHEET As String = "Sales"     'Budget or Sales

    'Source columns
    Const COL_QTY_NEW As String = "G"
    Const COL_REV_NEW As String = "N"
    Const COL_GP_NEW As String = "P"

    'Destination columns on Workings
    Const COL_QTY_WORK As String = "BU"
    Const COL_REV_WORK As String = "BV"
    Const COL_GP_WORK As String = "BW"

    'Item Code columns
    Const SOURCE_CODE_COL As String = "E"
    Const WORK_CODE_COL As String = "D"

    'TRUE for Sales (because item codes are stored as text)
    'FALSE for Budget
    Const CONVERT_SOURCE_TO_NUMBER As Boolean = False

'====================================================================

    On Error GoTo CleanFail

    Set wsNew = ThisWorkbook.Sheets(SOURCE_SHEET)
    Set wsWork = ThisWorkbook.Sheets("Workings")

    colMap(0, 0) = ColLetterToNum(COL_QTY_NEW)
    colMap(0, 1) = ColLetterToNum(COL_QTY_WORK)

    colMap(1, 0) = ColLetterToNum(COL_REV_NEW)
    colMap(1, 1) = ColLetterToNum(COL_REV_WORK)

    colMap(2, 0) = ColLetterToNum(COL_GP_NEW)
    colMap(2, 1) = ColLetterToNum(COL_GP_WORK)

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    mergeState = wsNew.UsedRange.MergeCells
    If IsNull(mergeState) Or mergeState = True Then
        wsNew.Cells.UnMerge
    End If

    lastRowNew = wsNew.Cells(wsNew.Rows.Count, ColLetterToNum(SOURCE_CODE_COL)).End(xlUp).Row
    lastRowWork = wsWork.Cells(wsWork.Rows.Count, ColLetterToNum(WORK_CODE_COL)).End(xlUp).Row

    Set codeDict = CreateObject("Scripting.Dictionary")

    codeArr = wsNew.Range( _
        wsNew.Cells(2, ColLetterToNum(SOURCE_CODE_COL)), _
        wsNew.Cells(lastRowNew, ColLetterToNum(SOURCE_CODE_COL))).Value

    For i = 1 To UBound(codeArr, 1)

        If CONVERT_SOURCE_TO_NUMBER Then
            itemCodeNew = Trim(CStr(Val(codeArr(i, 1))))
        Else
            itemCodeNew = Trim(CStr(codeArr(i, 1)))
        End If

        If itemCodeNew <> "" Then
            If Not codeDict.Exists(itemCodeNew) Then
                codeDict.Add itemCodeNew, i + 1
            End If
        End If

    Next i

    For workRow = 3 To lastRowWork

        itemCodeWork = Trim(CStr(wsWork.Cells(workRow, ColLetterToNum(WORK_CODE_COL)).Value))

        If codeDict.Exists(itemCodeWork) Then

            newRow = codeDict(itemCodeWork)

            For i = 0 To 2

                Set srcCell = wsNew.Cells(newRow, colMap(i, 0))
                Set dstCell = wsWork.Cells(workRow, colMap(i, 1))

                dstCell.Value = srcCell.Value
                dstCell.Interior.Color = srcCell.Interior.Color
                dstCell.Interior.Pattern = srcCell.Interior.Pattern

            Next i

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

