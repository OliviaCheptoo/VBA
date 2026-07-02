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
    Dim dashPos As Long
    Dim i As Long
    Dim srcCell As Range
    Dim dstCell As Range
    Dim colMap(0 To 5, 0 To 1) As Long
    Dim modelArr As Variant
    Dim codeDict As Object
    Dim mergeState As Variant

    ' === MONTHLY SETTINGS - update these column letters each month ===
    ' Left side  = column letter on "New"
    ' Right side = column letter on "Workings"
    Const COL_SALES_NEW As String = "BU"        ' Qty Sales (current month)
    Const COL_SALES_WORK As String = "AX"
    Const COL_SALES_TOTAL_NEW As String = "CB"   ' Sales Total
    Const COL_SALES_TOTAL_WORK As String = "AY"
    Const COL_PHYS_STOCK_NEW As String = "CO"    ' Physical Stock
    Const COL_PHYS_STOCK_WORK As String = "BG"
    Const COL_SEA_NEW As String = "CP"           ' Sea
    Const COL_SEA_WORK As String = "BH"
    Const COL_GRL_NEW As String = "CQ"           ' GRL
    Const COL_GRL_WORK As String = "BI"
    Const COL_UNSHIPPED_NEW As String = "CR"     ' Unshipped
    Const COL_UNSHIPPED_WORK As String = "BJ"
    ' ===================================================================

    On Error GoTo CleanFail

    Set wsNew = ThisWorkbook.Sheets("New")
    Set wsWork = ThisWorkbook.Sheets("Workings")

    ' --- COLUMN MAPPING: built from the letters above, (New col, Workings col) ---
    colMap(0, 0) = ColLetterToNum(COL_SALES_NEW):       colMap(0, 1) = ColLetterToNum(COL_SALES_WORK)
    colMap(1, 0) = ColLetterToNum(COL_SALES_TOTAL_NEW): colMap(1, 1) = ColLetterToNum(COL_SALES_TOTAL_WORK)
    colMap(2, 0) = ColLetterToNum(COL_PHYS_STOCK_NEW):  colMap(2, 1) = ColLetterToNum(COL_PHYS_STOCK_WORK)
    colMap(3, 0) = ColLetterToNum(COL_SEA_NEW):         colMap(3, 1) = ColLetterToNum(COL_SEA_WORK)
    colMap(4, 0) = ColLetterToNum(COL_GRL_NEW):         colMap(4, 1) = ColLetterToNum(COL_GRL_WORK)
    colMap(5, 0) = ColLetterToNum(COL_UNSHIPPED_NEW):   colMap(5, 1) = ColLetterToNum(COL_UNSHIPPED_WORK)

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' Only unmerge if the sheet actually has merged cells - skip the
    ' (expensive, whole-sheet) UnMerge call otherwise.
    ' UsedRange.MergeCells returns False if nothing is merged, True if the
    ' whole range is one merge, Null if it's a mix - so Null or True both
    ' mean "there's at least one merged cell somewhere".
    mergeState = wsNew.UsedRange.MergeCells
    If IsNull(mergeState) Or mergeState = True Then
        wsNew.Cells.UnMerge
    End If

    lastRowNew = wsNew.Cells(wsNew.Rows.Count, 1).End(xlUp).Row
    lastRowWork = wsWork.Cells(wsWork.Rows.Count, 4).End(xlUp).Row

    ' --- Build a one-time lookup: extracted item code -> row on "New" ---
    ' Replaces the old inner loop that re-scanned all of "New" for every
    ' row of "Workings" (O(n*m)). This scan is O(n), and the match below
    ' is O(1), so the whole thing runs in O(n+m) instead.
    Set codeDict = CreateObject("Scripting.Dictionary")

    If lastRowNew >= 2 Then
        modelArr = wsNew.Range(wsNew.Cells(2, 5), wsNew.Cells(lastRowNew, 5)).Value
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
        itemCodeWork = Trim(CStr(wsWork.Cells(workRow, 4).Value))
        If itemCodeWork <> "" Then
            If IsNumeric(itemCodeWork) Then
                If codeDict.Exists(itemCodeWork) Then
                    newRow = codeDict(itemCodeWork)
                    For i = 0 To UBound(colMap, 1)
                        Set srcCell = wsNew.Cells(newRow, colMap(i, 0))
                        Set dstCell = wsWork.Cells(workRow, colMap(i, 1))
                        dstCell.Value = srcCell.Value
                        dstCell.Interior.Color = srcCell.Interior.Color
                        dstCell.Interior.Pattern = srcCell.Interior.Pattern
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
        MsgBox "Done! Sales and Stocks updated in Workings.", vbInformation
    End If
End Sub

' Converts a column letter (e.g. "BV") to its column number (e.g. 74).
' Pure math, no dependency on any worksheet.
Function ColLetterToNum(ByVal colLetter As String) As Long
    Dim i As Long
    Dim c As Long
    ColLetterToNum = 0
    For i = 1 To Len(colLetter)
        c = Asc(UCase(Mid(colLetter, i, 1))) - 64
        ColLetterToNum = ColLetterToNum * 26 + c
    Next i
End Function
