Option Explicit

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
    
    Dim lookupColNew As Long
    Dim lookupColWork As Long


    ' ================================================================
    '                     MONTHLY SETTINGS
    ' ================================================================
    
    ' --- START ROWS ---
    Const START_ROW_NEW As Long = 2
    Const START_ROW_WORK As Long = 3
    
    
    ' --- LOOKUP KEY COLUMNS ---
    ' New = column containing text such as:
    '       "12345 - Samsung Oven"
    '
    ' Workings = column containing the item code:
    '       "12345"
    
    Const LOOKUP_COL_NEW As String = "E"
    Const LOOKUP_COL_WORK As String = "D"
    
    
    ' --- SALES / STOCK COLUMN MAPPING ---
    '
    ' Left side  = column on "New"
    ' Right side = column on "Workings"
    
    Const COL_SALES_NEW As String = "BV"
    Const COL_SALES_WORK As String = "AV"
    
    Const COL_SALES_TOTAL_NEW As String = "CB"
    Const COL_SALES_TOTAL_WORK As String = "AW"
    
    Const COL_PHYS_STOCK_NEW As String = "CO"
    Const COL_PHYS_STOCK_WORK As String = "BH"
    
    Const COL_SEA_NEW As String = "CP"
    Const COL_SEA_WORK As String = "BI"
    
    Const COL_GRL_NEW As String = "CQ"
    Const COL_GRL_WORK As String = "BJ"
    
    Const COL_UNSHIPPED_NEW As String = "CR"
    Const COL_UNSHIPPED_WORK As String = "BK"
    
    ' ================================================================
    '                    END MONTHLY SETTINGS
    ' ================================================================


    On Error GoTo CleanFail


    ' ----------------------------------------------------------------
    ' Set worksheets
    ' ----------------------------------------------------------------
    
    Set wsNew = ThisWorkbook.Sheets("New")
    Set wsWork = ThisWorkbook.Sheets("Workings")


    ' ----------------------------------------------------------------
    ' Convert configurable column letters to column numbers
    ' ----------------------------------------------------------------
    
    lookupColNew = ColLetterToNum(LOOKUP_COL_NEW)
    lookupColWork = ColLetterToNum(LOOKUP_COL_WORK)


    ' ----------------------------------------------------------------
    ' Build source/destination column mapping
    ' ----------------------------------------------------------------
    
    colMap(0, 0) = ColLetterToNum(COL_SALES_NEW)
    colMap(0, 1) = ColLetterToNum(COL_SALES_WORK)
    
    colMap(1, 0) = ColLetterToNum(COL_SALES_TOTAL_NEW)
    colMap(1, 1) = ColLetterToNum(COL_SALES_TOTAL_WORK)
    
    colMap(2, 0) = ColLetterToNum(COL_PHYS_STOCK_NEW)
    colMap(2, 1) = ColLetterToNum(COL_PHYS_STOCK_WORK)
    
    colMap(3, 0) = ColLetterToNum(COL_SEA_NEW)
    colMap(3, 1) = ColLetterToNum(COL_SEA_WORK)
    
    colMap(4, 0) = ColLetterToNum(COL_GRL_NEW)
    colMap(4, 1) = ColLetterToNum(COL_GRL_WORK)
    
    colMap(5, 0) = ColLetterToNum(COL_UNSHIPPED_NEW)
    colMap(5, 1) = ColLetterToNum(COL_UNSHIPPED_WORK)


    ' ----------------------------------------------------------------
    ' Improve performance
    ' ----------------------------------------------------------------
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual


    ' ----------------------------------------------------------------
    ' Unmerge cells if necessary
    ' ----------------------------------------------------------------
    
    mergeState = wsNew.UsedRange.MergeCells
    
    If IsNull(mergeState) Or mergeState = True Then
        wsNew.Cells.UnMerge
    End If


    ' ----------------------------------------------------------------
    ' Find last rows
    ' ----------------------------------------------------------------
    
    lastRowNew = wsNew.Cells(wsNew.Rows.Count, lookupColNew).End(xlUp).Row
    lastRowWork = wsWork.Cells(wsWork.Rows.Count, lookupColWork).End(xlUp).Row


    ' ----------------------------------------------------------------
    ' Build lookup dictionary
    '
    ' Key:
    '   Extracted item code from New
    '
    ' Value:
    '   Actual Excel row containing that item
    ' ----------------------------------------------------------------
    
    Set codeDict = CreateObject("Scripting.Dictionary")
    
    codeDict.CompareMode = vbTextCompare


    If lastRowNew >= START_ROW_NEW Then
        
        modelArr = wsNew.Range( _
            wsNew.Cells(START_ROW_NEW, lookupColNew), _
            wsNew.Cells(lastRowNew, lookupColNew) _
        ).Value
        
        
        For i = 1 To UBound(modelArr, 1)
            
            itemModelNew = Trim(CStr(modelArr(i, 1)))
            
            ' Find the first dash
            dashPos = InStr(itemModelNew, "-")
            
            
            ' Extract item code
            If dashPos > 1 Then
                extractedCode = Trim(Left(itemModelNew, dashPos - 1))
            Else
                extractedCode = itemModelNew
            End If
            
            
            ' Skip blank lookup cells
            If extractedCode <> "" Then
                
                ' First matching item wins
                If Not codeDict.Exists(extractedCode) Then
                    
                    codeDict.Add _
                        extractedCode, _
                        i + START_ROW_NEW - 1
                    
                End If
                
            End If
            
        Next i
        
    End If


    ' ----------------------------------------------------------------
    ' Loop through Workings and update matching items
    ' ----------------------------------------------------------------
    
    For workRow = START_ROW_WORK To lastRowWork
        
        ' Get item code from Workings
        itemCodeWork = Trim(CStr(wsWork.Cells(workRow, lookupColWork).Value))
        
        
        ' Skip blank item codes
        If itemCodeWork <> "" Then
            
            
            ' Look for matching item in New
            If codeDict.Exists(itemCodeWork) Then
                
                newRow = codeDict(itemCodeWork)
                
                
                ' ----------------------------------------------------
                ' Copy each mapped Sales / Stock value
                ' ----------------------------------------------------
                
                For i = 0 To UBound(colMap, 1)
                    
                    Set srcCell = wsNew.Cells( _
                        newRow, _
                        colMap(i, 0) _
                    )
                    
                    Set dstCell = wsWork.Cells( _
                        workRow, _
                        colMap(i, 1) _
                    )
                    
                    
                    ' Copy value
                    dstCell.Value = srcCell.Value
                    
                    
                    ' Copy fill colour / pattern
                    dstCell.Interior.Color = srcCell.Interior.Color
                    dstCell.Interior.Pattern = srcCell.Interior.Pattern
                    
                Next i
                
            End If
            
        End If
        
    Next workRow


CleanExit:

    ' ----------------------------------------------------------------
    ' Restore Excel settings
    ' ----------------------------------------------------------------
    
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    
    ' ----------------------------------------------------------------
    ' Completion message
    ' ----------------------------------------------------------------
    
    If Err.Number <> 0 Then
        
        MsgBox _
            "Error " & Err.Number & ": " & Err.Description, _
            vbExclamation
        
    Else
        
        MsgBox _
            "Done! Sept Sales and Stocks updated in Workings.", _
            vbInformation
        
    End If
    
    Exit Sub


CleanFail:

    Resume CleanExit

End Sub


' ====================================================================
' Converts a column letter into its column number
'
' Example:
'   "A"  -> 1
'   "E"  -> 5
'   "BV" -> 74
'   "CR" -> 96
' ====================================================================

Function ColLetterToNum(ByVal colLetter As String) As Long

    Dim i As Long
    Dim c As Long
    
    ColLetterToNum = 0
    
    
    For i = 1 To Len(colLetter)
        
        c = Asc(UCase(Mid(colLetter, i, 1))) - 64
        
        ColLetterToNum = ColLetterToNum * 26 + c
        
    Next i

End Function

