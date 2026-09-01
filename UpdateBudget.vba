Option Explicit

Sub UpdateMetrics()

    Dim wsSource As Worksheet
    Dim wsWork As Worksheet
    
    Dim dataArr As Variant
    Dim dataDicts() As Object
    
    Dim configurations As Variant
    Dim periodSourceColumns As Variant
    Dim offsets As Variant
    Dim totals As Variant
    
    Dim lastSource As Long
    Dim lastWork As Long
    
    Dim i As Long
    Dim j As Long
    Dim workRow As Long
    
    Dim itemCode As String
    Dim workCode As String
    Dim channel As String
    
    Dim qtyIndex As Long
    Dim revIndex As Long
    Dim gpIndex As Long
    
    Dim codeIndex As Long
    Dim channelIndex As Long
    
    Dim destQtyCol As Long
    Dim destRevCol As Long
    Dim destGPCol As Long
    
    Const SOURCE_SHEET As String = "Budget"
    Const WORK_SHEET As String = "Workings"
    
    Const COL_CODE As String = "E"
    Const COL_CHANNEL As String = "D"
    
    Const CONVERT_CODE_TO_NUMBER As Boolean = False
    
    
    '========================================================
    ' PERIOD SOURCE COLUMNS (Budget sheet)
    '
    ' Array:
    '
    ' Period Name
    ' Qty Column
    ' Rev Column
    ' GP Column
    '
    ' Add a new row here to support a new period (e.g. "LY",
    ' "FCST") without touching any code below. The period
    ' name here must match the period name used in the
    ' "configurations" array further down.
    '========================================================
    
    periodSourceColumns = Array( _
        Array("YTD", "T", "AG", "AT"), _
        Array("MTD", "M", "Z", "AM") _
    )
    
    
    '========================================================
    ' CONFIGURATION
    '
    ' Array:
    '
    ' Period            (must match a name in periodSourceColumns)
    ' Channel
    ' Destination Qty
    ' Destination Rev
    ' Destination GP
    '
    ' Blank channel = ALL CHANNELS
    '========================================================
    
    configurations = Array( _
        Array("YTD", "", "BQ", "BR", "BS"), _
        Array("MTD", "", "CE", "CF", "CG"), _
        Array("YTD", "RETAIL", "CS", "CT", "CU"), _
        Array("MTD", "RETAIL", "DG", "DH", "DI"), _
        Array("YTD", "B2B", "DU", "DV", "DW"), _
        Array("MTD", "B2B", "EI", "EJ", "EK"), _
        Array("YTD", "ONLINE", "EW", "EX", "EY"), _
        Array("MTD", "ONLINE", "FK", "FL", "FM"), _
        Array("YTD", "EXPORT", "FY", "FZ", "GA"), _
        Array("MTD", "EXPORT", "GM", "GN", "GO"), _
        Array("YTD", "MT", "HA", "HB", "HC"), _
        Array("MTD", "MT", "HP", "HQ", "HR") _
    )
    
    
    On Error GoTo CleanFail
    
    
    '========================================================
    ' PERFORMANCE SETTINGS
    '========================================================
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    
    Set wsSource = ThisWorkbook.Sheets(SOURCE_SHEET)
    Set wsWork = ThisWorkbook.Sheets(WORK_SHEET)
    
    
    '========================================================
    ' FIND LAST SOURCE ROW
    '========================================================
    
    lastSource = wsSource.Cells( _
                    wsSource.Rows.Count, _
                    ColLetterToNum(COL_CODE)).End(xlUp).Row
    
    
    '========================================================
    ' LOAD BUDGET INTO MEMORY ONCE
    '
    ' IMPORTANT:
    ' We start at D, NOT E, because Channel is in D.
    '
    ' D:AT
    '========================================================
    
    dataArr = wsSource.Range( _
                wsSource.Cells(2, ColLetterToNum("D")), _
                wsSource.Cells(lastSource, ColLetterToNum("AT")) _
              ).value
    
    
    '========================================================
    ' ARRAY INDEXES
    '
    ' Array starts at D = 1
    '
    ' D  = 1  Channel
    ' E  = 2  Code
    '
    ' Qty/Rev/GP indexes are now resolved per-period via
    ' GetPeriodOffsets() below, instead of being hardcoded.
    '========================================================
    
    channelIndex = ColLetterToNum("D") - ColLetterToNum("D") + 1
    codeIndex = ColLetterToNum("E") - ColLetterToNum("D") + 1
    
    '========================================================
    ' CREATE ONE DICTIONARY FOR EACH CONFIGURATION
    '========================================================
    
    ReDim dataDicts(LBound(configurations) To UBound(configurations))
    
    
    '========================================================
    ' BUILD ALL DICTIONARIES
    '========================================================
    
    For j = LBound(configurations) To UBound(configurations)
        
        Set dataDicts(j) = CreateObject("Scripting.Dictionary")
        
        
        '--------------------------------------------
        ' Determine source columns based on period
        ' (looked up from periodSourceColumns instead
        ' of a hardcoded If/Else)
        '--------------------------------------------
        
        offsets = GetPeriodOffsets(configurations(j)(0), periodSourceColumns)
        
        qtyIndex = offsets(0)
        revIndex = offsets(1)
        gpIndex = offsets(2)
        
        
        '--------------------------------------------
        ' LOOP THROUGH SOURCE ARRAY
        '--------------------------------------------
        
        For i = 1 To UBound(dataArr, 1)
            
            '----------------------------------------
            ' GET CHANNEL FROM ARRAY
            '----------------------------------------
            
            channel = UCase(Trim(CStr(dataArr(i, channelIndex))))
            
            
            '----------------------------------------
            ' CHANNEL FILTER
            '
            ' Blank = All Channels
            '----------------------------------------
            
            If configurations(j)(1) <> "" Then
                
                If channel <> UCase(configurations(j)(1)) Then
                    GoTo NextSourceRow
                End If
                
            End If
            
            
            '----------------------------------------
            ' GET ITEM CODE
            '----------------------------------------
            
            If CONVERT_CODE_TO_NUMBER Then
                
                itemCode = Trim(CStr(Val(dataArr(i, codeIndex))))
                
            Else
                
                itemCode = Trim(CStr(dataArr(i, codeIndex)))
                
            End If
            
            
            If itemCode <> "" Then
                
                '------------------------------------
                ' CODE ALREADY EXISTS
                '------------------------------------
                
                If dataDicts(j).Exists(itemCode) Then
                    
                    totals = dataDicts(j)(itemCode)
                    
                    
                    If IsNumeric(dataArr(i, qtyIndex)) Then
                        totals(0) = totals(0) + CDbl(dataArr(i, qtyIndex))
                    End If
                    
                    
                    If IsNumeric(dataArr(i, revIndex)) Then
                        totals(1) = totals(1) + CDbl(dataArr(i, revIndex))
                    End If
                    
                    
                    If IsNumeric(dataArr(i, gpIndex)) Then
                        totals(2) = totals(2) + CDbl(dataArr(i, gpIndex))
                    End If
                    
                    
                    dataDicts(j)(itemCode) = totals
                    
                    
                '------------------------------------
                ' NEW CODE
                '------------------------------------
                
                Else
                    
                    dataDicts(j).Add itemCode, Array( _
                        SafeNumber(dataArr(i, qtyIndex)), _
                        SafeNumber(dataArr(i, revIndex)), _
                        SafeNumber(dataArr(i, gpIndex)) _
                    )
                    
                End If
                
            End If
            
            
NextSourceRow:
            
        Next i
        
    Next j
    
    
    '========================================================
    ' FIND LAST WORKINGS ROW
    '========================================================
    
    lastWork = wsWork.Cells( _
                    wsWork.Rows.Count, _
                    4).End(xlUp).Row
    
    
    '========================================================
    ' UPDATE WORKINGS
    '========================================================
    
    For workRow = 3 To lastWork
        
        workCode = Trim(CStr(wsWork.Cells(workRow, 4).value))
        
        
        If workCode <> "" Then
            
            
            '--------------------------------------------
            ' PROCESS ALL CONFIGURATIONS
            '--------------------------------------------
            
            For j = LBound(configurations) To UBound(configurations)
                
                
                destQtyCol = ColLetterToNum(configurations(j)(2))
                destRevCol = ColLetterToNum(configurations(j)(3))
                destGPCol = ColLetterToNum(configurations(j)(4))
                
                
                If dataDicts(j).Exists(workCode) Then
                    
                    totals = dataDicts(j)(workCode)
                    
                    wsWork.Cells(workRow, destQtyCol).value = totals(0)
                    wsWork.Cells(workRow, destRevCol).value = totals(1)
                    wsWork.Cells(workRow, destGPCol).value = totals(2)
                    
                Else
                    
                    ' No matching data
                    wsWork.Cells(workRow, destQtyCol).value = 0
                    wsWork.Cells(workRow, destRevCol).value = 0
                    wsWork.Cells(workRow, destGPCol).value = 0
                    
                End If
                
            Next j
            
        End If
        
    Next workRow
    
    
    '========================================================
    ' RESTORE EXCEL
    '========================================================
    
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    
    MsgBox "YTD and MTD metrics updated successfully.", _
           vbInformation, "Update Complete"
    
    Exit Sub


'============================================================
' ERROR HANDLER
'============================================================

CleanFail:

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    MsgBox "Error " & Err.Number & ": " & Err.Description, _
           vbExclamation, "Update Failed"

End Sub



'============================================================
' SAFE NUMBER
'============================================================

Function SafeNumber(ByVal value As Variant) As Double

    If IsError(value) Then
        
        SafeNumber = 0
        
    ElseIf IsNumeric(value) Then
        
        SafeNumber = CDbl(value)
        
    Else
        
        SafeNumber = 0
        
    End If

End Function



'============================================================
' GET PERIOD OFFSETS
'
' Looks up a period name (e.g. "YTD") in periodSourceColumns
' and returns its Qty/Rev/GP column indexes, already converted
' to be relative to column D (i.e. ready to use directly
' against dataArr).
'
' Raises an error if the period name isn't found, so a typo
' in "configurations" fails loudly instead of silently
' returning zeros.
'============================================================

Function GetPeriodOffsets(ByVal periodName As String, ByVal periodCols As Variant) As Variant

    Dim k As Long
    Dim baseCol As Long
    
    baseCol = ColLetterToNum("D")
    
    For k = LBound(periodCols) To UBound(periodCols)
        
        If UCase(Trim(periodCols(k)(0))) = UCase(Trim(periodName)) Then
            
            GetPeriodOffsets = Array( _
                ColLetterToNum(periodCols(k)(1)) - baseCol + 1, _
                ColLetterToNum(periodCols(k)(2)) - baseCol + 1, _
                ColLetterToNum(periodCols(k)(3)) - baseCol + 1 _
            )
            
            Exit Function
            
        End If
        
    Next k
    
    Err.Raise vbObjectError + 1, "GetPeriodOffsets", _
        "Unknown period '" & periodName & "' - add it to periodSourceColumns."

End Function



'============================================================
' COLUMN LETTER TO NUMBER
'============================================================

Function ColLetterToNum(ByVal colLetter As String) As Long

    Dim i As Long
    Dim c As Long
    
    ColLetterToNum = 0
    
    For i = 1 To Len(colLetter)
        
        c = Asc(UCase(Mid(colLetter, i, 1))) - 64
        
        ColLetterToNum = ColLetterToNum * 26 + c
        
    Next i

End Function
