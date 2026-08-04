
Option Explicit

Sub UpdateSalesMetrics()

    Dim wsSource As Worksheet
    Dim wsWork As Worksheet
    
    Dim dataArr As Variant
    Dim dataDicts() As Object
    Dim configurations As Variant
    Dim totals As Variant
    
    Dim lastSource As Long
    Dim lastWork As Long
    
    Dim i As Long
    Dim j As Long
    Dim workRow As Long
    
    Dim itemCode As String
    Dim workCode As String
    Dim channel As String
    
    Dim saleDate As Variant
    
    Dim dateIndex As Long
    Dim channelIndex As Long
    Dim codeIndex As Long
    Dim qtyIndex As Long
    Dim revIndex As Long
    Dim gpIndex As Long
    
    Dim destQtyCol As Long
    Dim destRevCol As Long
    Dim destGPCol As Long
    
    Dim actualReportMonth As Long
    
    
'==================================================================
'                    CHANGE THESE SETTINGS ONLY
'==================================================================

    Const SOURCE_SHEET As String = "Sales"
    Const WORK_SHEET As String = "Workings"

    Const REPORT_YEAR As Long = 2026

    ' 0 = automatically use current month
    ' 1 = January
    ' 2 = February
    ' ...
    ' 12 = December
    Const REPORT_MONTH As Long = 8

    Const CONVERT_CODE_TO_NUMBER As Boolean = False

'==================================================================


    On Error GoTo CleanFail


    '==================================================================
    ' PERFORMANCE SETTINGS
    '==================================================================

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual


    Set wsSource = ThisWorkbook.Sheets(SOURCE_SHEET)
    Set wsWork = ThisWorkbook.Sheets(WORK_SHEET)


    '==================================================================
    ' CONFIGURATION
    '
    ' Array:
    '
    ' Period
    ' Channel
    ' Destination Qty
    ' Destination Revenue
    ' Destination GP
    '
    ' Blank channel = ALL CHANNELS
    ' MT = Modern Trade
    '==================================================================

    configurations = Array( _
        Array("YTD", "", "BZ", "CA", "CB"), _
        Array("YTD", "RETAIL", "CN", "CO", "CP"), _
        Array("YTD", "B2B", "DB", "DC", "DD"), _
        Array("YTD", "ONLINE", "DP", "DQ", "DR"), _
        Array("YTD", "EXPORT", "ED", "EE", "EF"), _
        Array("YTD", "MODERN TRADE", "ER", "ES", "ET"), _
        Array("MTD", "", "FG", "FH", "FI"), _
        Array("MTD", "RETAIL", "FV", "FW", "FX"), _
        Array("MTD", "B2B", "GJ", "GK", "GL"), _
        Array("MTD", "ONLINE", "GX", "GY", "GZ"), _
        Array("MTD", "EXPORT", "HL", "HM", "HN"), _
        Array("MTD", "MODERN TRADE", "HZ", "IA", "IB") _
    )


    '==================================================================
    ' DETERMINE REPORT MONTH
    '==================================================================

    If REPORT_MONTH = 0 Then
        actualReportMonth = Month(Date)
    Else
        actualReportMonth = REPORT_MONTH
    End If


    '==================================================================
    ' FIND LAST SOURCE ROW
    ' Code is in column R
    '==================================================================

    lastSource = wsSource.Cells( _
                    wsSource.Rows.Count, _
                    ColLetterToNum("R")).End(xlUp).Row


    '==================================================================
    ' LOAD SALES DATA INTO MEMORY ONCE
    '
    ' E:AC contains everything required:
    '
    ' E  = Date
    ' H  = Channel
    ' R  = Item Code
    ' T  = Qty
    ' AA = Revenue
    ' AC = GP
    '==================================================================

    dataArr = wsSource.Range( _
                wsSource.Cells(2, ColLetterToNum("E")), _
                wsSource.Cells(lastSource, ColLetterToNum("AC")) _
              ).value


    '==================================================================
    ' ARRAY INDEXES
    '
    ' Array starts at E:
    '
    ' E  = 1
    ' H  = 4
    ' R  = 14
    ' T  = 16
    ' AA = 23
    ' AC = 25
    '==================================================================

    dateIndex = ColLetterToNum("E") - ColLetterToNum("E") + 1
    channelIndex = ColLetterToNum("H") - ColLetterToNum("E") + 1
    codeIndex = ColLetterToNum("R") - ColLetterToNum("E") + 1
    qtyIndex = ColLetterToNum("T") - ColLetterToNum("E") + 1
    revIndex = ColLetterToNum("AA") - ColLetterToNum("E") + 1
    gpIndex = ColLetterToNum("AC") - ColLetterToNum("E") + 1


    '==================================================================
    ' CREATE ONE DICTIONARY FOR EACH CONFIGURATION
    '==================================================================

    ReDim dataDicts(LBound(configurations) To UBound(configurations))


    '==================================================================
    ' BUILD ALL 12 DICTIONARIES
    '==================================================================

    For j = LBound(configurations) To UBound(configurations)

        Set dataDicts(j) = CreateObject("Scripting.Dictionary")


        '--------------------------------------------------------------
        ' LOOP THROUGH SALES ARRAY
        '--------------------------------------------------------------

        For i = 1 To UBound(dataArr, 1)


            '==========================================================
            ' DATE
            '==========================================================

            saleDate = dataArr(i, dateIndex)

            If Not IsDate(saleDate) Then
                GoTo NextSalesRow
            End If


            '==========================================================
            ' YEAR FILTER
            '==========================================================

            If Year(CDate(saleDate)) <> REPORT_YEAR Then
                GoTo NextSalesRow
            End If


            '==========================================================
            ' MTD FILTER
            '
            ' YTD = entire report year
            '
            ' MTD = selected report month
            '==========================================================

            If configurations(j)(0) = "MTD" Then

                If Month(CDate(saleDate)) <> actualReportMonth Then
                    GoTo NextSalesRow
                End If

            End If


            '==========================================================
            ' CHANNEL
            '==========================================================

            channel = UCase(Trim(CStr(dataArr(i, channelIndex))))


            '==========================================================
            ' CHANNEL FILTER
            '
            ' Blank = ALL CHANNELS
            '==========================================================

            If configurations(j)(1) <> "" Then

                If channel <> UCase(configurations(j)(1)) Then
                    GoTo NextSalesRow
                End If

            End If


            '==========================================================
            ' ITEM CODE
            '==========================================================

            If CONVERT_CODE_TO_NUMBER Then

                itemCode = Trim(CStr(Val(dataArr(i, codeIndex))))

            Else

                itemCode = Trim(CStr(dataArr(i, codeIndex)))

            End If


            If itemCode <> "" Then


                '======================================================
                ' CODE ALREADY EXISTS
                '======================================================

                If dataDicts(j).Exists(itemCode) Then

                    totals = dataDicts(j)(itemCode)


                    '--------------------------------------------------
                    ' QTY
                    '--------------------------------------------------

                    If IsNumeric(dataArr(i, qtyIndex)) Then
                        totals(0) = totals(0) + CDbl(dataArr(i, qtyIndex))
                    End If


                    '--------------------------------------------------
                    ' REVENUE
                    '--------------------------------------------------

                    If IsNumeric(dataArr(i, revIndex)) Then
                        totals(1) = totals(1) + CDbl(dataArr(i, revIndex))
                    End If


                    '--------------------------------------------------
                    ' GP
                    '--------------------------------------------------

                    If IsNumeric(dataArr(i, gpIndex)) Then
                        totals(2) = totals(2) + CDbl(dataArr(i, gpIndex))
                    End If


                    dataDicts(j)(itemCode) = totals


                '======================================================
                ' NEW CODE
                '======================================================

                Else

                    dataDicts(j).Add itemCode, Array( _
                        SafeNumber(dataArr(i, qtyIndex)), _
                        SafeNumber(dataArr(i, revIndex)), _
                        SafeNumber(dataArr(i, gpIndex)) _
                    )

                End If

            End If


NextSalesRow:

        Next i

    Next j


    '==================================================================
    ' FIND LAST WORKINGS ROW
    '
    ' Workings item codes are in column D
    '==================================================================

    lastWork = wsWork.Cells( _
                    wsWork.Rows.Count, _
                    4).End(xlUp).Row


    '==================================================================
    ' UPDATE WORKINGS
    '
    ' IMPORTANT:
    ' Data starts from ROW 5
    '==================================================================

    For workRow = 5 To lastWork

        workCode = Trim(CStr(wsWork.Cells(workRow, 4).value))


        If workCode <> "" Then


            '----------------------------------------------------------
            ' PROCESS ALL 12 CONFIGURATIONS
            '----------------------------------------------------------

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

                    ' No sales found for this item/filter
                    wsWork.Cells(workRow, destQtyCol).value = 0
                    wsWork.Cells(workRow, destRevCol).value = 0
                    wsWork.Cells(workRow, destGPCol).value = 0

                End If

            Next j

        End If

    Next workRow


    '==================================================================
    ' RESTORE EXCEL
    '==================================================================

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True


    '==================================================================
    ' COMPLETION MESSAGE
    '==================================================================

    MsgBox "Sales metrics updated successfully." & vbCrLf & vbCrLf & _
           "Report Year: " & REPORT_YEAR & vbCrLf & _
           "MTD Month: " & Format(DateSerial(REPORT_YEAR, actualReportMonth, 1), "mmmm"), _
           vbInformation, "Sales Update Complete"

    Exit Sub


'======================================================================
' ERROR HANDLER
'======================================================================

CleanFail:

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Error " & Err.Number & ": " & Err.Description, _
           vbExclamation, "Sales Update Failed"

End Sub



'======================================================================
' SAFE NUMBER
'======================================================================

Function SafeNumber(ByVal value As Variant) As Double

    If IsError(value) Then

        SafeNumber = 0

    ElseIf IsNumeric(value) Then

        SafeNumber = CDbl(value)

    Else

        SafeNumber = 0

    End If

End Function



'======================================================================
' COLUMN LETTER TO NUMBER
'======================================================================

Function ColLetterToNum(ByVal colLetter As String) As Long

    Dim i As Long
    Dim c As Long

    ColLetterToNum = 0

    For i = 1 To Len(colLetter)

        c = Asc(UCase(Mid(colLetter, i, 1))) - 64

        ColLetterToNum = ColLetterToNum * 26 + c

    Next i

End Function



