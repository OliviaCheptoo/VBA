Option Explicit

'==================================================================
' MASTER ENTRY POINT
' Runs all three stages in the required order:
'   1. Clean & restructure the raw data sheet
'   2. Insert totals rows + SUM formulas by brand/category/subcat
'   3. Insert brand-separator rows + borders
'
' Application settings and error handling are managed ONCE here,
' instead of three times (which caused redundant screen flicker
' and three separate MsgBox popups in the original scripts).
'==================================================================
Sub RunFullDataPipeline()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    On Error GoTo PipelineFail

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Step1_CleanDataSheet ws
    Step2_InsertTotalsAndFillFormulas ws
    Step3_FormatBrandSeparationAndBorders ws

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "Pipeline complete: data cleaned, totals inserted, and formatting applied.", _
           vbInformation, "Done"

    Exit Sub

PipelineFail:

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "The pipeline stopped because of an error:" & vbCrLf & vbCrLf & _
           Err.Description, vbExclamation, "Pipeline Failed"

End Sub


'==================================================================
' STAGE 1 — formerly "CleanDataSheet"
' Unmerges cells, removes blank/Totals rows, splits Code/Model out
' of column E, drops Status (D) and the original combined column
' (E), sorts by custom brand order, and restructures the two
' header rows from column G onward.
'==================================================================
Private Sub Step1_CleanDataSheet(ws As Worksheet)

    Dim lastRow As Long, lastCol As Long, i As Long
    Dim separatorPos As Long
    Dim cellValue As String, codeValue As String, modelValue As String, statusValue As String
    Dim headerValue As String, headerLower As String, monthValue As String
    Dim section As String, totalsCount As Long

    ' 1. Unmerge all cells
    ws.UsedRange.UnMerge

    ' 2. Delete rows where column F is blank or "Totals"
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    For i = lastRow To 2 Step -1
        cellValue = Trim(CStr(ws.Cells(i, "F").Value))
        If cellValue = "" Or LCase(cellValue) = "totals" Then
            ws.Rows(i).Delete
        End If
    Next i

    ' 3. Insert two new columns G:H
    ws.Columns("G:H").Insert Shift:=xlToRight
    ws.Range("G1").Value = "Code"
    ws.Range("H1").Value = "Model Number"

    ' 4. Extract Code and Model from column E ("12345 -- ABC123")
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    For i = 2 To lastRow
        cellValue = Trim(CStr(ws.Cells(i, "E").Value))
        separatorPos = InStr(1, cellValue, "--")
        If separatorPos > 0 Then
            codeValue = Trim(Left(cellValue, separatorPos - 1))
            modelValue = Trim(Mid(cellValue, separatorPos + 2))

            If IsNumeric(codeValue) Then
                ws.Cells(i, "G").Value = CDbl(codeValue)
                ws.Cells(i, "G").NumberFormat = "0"
            Else
                ws.Cells(i, "G").Value = codeValue
            End If

            ws.Cells(i, "H").Value = modelValue
        End If
    Next i

    ' 5. Highlight EOL model numbers in red (column D = Status)
    For i = 2 To lastRow
        statusValue = Trim(CStr(ws.Cells(i, "D").Value))
        If LCase(statusValue) = "end of line - eol" Then
            ws.Cells(i, "H").Font.Color = RGB(255, 0, 0)
        End If
    Next i

    ' 6. Delete original column E
    ws.Columns("E").Delete

    ' 7. Delete Status column D
    ws.Columns("D").Delete

    ' 8. Sort by custom brand order
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add2 _
            Key:=ws.Range("A2:A" & lastRow), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            CustomOrder:="Von,Hisense,Bosch,SMEG,Simfer", _
            DataOption:=xlSortNormal
        .SetRange ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

    ' 9. Insert new row above row 1
    ws.Rows(1).Insert Shift:=xlDown

    ' 10. Restructure headers from column G onward:
    '     Purchase/Shipment/GRL -> Totals -> Sales -> Totals -> Stock -> stop at Physical Stock
    lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column

    section = "PURCHASE"
    totalsCount = 0

    For i = 7 To lastCol

        headerValue = Trim(CStr(ws.Cells(2, i).Value))
        headerLower = LCase(headerValue)

        If InStr(1, headerLower, "physical stock", vbTextCompare) > 0 Then
            Exit For
        End If

        If InStr(1, headerLower, "totals", vbTextCompare) > 0 Then
            totalsCount = totalsCount + 1
            If totalsCount = 1 Then
                section = "SALES"
            ElseIf totalsCount = 2 Then
                section = "STOCK"
            End If
            GoTo NextColumn
        End If

        If section = "PURCHASE" Then

            monthValue = GetMonthName(headerValue)
            If monthValue <> "" Then ws.Cells(1, i).Value = monthValue

            If InStr(1, headerLower, "purchase", vbTextCompare) > 0 Then
                ws.Cells(2, i).Value = "P"
            ElseIf InStr(1, headerLower, "shipment", vbTextCompare) > 0 Then
                ws.Cells(2, i).Value = "S"
            ElseIf InStr(1, headerLower, "grl", vbTextCompare) > 0 Then
                ws.Cells(2, i).Value = "G"
            End If

        ElseIf section = "SALES" Then

            monthValue = GetMonthName(headerValue)
            If monthValue <> "" Then
                ws.Cells(1, i).Value = "Sales"
                ws.Cells(2, i).Value = monthValue
            End If

        ElseIf section = "STOCK" Then

            monthValue = GetMonthName(headerValue)
            If monthValue <> "" Then
                ws.Cells(1, i).Value = "Stock"
                ws.Cells(2, i).Value = monthValue
            End If

        End If

NextColumn:
    Next i

    ' 11. Set entire sheet to Dubai 8
    With ws.Cells.Font
        .Name = "Dubai"
        .Size = 8
    End With

End Sub


'==================================================================
' STAGE 2 — formerly "InsertTotalsAndFillFormulas_BlueAccent5_Fixed_Bold"
' Inserts a totals row wherever Brand (A) / Category (B) / Subcat (C)
' changes, tags it in column F, colors it by level (Accent5 tint
' 0.4/0.6/0.8), bolds it, then fills SUM formulas from column G
' based on that tint coding.
'==================================================================
Private Sub Step2_InsertTotalsAndFillFormulas(ws As Worksheet)

    Dim lastRow As Long, lastCol As Long, i As Long, idx As Long, j As Long
    Dim insertPoints As Collection
    Dim brand As String, category As String, subcat As String
    Dim targetRange As Range
    Dim formulaCol As Long
    Dim cell As Range
    Dim endRow As Long
    Dim sumRange As String

    formulaCol = 7 ' Column G for totals

    lastRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    lastCol = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    Set insertPoints = New Collection

    For i = 2 To lastRow
        brand = CStr(ws.Cells(i, 1).Value)
        category = CStr(ws.Cells(i, 2).Value)
        subcat = CStr(ws.Cells(i, 3).Value)

        If brand <> "" Then
            If ws.Cells(i - 1, 1).Value <> brand Then insertPoints.Add Array(i, brand, "brand")
        End If
        If category <> "" Then
            If ws.Cells(i - 1, 2).Value <> category Then insertPoints.Add Array(i, category, "category")
        End If
        If subcat <> "" Then
            If ws.Cells(i - 1, 3).Value <> subcat Then insertPoints.Add Array(i, subcat, "subcat")
        End If
    Next i

    For idx = insertPoints.Count To 1 Step -1
        i = insertPoints(idx)(0)
        ws.Rows(i).Insert Shift:=xlDown
        ws.Cells(i, 6).Value = insertPoints(idx)(1)
        ws.Rows(i).Font.Bold = True

        Set targetRange = ws.Range(ws.Cells(i, 1), ws.Cells(i, lastCol))
        Select Case insertPoints(idx)(2)
            Case "subcat"
                targetRange.Interior.ThemeColor = xlThemeColorAccent5
                targetRange.Interior.TintAndShade = 0.8
            Case "category"
                targetRange.Interior.ThemeColor = xlThemeColorAccent5
                targetRange.Interior.TintAndShade = 0.6
            Case "brand"
                targetRange.Interior.ThemeColor = xlThemeColorAccent5
                targetRange.Interior.TintAndShade = 0.4
        End Select
    Next idx

    lastRow = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row

    For i = 2 To lastRow
        Set cell = ws.Cells(i, 6)

        If cell.Interior.ThemeColor = xlThemeColorAccent5 And Abs(cell.Interior.TintAndShade - 0.8) < 0.01 Then
            endRow = i + 1
            Do While endRow <= lastRow And ws.Cells(endRow, 3).Value <> ""
                endRow = endRow + 1
            Loop
            ws.Cells(i, formulaCol).Formula = "=SUM(G" & i + 1 & ":G" & endRow - 1 & ")"

        ElseIf cell.Interior.ThemeColor = xlThemeColorAccent5 And Abs(cell.Interior.TintAndShade - 0.6) < 0.01 Then
            sumRange = ""
            For j = i + 1 To lastRow
                If ws.Cells(j, 6).Interior.ThemeColor = xlThemeColorAccent5 And Abs(ws.Cells(j, 6).Interior.TintAndShade - 0.6) < 0.01 Then Exit For
                If ws.Cells(j, 6).Interior.ThemeColor = xlThemeColorAccent5 And Abs(ws.Cells(j, 6).Interior.TintAndShade - 0.8) < 0.01 Then
                    If sumRange = "" Then sumRange = "G" & j Else sumRange = sumRange & ",G" & j
                End If
            Next j
            If sumRange <> "" Then ws.Cells(i, formulaCol).Formula = "=SUM(" & sumRange & ")"

        ElseIf cell.Interior.ThemeColor = xlThemeColorAccent5 And Abs(cell.Interior.TintAndShade - 0.4) < 0.01 Then
            sumRange = ""
            For j = i + 1 To lastRow
                If ws.Cells(j, 6).Interior.ThemeColor = xlThemeColorAccent5 And Abs(ws.Cells(j, 6).Interior.TintAndShade - 0.4) < 0.01 Then Exit For
                If ws.Cells(j, 6).Interior.ThemeColor = xlThemeColorAccent5 And Abs(ws.Cells(j, 6).Interior.TintAndShade - 0.6) < 0.01 Then
                    If sumRange = "" Then sumRange = "G" & j Else sumRange = sumRange & ",G" & j
                End If
            Next j
            If sumRange <> "" Then ws.Cells(i, formulaCol).Formula = "=SUM(" & sumRange & ")"
        End If

        If ws.Cells(i, formulaCol).HasFormula Then
            ws.Range(ws.Cells(i, formulaCol), ws.Cells(i, lastCol)).FillRight
        End If
    Next i

End Sub


'==================================================================
' STAGE 3 — formerly "FormatBrandSeparationAndBorders"
' Uses the tint-0.4 (brand-level) coloring from Stage 2 to find
' brand block boundaries, inserts blank separator rows between
' brands, then draws box borders around each data row.
'==================================================================
Private Sub Step3_FormatBrandSeparationAndBorders(ws As Worksheet)

    Dim lastRow As Long, lastCol As Long, i As Long, b As Integer
    Dim brandCount As Integer
    Dim brandInfo() As Long
    Dim cell As Range, rowRange As Range
    Dim borderColor As Long
    Dim isBlankRow As Boolean
    Dim brandRowHeight As Double

    borderColor = RGB(70, 130, 180)

    lastRow = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row
    lastCol = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    ' Step 1: identify brand block boundaries
    ReDim brandInfo(1 To 200, 1 To 2)
    brandCount = 0
    For i = 2 To lastRow
        Set cell = ws.Cells(i, 6)
        If cell.Interior.ThemeColor = xlThemeColorAccent5 And _
           Abs(cell.Interior.TintAndShade - 0.4) < 0.01 Then
            If brandCount > 0 Then brandInfo(brandCount, 2) = i - 1
            brandCount = brandCount + 1
            brandInfo(brandCount, 1) = i
        End If
    Next i
    If brandCount > 0 Then brandInfo(brandCount, 2) = lastRow

    ' Step 2: insert blank separator rows between brands
    brandRowHeight = ws.Rows(brandInfo(1, 1)).RowHeight
    For b = brandCount To 2 Step -1
        ws.Rows(brandInfo(b, 1)).Insert Shift:=xlDown
        With ws.Rows(brandInfo(b, 1))
            .Interior.ColorIndex = xlNone
            .Borders.LineStyle = xlNone
            .Font.ColorIndex = xlAutomatic
            .Font.Bold = False
            .RowHeight = brandRowHeight
        End With
    Next b

    ' Step 3: re-scan after row insertion
    lastRow = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row
    brandCount = 0
    ReDim brandInfo(1 To 200, 1 To 2)
    For i = 2 To lastRow
        Set cell = ws.Cells(i, 6)
        If cell.Interior.ThemeColor = xlThemeColorAccent5 And _
           Abs(cell.Interior.TintAndShade - 0.4) < 0.01 Then
            If brandCount > 0 Then brandInfo(brandCount, 2) = i - 1
            brandCount = brandCount + 1
            brandInfo(brandCount, 1) = i
        End If
    Next i
    If brandCount > 0 Then brandInfo(brandCount, 2) = lastRow

    ' Step 4: apply box borders row by row, skipping blank separators
    For b = 1 To brandCount
        For i = brandInfo(b, 1) To brandInfo(b, 2)
            Set cell = ws.Cells(i, 6)
            isBlankRow = (cell.Interior.ColorIndex = xlNone Or _
                          cell.Interior.ColorIndex = 0) And cell.Value = ""
            If isBlankRow Then GoTo NextRow

            Set rowRange = ws.Range(ws.Cells(i, 1), ws.Cells(i, lastCol))
            rowRange.BorderAround xlContinuous, xlThin, xlColorIndexNone, borderColor
            rowRange.Borders(xlInsideVertical).LineStyle = xlContinuous
            rowRange.Borders(xlInsideVertical).Weight = xlThin
            rowRange.Borders(xlInsideVertical).Color = borderColor

            If cell.Interior.ThemeColor = xlThemeColorAccent5 Then
                rowRange.Font.Color = RGB(0, 0, 0)
            End If
NextRow:
        Next i
    Next b

End Sub


'==================================================================
' HELPER — used by Stage 1
' Converts "Jan 2026" / "January 2026" etc. into a 3-letter code.
'==================================================================
Function GetMonthName(ByVal headerText As String) As String

    Dim firstWord As String
    Dim firstThree As String

    headerText = Trim(headerText)
    If Len(headerText) = 0 Then
        GetMonthName = ""
        Exit Function
    End If

    firstWord = Split(headerText, " ")(0)
    firstThree = LCase(Left(firstWord, 3))

    Select Case firstThree
        Case "jan": GetMonthName = "Jan"
        Case "feb": GetMonthName = "Feb"
        Case "mar": GetMonthName = "Mar"
        Case "apr": GetMonthName = "Apr"
        Case "may": GetMonthName = "May"
        Case "jun": GetMonthName = "Jun"
        Case "jul": GetMonthName = "Jul"
        Case "aug": GetMonthName = "Aug"
        Case "sep": GetMonthName = "Sep"
        Case "oct": GetMonthName = "Oct"
        Case "nov": GetMonthName = "Nov"
        Case "dec": GetMonthName = "Dec"
        Case Else: GetMonthName = ""
    End Select

End Function
