Sub FormatBrandSeparationAndBorders()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long, i As Long, b As Integer
    Dim brandCount As Integer
    Dim brandInfo() As Long
    Dim cell As Range, rowRange As Range
    Dim borderColor As Long
    Dim isBlankRow As Boolean

    Set ws = ActiveSheet
    borderColor = RGB(70, 130, 180)

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    lastRow = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row
    lastCol = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    ' --- STEP 1: Identify brand block boundaries ---
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

    ' --- STEP 2: Insert blank separator rows between brands ---
    Dim brandRowHeight As Double
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

    ' --- STEP 3: Re-scan after row insertion ---
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

    ' --- STEP 4: Apply box borders row by row, skip blank separators ---
    For b = 1 To brandCount
        For i = brandInfo(b, 1) To brandInfo(b, 2)
            Set cell = ws.Cells(i, 6)

            ' Skip blank separator rows
            isBlankRow = (cell.Interior.ColorIndex = xlNone Or _
                          cell.Interior.ColorIndex = 0) And cell.Value = ""
            If isBlankRow Then GoTo NextRow

            ' Box border on every cell in the row
            Set rowRange = ws.Range(ws.Cells(i, 1), ws.Cells(i, lastCol))
            rowRange.BorderAround xlContinuous, xlThin, xlColorIndexNone, borderColor
            rowRange.Borders(xlInsideVertical).LineStyle = xlContinuous
            rowRange.Borders(xlInsideVertical).Weight = xlThin
            rowRange.Borders(xlInsideVertical).Color = borderColor

            ' Black text on header rows
            If cell.Interior.ThemeColor = xlThemeColorAccent5 Then
                rowRange.Font.Color = RGB(0, 0, 0)
            End If

NextRow:
        Next i
    Next b

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    MsgBox "Done! " & brandCount & " brand blocks formatted.", vbInformation
End Sub

