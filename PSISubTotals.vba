Sub InsertTotalsAndFillFormulas_BlueAccent5_Fixed_Bold()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long, i As Long, idx As Long
    Dim insertPoints As Collection
    Dim brand As String, category As String, subcat As String
    Dim targetRange As Range
    Dim formulaCol As Long
    
    Set ws = ActiveSheet
    formulaCol = 7 ' Column G for totals
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' --- Step 1: Find last row/column ---
    lastRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    lastCol = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    Set insertPoints = New Collection
    
    ' --- Step 2: Record where to insert totals ---
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
    
    ' --- Step 3: Insert rows in reverse order, color them, and make bold ---
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
    
    ' --- Step 4: Fill formulas based on theme color + tint ---
    lastRow = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row
    
    For i = 2 To lastRow
        Dim cell As Range
        Set cell = ws.Cells(i, 6)
        
        If cell.Interior.ThemeColor = xlThemeColorAccent5 And Abs(cell.Interior.TintAndShade - 0.8) < 0.01 Then
            Dim endRow As Long
            endRow = i + 1
            Do While endRow <= lastRow And ws.Cells(endRow, 3).Value <> ""
                endRow = endRow + 1
            Loop
            ws.Cells(i, formulaCol).Formula = "=SUM(G" & i + 1 & ":G" & endRow - 1 & ")"
        
        ElseIf cell.Interior.ThemeColor = xlThemeColorAccent5 And Abs(cell.Interior.TintAndShade - 0.6) < 0.01 Then
            Dim sumRange As String, j As Long
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
    
    ' --- Step 5: Restore ---
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
