
Option Explicit

Sub CleanDataSheet()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long
    
    Dim separatorPos As Long
    Dim cellValue As String
    Dim codeValue As String
    Dim modelValue As String
    Dim statusValue As String
    
    Dim headerValue As String
    Dim headerLower As String
    Dim monthValue As String
    
    Dim section As String
    Dim totalsCount As Long
    
    On Error GoTo CleanFail
    
    Set ws = ActiveSheet
    
    '====================================================
    ' PERFORMANCE SETTINGS
    '====================================================
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    
    '====================================================
    ' 1. UNMERGE ALL CELLS
    '====================================================
    ws.UsedRange.UnMerge
    
    
    '====================================================
    ' 2. DELETE ROWS WHERE COLUMN F IS BLANK OR "TOTALS"
    '====================================================
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    
    For i = lastRow To 2 Step -1
        
        cellValue = Trim(CStr(ws.Cells(i, "F").Value))
        
        If cellValue = "" Or LCase(cellValue) = "totals" Then
            ws.Rows(i).Delete
        End If
        
    Next i
    
    
    '====================================================
    ' 3. INSERT TWO NEW COLUMNS G:H
    '
    ' Existing G:H are shifted right.
    ' Nothing gets overwritten.
    '====================================================
    ws.Columns("G:H").Insert Shift:=xlToRight
    
    ws.Range("G1").Value = "Code"
    ws.Range("H1").Value = "Model Number"
    
    
    '====================================================
    ' 4. EXTRACT CODE AND MODEL FROM COLUMN E
    '
    ' Example:
    '
    ' 12345 -- ABC123
    '
    ' G = 12345
    ' H = ABC123
    '====================================================
    
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    
    For i = 2 To lastRow
        
        cellValue = Trim(CStr(ws.Cells(i, "E").Value))
        
        separatorPos = InStr(1, cellValue, "--")
        
        If separatorPos > 0 Then
            
            ' Code = everything before --
            codeValue = Trim(Left(cellValue, separatorPos - 1))
            
            ' Model = everything after --
            modelValue = Trim(Mid(cellValue, separatorPos + 2))
            
            ' Write code as a whole number
            If IsNumeric(codeValue) Then
                
                ws.Cells(i, "G").Value = CDbl(codeValue)
                ws.Cells(i, "G").NumberFormat = "0"
                
            Else
                
                ws.Cells(i, "G").Value = codeValue
                
            End If
            
            ' Write model
            ws.Cells(i, "H").Value = modelValue
            
        End If
        
    Next i
    
    
    '====================================================
    ' 5. HIGHLIGHT EOL MODEL NUMBERS IN RED
    '
    ' Column D = Status
    '
    ' End of Line - EOL
    ' Active Sales Item
    '====================================================
    
    For i = 2 To lastRow
        
        statusValue = Trim(CStr(ws.Cells(i, "D").Value))
        
        If LCase(statusValue) = "end of line - eol" Then
            
            ws.Cells(i, "H").Font.Color = RGB(255, 0, 0)
            
        End If
        
    Next i
    
    
    '====================================================
    ' 6. DELETE ORIGINAL COLUMN E
    '====================================================
    ws.Columns("E").Delete
    
    
    '====================================================
    ' 7. DELETE STATUS COLUMN D
    '====================================================
    ws.Columns("D").Delete
    
    
    '====================================================
    ' 8. SORT BY CUSTOM BRAND ORDER
    '
    ' 1. Von
    ' 2. Hisense
    ' 3. Bosch
    ' 4. SMEG
    ' 5. Simfer
    ' 6. Everything else
    '====================================================
    
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
    
    
    '====================================================
    ' 9. INSERT NEW ROW ABOVE ROW 1
    '====================================================
    ws.Rows(1).Insert Shift:=xlDown
    
    
    '====================================================
    ' 10. RESTRUCTURE HEADERS
    '
    ' STARTING FROM COLUMN G
    '
    ' SECTION 1:
    ' Purchases / Shipment / GRL
    '
    '       Row 1     Row 2
    '       Feb       P
    '       Feb       S
    '       Feb       G
    '
    ' FIRST TOTALS
    '       ?
    '
    ' SECTION 2:
    ' Sales
    '
    '       Row 1     Row 2
    '       Sales     Feb
    '       Sales     Mar
    '       Sales     Apr
    '
    ' SECOND TOTALS
    '       ?
    '
    ' SECTION 3:
    ' Stock
    '
    '       Row 1     Row 2
    '       Stock     Feb
    '       Stock     Mar
    '       Stock     Apr
    '
    ' STOP AT PHYSICAL STOCK
    '====================================================
    
    lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    
    section = "PURCHASE"
    totalsCount = 0
    
    
    For i = 7 To lastCol
        
        headerValue = Trim(CStr(ws.Cells(2, i).Value))
        headerLower = LCase(headerValue)
        
        
        '-----------------------------------------------
        ' STOP AT PHYSICAL STOCK
        '-----------------------------------------------
        If InStr(1, headerLower, "physical stock", vbTextCompare) > 0 Then
            Exit For
        End If
        
        
        '-----------------------------------------------
        ' DETECT TOTALS
        '-----------------------------------------------
        If InStr(1, headerLower, "totals", vbTextCompare) > 0 Then
            
            totalsCount = totalsCount + 1
            
            ' Leave the Totals column unchanged
            
            If totalsCount = 1 Then
                
                section = "SALES"
                
            ElseIf totalsCount = 2 Then
                
                section = "STOCK"
                
            End If
            
            GoTo NextColumn
            
        End If
        
        
        '================================================
        ' PURCHASE / SHIPMENT / GRL
        '================================================
        
        If section = "PURCHASE" Then
            
            monthValue = GetMonthName(headerValue)
            
            If monthValue <> "" Then
                
                ws.Cells(1, i).Value = monthValue
                
            End If
            
            
            If InStr(1, headerLower, "purchase", vbTextCompare) > 0 Then
                
                ws.Cells(2, i).Value = "P"
                
            ElseIf InStr(1, headerLower, "shipment", vbTextCompare) > 0 Then
                
                ws.Cells(2, i).Value = "S"
                
            ElseIf InStr(1, headerLower, "grl", vbTextCompare) > 0 Then
                
                ws.Cells(2, i).Value = "G"
                
            End If
            
            
        '================================================
        ' SALES
        '================================================
        
        ElseIf section = "SALES" Then
            
            monthValue = GetMonthName(headerValue)
            
            If monthValue <> "" Then
                
                ws.Cells(1, i).Value = "Sales"
                ws.Cells(2, i).Value = monthValue
                
            End If
            
            
        '================================================
        ' STOCK
        '================================================
        
        ElseIf section = "STOCK" Then
            
            monthValue = GetMonthName(headerValue)
            
            If monthValue <> "" Then
                
                ws.Cells(1, i).Value = "Stock"
                ws.Cells(2, i).Value = monthValue
                
            End If
            
        End If
        
        
NextColumn:
        
    Next i
    
    
    '====================================================
    ' 11. SET ENTIRE SHEET TO DUBAI 8
    '====================================================
    
    With ws.Cells.Font
        .Name = "Dubai"
        .Size = 8
    End With
    
    
    '====================================================
    ' 12. RESTORE EXCEL SETTINGS
    '====================================================
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    
    '====================================================
    ' 13. COMPLETION MESSAGE
    '====================================================
    
    MsgBox "Data cleaning, sorting and header restructuring is complete.", _
           vbInformation, "Completed"
    
    Exit Sub


'========================================================
' ERROR HANDLER
'========================================================

CleanFail:

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "The macro stopped because of an error:" & vbCrLf & vbCrLf & _
           Err.Description, _
           vbExclamation, "Cleaning Failed"

End Sub



'========================================================
' GET MONTH NAME
'
' Handles both:
'
' Jan 2026
' January 2026
' Feb 2026
' February 2026
' etc.
'
' Returns:
'
' Jan
' Feb
' Mar
' etc.
'========================================================

Function GetMonthName(ByVal headerText As String) As String

    Dim firstWord As String
    Dim firstThree As String
    
    headerText = Trim(headerText)
    
    If Len(headerText) = 0 Then
        GetMonthName = ""
        Exit Function
    End If
    
    
    ' Get first word
    firstWord = Split(headerText, " ")(0)
    
    ' First three characters
    firstThree = LCase(Left(firstWord, 3))
    
    
    Select Case firstThree
        
        Case "jan"
            GetMonthName = "Jan"
            
        Case "feb"
            GetMonthName = "Feb"
            
        Case "mar"
            GetMonthName = "Mar"
            
        Case "apr"
            GetMonthName = "Apr"
            
        Case "may"
            GetMonthName = "May"
            
        Case "jun"
            GetMonthName = "Jun"
            
        Case "jul"
            GetMonthName = "Jul"
            
        Case "aug"
            GetMonthName = "Aug"
            
        Case "sep"
            GetMonthName = "Sep"
            
        Case "oct"
            GetMonthName = "Oct"
            
        Case "nov"
            GetMonthName = "Nov"
            
        Case "dec"
            GetMonthName = "Dec"
            
        Case Else
            GetMonthName = ""
            
    End Select

End Function



