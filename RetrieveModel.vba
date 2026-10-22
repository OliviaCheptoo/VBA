Option Explicit

Function ExtractModelNumber(ByVal desc As String) As String
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim result As String
    
    ' Create RegExp object
    Set regex = CreateObject("VBScript.RegExp")
    ' Pattern to match model numbers (alphanumeric with dashes)
    regex.Pattern = "\b[A-Z]{1,}[A-Z0-9-]*[0-9]+[A-Z0-9-]*\b"
    regex.Global = True
    
    ' Find matches
    If regex.test(desc) Then
        Set matches = regex.Execute(desc)
        For Each match In matches
            ' Append all potential model numbers, separated by commas
            If result = "" Then
                result = match.Value
            Else
                result = result & ", " & match.Value
            End If
        Next match
    End If
    
    ' Return extracted model numbers
    ExtractModelNumber = result
End Function

Sub RetrieveModelNumbers()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' Run on whichever sheet is currently active
    Set ws = ActiveSheet
    
    ' Find the last row in column A
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Application.ScreenUpdating = False
    
    ' Insert a new column B (existing column B and beyond shift right)
    ws.Columns(2).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ws.Cells(1, 2).Value = "Model Number"
    
    ' Loop through column A (row 1 assumed to be headers)
    For i = 2 To lastRow
        ws.Cells(i, 2).Value = ExtractModelNumber(CStr(ws.Cells(i, 1).Value))
    Next i
    
    Application.ScreenUpdating = True
    
    MsgBox "Model numbers extracted successfully!", vbInformation
End Sub
