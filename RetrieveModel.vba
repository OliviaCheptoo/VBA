Option Explicit

Function ExtractModelNumber(ByVal desc As String) As String
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim result As String
    
    ' Create RegExp object
    Set regex = CreateObject("VBScript.RegExp")
    ' Updated pattern to match model numbers (alphanumeric with dashes)
    regex.Pattern = "\b[A-Z]{1,}[A-Z0-9-]*[0-9]+[A-Z0-9-]*\b" ' Matches model numbers
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
    Dim modelNumbers As String
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("description")
    
    ' Find the last row in column A
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Loop through each cell in column A
    For i = 2 To lastRow ' Assuming row 1 is headers
        modelNumbers = ExtractModelNumber(ws.Cells(i, 1).Value) ' Extract model numbers
        ws.Cells(i, 2).Value = modelNumbers ' Write results to column B
    Next i
    
    MsgBox "Model numbers extracted successfully!", vbInformation
End Sub

