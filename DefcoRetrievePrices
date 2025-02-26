Sub VerifyPrices()
    Dim wsDefco As Worksheet
    Dim wsHal As Worksheet
    Dim defcoLastRow As Long
    Dim halLastRow As Long
    Dim i As Long, j As Long
    Dim modelNumber As String
    Dim defcoPrice As Double
    Dim found As Boolean
    
    ' Set the worksheets
    Set wsDefco = ThisWorkbook.Sheets("Defco Verify Prices")
    Set wsHal = ThisWorkbook.Sheets("Hal")
    
    ' Get the last row in both sheets
    defcoLastRow = wsDefco.Cells(wsDefco.Rows.Count, "B").End(xlUp).Row
    halLastRow = wsHal.Cells(wsHal.Rows.Count, "B").End(xlUp).Row
    
    ' Loop through each row in Defco Verify Prices
    For i = 2 To defcoLastRow
        modelNumber = Trim(wsDefco.Cells(i, 2).Value) ' Trim to remove extra spaces
        defcoPrice = wsDefco.Cells(i, 6).Value
        found = False
        
        ' Loop through each row in Hal to find the model number in the description
        For j = 2 To halLastRow
            ' Check if the model number exists in the description (case-insensitive)
            If InStr(1, Trim(wsHal.Cells(j, 2).Value), modelNumber, vbTextCompare) > 0 Then
                ' Check if the price matches (with a small tolerance for floating-point precision)
                If Abs(wsHal.Cells(j, 3).Value - defcoPrice) < 0.01 Then
                    wsDefco.Cells(i, 7).Value = "Match Found"
                    found = True
                    Exit For
                End If
            End If
        Next j
        
        ' If the price was not found, return "Not Found"
        If Not found Then
            wsDefco.Cells(i, 7).Value = "Not Found"
        End If
    Next i
    
    MsgBox "Price verification complete! Check column G for results.", vbInformation
End Sub
