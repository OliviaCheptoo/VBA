Sub VerifyPricesBesco()
    Dim wsBesco As Worksheet
    Dim wsHal As Worksheet
    Dim bescoLastRow As Long
    Dim halLastRow As Long
    Dim i As Long, j As Long
    Dim itemCode As String
    Dim bescoPrice As Double
    Dim found As Boolean
    
    ' Set the worksheets
    Set wsBesco = ThisWorkbook.Sheets("Besco")
    Set wsHal = ThisWorkbook.Sheets("Hal")
    
    ' Get the last row in both sheets
    bescoLastRow = wsBesco.Cells(wsBesco.Rows.Count, "A").End(xlUp).Row
    halLastRow = wsHal.Cells(wsHal.Rows.Count, "A").End(xlUp).Row
    
    ' Loop through each row in Besco
    For i = 2 To bescoLastRow
        itemCode = Trim(wsBesco.Cells(i, 1).Value) ' Item code in column A
        bescoPrice = wsBesco.Cells(i, 2).Value ' Price in column B
        found = False
        
        ' Loop through each row in Hal to find the item code
        For j = 2 To halLastRow
            ' Check if the item code matches
            If Trim(wsHal.Cells(j, 1).Value) = itemCode Then
                ' Check if the price matches (with a small tolerance for floating-point precision)
                If Abs(wsHal.Cells(j, 2).Value - bescoPrice) < 0.01 Then
                    wsBesco.Cells(i, 3).Value = "Match Found" ' Output in column C
                    found = True
                    Exit For
                End If
            End If
        Next j
        
        ' If the price was not found, return "Not Found"
        If Not found Then
            wsBesco.Cells(i, 3).Value = "Not Found"
        End If
    Next i
    
    MsgBox "Price verification complete! Check column C for results.", vbInformation
End Sub
