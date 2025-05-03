Sub ReturnClosestOrMatchingHalPrice()
    Dim wsHiHome As Worksheet
    Dim wsHal As Worksheet
    Dim hiHomeLastRow As Long
    Dim halLastRow As Long
    Dim i As Long, j As Long
    Dim modelNumber As String
    Dim hiHomePrice As Double
    Dim halPrice As Variant
    Dim cellValue As Variant
    Dim exactMatchFound As Boolean
    Dim modelFound As Boolean
    Dim closestPrice As Double
    Dim smallestDiff As Double
    
    ' Set the worksheets
    Set wsHiHome = ThisWorkbook.Sheets("HiHome")
    Set wsHal = ThisWorkbook.Sheets("Hal")
    
    ' Get the last row in both sheets
    hiHomeLastRow = wsHiHome.Cells(wsHiHome.Rows.Count, "J").End(xlUp).Row
    halLastRow = wsHal.Cells(wsHal.Rows.Count, "P").End(xlUp).Row
    
    ' Loop through each row in HiHome
    For i = 2 To hiHomeLastRow
        modelNumber = Trim(wsHiHome.Cells(i, 10).Value) ' Column J
        cellValue = wsHiHome.Cells(i, 11).Value          ' Column K
        wsHiHome.Cells(i, 12).Value = ""                 ' Clear column L
        
        exactMatchFound = False
        modelFound = False
        smallestDiff = 9999999 ' Arbitrary large number
        closestPrice = 0
        
        If modelNumber = "" Then GoTo NextRow
        
        ' Ensure HiHome price is numeric
        If IsNumeric(cellValue) Then
            hiHomePrice = CDbl(cellValue)
        Else
            wsHiHome.Cells(i, 12).Value = "Invalid Price"
            GoTo NextRow
        End If
        
        ' Search HAL for matching model
        For j = 2 To halLastRow
            If InStr(1, Trim(wsHal.Cells(j, 16).Value), modelNumber, vbTextCompare) > 0 Then ' Column P
                modelFound = True
                halPrice = wsHal.Cells(j, 20).Value ' Column T
                
                If IsNumeric(halPrice) Then
                    If halPrice = hiHomePrice Then
                        wsHiHome.Cells(i, 12).Value = hiHomePrice
                        exactMatchFound = True
                        Exit For
                    Else
                        ' Update closest price if this is the closest so far
                        If Abs(halPrice - hiHomePrice) < smallestDiff Then
                            smallestDiff = Abs(halPrice - hiHomePrice)
                            closestPrice = halPrice
                        End If
                    End If
                End If
            End If
        Next j
        
        ' If no exact match but model was found
        If Not exactMatchFound And modelFound Then
            wsHiHome.Cells(i, 12).Value = closestPrice
        End If
        
        ' If model not found at all
        If Not modelFound Then
            wsHiHome.Cells(i, 12).Value = "No Match"
        End If
        
NextRow:
    Next i

    MsgBox "Price verification complete. Check column L for results.", vbInformation
End Sub
