Sub FillInvoiceNumbers()
    Dim wsDefco As Worksheet, wsHAL As Worksheet
    Dim lastRowDefco As Long, lastRowHAL As Long
    Dim i As Long, j As Long
    Dim modelNumber As String, defcoPrice As Double, shopNumber As String
    Dim bestInvoice As String
    Dim found As Boolean
    
    ' Set worksheet references
    Set wsDefco = ThisWorkbook.Sheets("DefcoInvoices")
    Set wsHAL = ThisWorkbook.Sheets("HALDefcoSellin")
    
    ' Get last row in each sheet
    lastRowDefco = wsDefco.Cells(wsDefco.Rows.Count, "B").End(xlUp).Row
    lastRowHAL = wsHAL.Cells(wsHAL.Rows.Count, "D").End(xlUp).Row
    
    ' Turn off screen updating for performance
    Application.ScreenUpdating = False
    
    ' Loop through DefcoInvoices sheet
    For i = 2 To lastRowDefco ' Assuming headers in row 1
        modelNumber = Trim(wsDefco.Cells(i, 2).value) ' Trim to remove extra spaces
        defcoPrice = CleanValue(wsDefco.Cells(i, 4).value) ' Clean and validate the price
        shopNumber = wsDefco.Cells(i, 3).value ' Column C
        found = False
        bestInvoice = "" ' Reset best match for this row
        
        ' Loop through HALDefcoSellin sheet
        For j = 2 To lastRowHAL
            ' Check if the model number exists in the description (case-insensitive)
            If InStr(1, Trim(wsHAL.Cells(j, 4).value), modelNumber, vbTextCompare) > 0 Then
                ' Check if the price in HAL is numeric and the defcoPrice is also numeric
                If IsNumeric(wsHAL.Cells(j, 5).value) And IsNumeric(defcoPrice) Then
                    ' Check if the price matches (with a small tolerance for floating-point precision)
                    If Abs(CDbl(wsHAL.Cells(j, 5).value) - defcoPrice) < 0.01 Then
                        ' Check if shop number matches
                        If wsHAL.Cells(j, 3).value = shopNumber Then
                            ' Store the invoice number as a potential best match
                            bestInvoice = wsHAL.Cells(j, 1).value
                            found = True
                            Exit For ' Exit inner loop if a match is found
                        End If
                    End If
                End If
            End If
        Next j
        
        ' Assign the best invoice number found, if any
        If found Then
            wsDefco.Cells(i, 5).value = bestInvoice
        Else
            wsDefco.Cells(i, 5).value = "Not Found"
        End If
    Next i
    
    ' Turn on screen updating after processing
    Application.ScreenUpdating = True
    
    MsgBox "Invoice numbers updated successfully!", vbInformation
End Sub

Function CleanValue(value As Variant) As Double
    Dim cleanedValue As String
    cleanedValue = Trim(CStr(value)) ' Convert to string and trim spaces
    ' Remove any non-numeric characters (except for decimal point)
    cleanedValue = Replace(cleanedValue, "$", "") ' Remove dollar sign
    cleanedValue = Replace(cleanedValue, ",", "") ' Remove commas
    ' Check if the cleaned value is numeric
    If IsNumeric(cleanedValue) Then
        CleanValue = CDbl(cleanedValue) ' Convert to Double
    Else
        CleanValue = 0 ' Return 0 or handle as needed
    End If
End Function
