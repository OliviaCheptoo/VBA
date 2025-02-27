Sub FilterAndCreateSheets()
    Dim ws As Worksheet
    Dim uniqueShops As Collection
    Dim shopName As Variant
    Dim lastRow As Long
    Dim i As Long
    Dim newSheet As Worksheet
    Dim noShopsSheet As Worksheet
    
    ' Set the worksheet you want to filter
    Set ws = ThisWorkbook.Sheets("DefcoStocks") ' Change "Sheet1" to your sheet name
    
    ' Find the last row in column E
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    
    ' Create a collection to hold unique shop names
    Set uniqueShops = New Collection
    
    ' Loop through column E to get unique shop names
    On Error Resume Next ' Ignore errors for duplicate keys
    For i = 2 To lastRow ' Assuming row 1 is headers
        If ws.Cells(i, "E").Value = "" Then
            uniqueShops.Add "NoShops", "NoShops" ' Add to collection for blank values
        Else
            uniqueShops.Add ws.Cells(i, "E").Value, CStr(ws.Cells(i, "E").Value)
        End If
    Next i
    On Error GoTo 0 ' Resume normal error handling
    
    ' Loop through each unique shop name and create a new sheet
    For Each shopName In uniqueShops
        ' Check if the sheet already exists
        On Error Resume Next
        Set newSheet = ThisWorkbook.Sheets(shopName)
        On Error GoTo 0
        
        ' If the sheet does not exist, create it; if it exists, clear it
        If newSheet Is Nothing Then
            Set newSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            newSheet.Name = shopName
        Else
            newSheet.Cells.Clear ' Clear existing data
        End If
        
        ' Filter the original data and copy to the new sheet
        ws.Rows(1).AutoFilter Field:=5, Criteria1:=shopName ' Column E is the 5th column
        ws.Range("A1").CurrentRegion.SpecialCells(xlCellTypeVisible).Copy Destination:=newSheet.Range("A1")
        
        ' Clear the filter
        ws.AutoFilterMode = False
        
        ' Reset newSheet for the next iteration
        Set newSheet = Nothing
    Next shopName
    
    MsgBox "Sheets created for each shop!"
End Sub
