Sub RemoveAllTables()

    Dim ws As Worksheet
    Dim tbl As ListObject

    Application.ScreenUpdating = False

    For Each ws In ThisWorkbook.Worksheets
        Do While ws.ListObjects.Count > 0
            ws.ListObjects(1).Unlist
        Loop
    Next ws

    Application.ScreenUpdating = True

    MsgBox "All tables have been converted to normal ranges."

End Sub
