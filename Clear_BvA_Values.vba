Sub ClearValuesKeepFormulas()

    Dim ws As Worksheet
    Dim checkRange As Range
    Dim constants As Range

    Set ws = ThisWorkbook.Sheets("Workings")

    ' === SETTINGS ===
    Const START_ROW As Long = 5
    Const END_ROW As Long = 400
    Const START_COL As String = "BT"
    Const END_COL As String = "IF"
    ' =================

    ' Define range to check
    Set checkRange = ws.Range( _
        START_COL & START_ROW & ":" & _
        END_COL & END_ROW _
    )

    ' Find cells containing hard-coded values
    ' Formula cells are automatically excluded
    On Error Resume Next
    Set constants = checkRange.SpecialCells(xlCellTypeConstants)
    On Error GoTo 0

    ' Clear hard-coded values
    If Not constants Is Nothing Then
        constants.ClearContents
    End If

    ' Notify when finished
    MsgBox "Done! Hard-coded values have been cleared." & vbCrLf & _
           "Formula cells were left untouched.", _
           vbInformation, _
           "Completed"

End Sub
