Dim modelCol As Range
Dim totalHits As Long
Dim unmatched As String

Sub CleanModelNumbers()
    Dim ws As Worksheet

    Set ws = ActiveSheet
    totalHits = 0
    unmatched = ""

    ' Column B, from row 2 down
    Set modelCol = ws.Range("B2:B" & ws.Cells(ws.Rows.Count, "B").End(xlUp).Row)

    Application.ScreenUpdating = False

    ' ReplaceModel "exact original cell text", "clean model number"
    ReplaceModel "DIA098E50B, DIB97IM50M", "DIB97IM50M"
    ReplaceModel "PCR915B91E, PCR9A5B90", "PCR9A5B90"
    ReplaceModel "GIN81AEF0G, GIN81AE3K", "GIN81AEF0G"
    ReplaceModel "H6040DECSP, H6040DECSP-T", "H6040DECSP"
    ReplaceModel "PRY626B70E, PRY6A6B70", "PRY6A6B70"
    ReplaceModel "KV694R, M3", "KV694R"
    ReplaceModel "DWP64CC60Z, DWW06W460B", "DWP64CC60Z"
    ReplaceModel "PRS926B70E, PRS9A6D70", "PRS9A6D70"
    ReplaceModel "DHZ5276, DWP94-50B", "DHZ5276"
    ReplaceModel "DWZ6HK1H4, DWK96PJ", "DWZ6HK1H4"
    ReplaceModel "F8, KIT2-D196, H37", "KIT2-D196"
    ReplaceModel "KV26N, M3", "KV26N"
    ReplaceModel "F3, D155, H16", "D155"
    ReplaceModel "DWZ6HP1U2, DWP64, DWP94-50M", "DWZ6HP1U2"
    ReplaceModel "DHZ2700, DHU665, DHU965", "DHZ2700"
    ReplaceModel "DIZ2CB1I4, DIB97", "DIZ2CB1I4"
    ReplaceModel "DWA0LK6A, DWK96", "DWA0LK6A"
    ReplaceModel "DHZ5385, DZZ2CB1B4", "DHZ5385"
    ReplaceModel "DWZ0DX0A0, DWB64", "DWZ0DX0A0"
    ReplaceModel "D233, H20", "D233"
    ReplaceModel "H9414, VBHS941X", "VBHS941X"
    ReplaceModel "DWZ0GK0A0, DWK96", "DWZ0GK0A0"
    ReplaceModel "H6404VGRM, VBHS6401X", "VBHS6401X"
    ReplaceModel "B640, CWBJ-26449, HAH902BG01, HAH602BG01", "CWBJ-26449"
    ReplaceModel "B642, KSC01BLMWW", "KSC01BLMWW"
    ReplaceModel "CWBJ-26449, HAH902BG01, HAH602BG01", "CWBJ-26449"
    ReplaceModel "KSG6P4MB, M3", "KSG6P4MB"
    ReplaceModel "DWW09W450B, DWP94BC50B", "DWP94BC50B"

    Application.ScreenUpdating = True

    If Len(unmatched) > 0 Then
        MsgBox "Replaced " & totalHits & " cell(s)." & vbLf & vbLf & _
               "No match (not in this data):" & unmatched, vbInformation, "Model number cleanup"
    Else
        MsgBox "Replaced " & totalHits & " cell(s).", vbInformation, "Model number cleanup"
    End If
End Sub

Private Sub ReplaceModel(findText As String, replaceText As String)
    Dim hits As Long
    hits = Application.WorksheetFunction.CountIf(modelCol, findText)
    If hits = 0 Then
        unmatched = unmatched & vbLf & "  " & findText
    Else
        modelCol.Replace What:=findText, Replacement:=replaceText, _
            LookAt:=xlWhole, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    End If
    totalHits = totalHits + hits
End Sub
