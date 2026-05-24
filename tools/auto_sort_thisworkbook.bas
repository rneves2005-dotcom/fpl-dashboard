' ---- Paste this into the ThisWorkbook code module (NOT a regular module) ----
' Right-click "ThisWorkbook" in the VBA Project Explorer → View Code, then paste.

Private Sub Workbook_Open()
    StartAutoSort
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    StopAutoSort
End Sub

' Optional: reset the 5-min timer every time you edit a cell, so sort only
' fires after 5 min of INACTIVITY (avoids sort-while-typing surprises).
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If Sh.Name <> "Players" Then Exit Sub
    If Not SortEnabled Then Exit Sub
    ' Reset countdown
    On Error Resume Next
    Application.OnTime NextSortTime, "DoAutoSort", , False
    On Error GoTo 0
    NextSortTime = Now + TimeSerial(0, SORT_INTERVAL_MINUTES, 0)
    Application.OnTime NextSortTime, "DoAutoSort"
End Sub
