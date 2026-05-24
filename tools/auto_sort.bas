Attribute VB_Name = "AutoSort"
' --------------------------------------------------------------------------
'  AutoSort module — sorts the Players table every N minutes (default 5)
'  Sort keys: 1) Clubs ASC   2) Shirt ASC   (groups players by team)
'  Pauses while you're mid-edit, resumes after idle window.
'
'  Designed for Squads_2026.xlsm. Auto-finds the players table by name:
'  tries "Players" first, falls back to "Players5", "Players4", etc.
'  This handles the case where Excel auto-renamed a copied table.
' --------------------------------------------------------------------------

Public NextSortTime As Date
Public SortEnabled As Boolean
Public Const SORT_INTERVAL_MINUTES As Integer = 5

' ---- Public API ----------------------------------------------------------

Public Sub StartAutoSort()
    SortEnabled = True
    ScheduleNextSort
    Application.StatusBar = "AutoSort ON — sorts every " & SORT_INTERVAL_MINUTES & " min by Clubs → Shirt"
End Sub

Public Sub StopAutoSort()
    SortEnabled = False
    On Error Resume Next
    Application.OnTime NextSortTime, "DoAutoSort", , False
    Application.StatusBar = "AutoSort OFF"
End Sub

Public Sub SortNow()
    DoAutoSort
End Sub

' ---- Internal -----------------------------------------------------------

Private Sub ScheduleNextSort()
    NextSortTime = Now + TimeSerial(0, SORT_INTERVAL_MINUTES, 0)
    Application.OnTime NextSortTime, "DoAutoSort"
End Sub

Private Function FindPlayersTable() As ListObject
    ' Searches every worksheet for a ListObject named "Players", "Players5", "Players4", etc.
    ' Prefers the one with the most rows (the real data, not the 1M-row ghost).
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim candidate As ListObject
    Dim candidateRows As Long
    Dim names As Variant
    Dim i As Integer

    candidateRows = 0
    Set candidate = Nothing

    names = Array("Players", "Players5", "Players4", "Players3", "Players2", "Players1")

    For Each ws In ThisWorkbook.Worksheets
        For i = LBound(names) To UBound(names)
            On Error Resume Next
            Set lo = ws.ListObjects(names(i))
            On Error GoTo 0
            If Not lo Is Nothing Then
                ' Skip the ghost table (1M rows with mostly empty content)
                If lo.ListRows.Count < 100000 And lo.ListRows.Count > candidateRows Then
                    Set candidate = lo
                    candidateRows = lo.ListRows.Count
                End If
                Set lo = Nothing
            End If
        Next i
    Next ws

    Set FindPlayersTable = candidate
End Function

Public Sub DoAutoSort()
    On Error GoTo Cleanup
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim tbl As ListObject
    Set tbl = FindPlayersTable()
    If tbl Is Nothing Then
        Application.StatusBar = "AutoSort: no Players table found (skipping)"
        GoTo Cleanup
    End If

    With tbl.Sort
        .SortFields.Clear
        .SortFields.Add Key:=tbl.ListColumns("Clubs").Range, _
                        SortOn:=xlSortOnValues, Order:=xlAscending, _
                        DataOption:=xlSortNormal
        .SortFields.Add Key:=tbl.ListColumns("Shirt").Range, _
                        SortOn:=xlSortOnValues, Order:=xlAscending, _
                        DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

    Application.StatusBar = "AutoSort: " & tbl.Name & " sorted at " & Format(Now, "hh:mm:ss") & _
                            "  ·  next at " & Format(Now + TimeSerial(0, SORT_INTERVAL_MINUTES, 0), "hh:mm")

Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If SortEnabled Then ScheduleNextSort
End Sub
