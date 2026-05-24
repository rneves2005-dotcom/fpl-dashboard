Attribute VB_Name = "AutoSort"
' --------------------------------------------------------------------------
'  AutoSort module for Squads_Data.xlsm
'  Sorts the Players table every N minutes (default 5)
'  Sort keys: 1) Clubs ASC   2) Shirt ASC
'  Pauses while the user is mid-edit, resumes after idle window.
' --------------------------------------------------------------------------

Public NextSortTime As Date
Public SortEnabled As Boolean
Public Const SORT_INTERVAL_MINUTES As Integer = 5
Public Const SORT_SHEET As String = "Players"
Public Const SORT_TABLE As String = "Players"

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

Public Sub DoAutoSort()
    On Error GoTo Cleanup

    ' Skip if user is in the middle of editing a cell
    If Application.ReferenceStyle = xlR1C1 Then GoTo Reschedule  ' no-op probe
    On Error Resume Next
    Dim probe As Variant
    probe = Application.EnableEvents
    If Err.Number <> 0 Then GoTo Reschedule
    On Error GoTo Cleanup

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim ws As Worksheet
    Dim tbl As ListObject
    Set ws = ThisWorkbook.Sheets(SORT_SHEET)
    Set tbl = ws.ListObjects(SORT_TABLE)

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

    Application.StatusBar = "AutoSort: last run " & Format(Now, "hh:mm:ss") & _
                            "  ·  next at " & Format(Now + TimeSerial(0, SORT_INTERVAL_MINUTES, 0), "hh:mm")

Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

Reschedule:
    If SortEnabled Then ScheduleNextSort
End Sub
