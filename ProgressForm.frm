Option Explicit

'==== Private state ====
Private startTick As Double
Private lastUpdate As Double
Private emaSecPerItem As Double
Private Const SMOOTH As Double = 0.2   ' Exponential smoothing factor for ETA
Private maxBarWidth As Single          ' Captured from the design-time width

Public Paused As Boolean
Public Cancelled As Boolean

' Utility: format seconds as h:mm:ss
Private Function HMS(ByVal secs As Double) As String
    If secs < 0 Or Not IsNumeric(secs) Then
        HMS = "--:--:--"
        Exit Function
    End If
    Dim h As Long, m As Long, s As Long
    h = CLng(secs \ 3600)
    m = CLng((secs - h * 3600) \ 60)
    s = CLng(secs - h * 3600 - m * 60)
    HMS = h & ":" & Format$(m, "00") & ":" & Format$(s, "00")
End Function

' Call once, right after showing modeless
Public Sub Init(totalCount As Long, Optional captionText As String = "Reviewing records")
    Me.Caption = captionText
    Me.lblProcessed.Caption = "0"
    Me.lblRemaining.Caption = CStr(totalCount)
    Me.lblPercentage.Caption = "0%"
    Me.lblElapsed.Caption = "0:00:00"
    Me.lblETR.Caption = "--:--:--"

    CenterUserFormOnActiveMonitor Me
    
    ' Capture the design-time width of the bar as the maximum; then collapse to zero
    maxBarWidth = lblProcessedBarFill.Width
    lblProcessedBarFill.Width = 0

    Me.txtLog.ControlSource = vbNullString
    Me.txtLog.Value = vbNullString
    Me.txtLog.SelStart = 0

    Paused = False
    Cancelled = False
    Me.btnPause.Caption = "Pause"

    startTick = Timer
    lastUpdate = startTick
    emaSecPerItem = 0#
End Sub

' Append a time-stamped line to the log
Public Sub LogLine(ByVal lineText As String)
    Dim newLine As String
    newLine = Format$(Now, "hh:nn:ss") & "  " & CStr(lineText)

    With Me.txtLog
        If Len(.Text) > 0 Then
            .Text = .Text & vbCrLf & newLine  ' Append the line
        Else
            .Text = newLine  ' Start with the first line
        End If
        .SelStart = Len(.Text)
        .SelLength = 0
    End With
    DoEvents
End Sub

' Update counters, percent, bar, elapsed, ETA. Call this once per record (or more).
Public Sub UpdateProgress(ByVal done As Long, ByVal totalCount As Long, Optional ByVal status As String = "")
    Dim nowT As Double: nowT = Timer
    If nowT < lastUpdate Then nowT = nowT + 86400# ' handle midnight rollover

    ' Update EMA of seconds per item for smoother ETA
    If done > 0 Then
        Dim secPerItem As Double
        secPerItem = (nowT - startTick) / done
        If emaSecPerItem = 0 Then
            emaSecPerItem = secPerItem
        Else
            emaSecPerItem = (1 - SMOOTH) * emaSecPerItem + SMOOTH * secPerItem
        End If
    End If
    lastUpdate = nowT

    ' Numbers
    lblProcessed.Caption = CStr(done)
    lblRemaining.Caption = CStr(Application.Max(totalCount - done, 0))

    ' Percent & bar
    Dim pct As Double: pct = 0#
    If totalCount > 0 Then pct = done / totalCount
    lblPercentage.Caption = Format$(pct, "0%")
    lblProcessedBarFill.Width = maxBarWidth * pct

    ' Time
    Dim elapsed As Double: elapsed = nowT - startTick
    If elapsed < 0 Then elapsed = elapsed + 86400#
    lblElapsed.Caption = HMS(elapsed)

    Dim remain As Double
    If totalCount > 0 Then
        remain = (totalCount - done) * IIf(emaSecPerItem > 0, emaSecPerItem, 0)
    Else
        remain = 0
    End If
    lblETR.Caption = HMS(remain)

    ' Optional status line to log
    If Len(status) > 0 Then LogLine (status)

    DoEvents
End Sub

' Blocks while paused; returns False if cancelled while waiting
Public Function WaitIfPaused() As Boolean
    Do While Paused And Not Cancelled
        DoEvents
        Application.Wait Now + TimeValue("0:00:00") ' yield without sleeping too long
    Loop
    WaitIfPaused = Not Cancelled
End Function

Private Sub btnPause_Click()
    Paused = Not Paused
    btnPause.Caption = IIf(Paused, "Resume", "Pause")
    If Paused Then
        LogLine "Paused by user."
    Else
        LogLine "Resumed."
    End If
End Sub

Private Sub btnCancel_Click()
    Cancelled = True
    btnCancel.Enabled = False
    LogLine "Cancel requested. Finishing current step"
End Sub

Private Sub bSettings_Click()
    ToggleThisWorkbookVisibility
End Sub


Private Sub bOAIS_Click()
    ' If not connected, try to connect; else toggle external frame if present.
    If lblOAIS.BackColor = vbRed Then
        ConnectToRunningOAIS
        SetOAISStatus lblOAIS, Not (iCS Is Nothing), , , vbWhite, vbWhite
        Exit Sub
    End If

    ' Optional: toggle an external host frame if your environment exposes one.
    On Error Resume Next
    If Not (iFrame Is Nothing) Then
        ' Late-bound: property may not exist in all hosts; safe-guarded.
        If LCase$(CStr(CallByName(iFrame, "WindowState", VbGet))) = "0" Then
            CallByName iFrame, "WindowState", VbLet, 1   ' minimize
        Else
            CallByName iFrame, "WindowState", VbLet, 0   ' normal
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub UserForm_Initialize()

    Me.txtLog.ControlSource = vbNullString

    InitializeOAISSession lblOAIS, , , vbWhite, vbWhite

    'A_Record_Review

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' X button behaves like Cancel to avoid orphaned background work
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        btnCancel_Click
    End If
End Sub


