Option Explicit
'-------------------------------------------------------------------------------
' Form: ProgressForm
' Role   : Modeless status window orchestrated by modProgressUI to visualize the
'          record review pipeline, surface logs, and allow pause/cancel control.
' Coordinates:
'   - Receives updates and lifecycle events from MainMod.A_Record_Review via modProgressUI.
'   - Launches EmailForm when review work finishes so drafting can begin immediately.
'   - Relies on modUIHelpers for title bar suppression and shared button state helpers.
'-------------------------------------------------------------------------------

#Const DEBUG_PAUSE_WAIT = False

#If VBA7 Then
    Private Declare PtrSafe Function SendMessageLongPtr Lib "user32" Alias "SendMessageA" ( _
        ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Private Declare PtrSafe Function SendMessageByRef Lib "user32" Alias "SendMessageA" ( _
        ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByRef lParam As Any) As LongPtr
    Private Declare PtrSafe Function GetFocus Lib "user32" () As LongPtr
    Private Declare PtrSafe Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Function SendMessageLongPtr Lib "user32" Alias "SendMessageA" ( _
        ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function SendMessageByRef Lib "user32" Alias "SendMessageA" ( _
        ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
    Private Declare Function GetFocus Lib "user32" () As Long
    Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const EM_GETLINECOUNT As Long = &HBA
Private Const EM_LINEFROMCHAR As Long = &HC9
Private Const EM_CHARFROMPOS As Long = &HD7
Private Const EM_GETRECT As Long = &HB2
Private Const EM_SETSEL As Long = &HB1
Private Const EM_SCROLLCARET As Long = &HB7
'==== Private state ====
Private maxBarWidth As Single          ' Captured from the design-time width
Private nextFormName As String
Private mTitleBarHidden As Boolean

Private sessionStart As Date
Private pauseStarted As Date
Private pausedSeconds As Double
Private startTick As Double
Private lastUpdate As Double

#If VBA7 Then
    Private timerId As LongPtr
#Else
    Private timerId As Long
#End If
Private timerBusy As Boolean

Public TotalCount As Long
Public CompletedCount As Long

Public Paused As Boolean
Public Cancelled As Boolean

Private Function TryGetFormControl(ByVal controlName As String) As MSForms.Control
    On Error Resume Next
    Set TryGetFormControl = Me.Controls(controlName)
    On Error GoTo 0
End Function

Private Function EnsureRequiredControls() As Boolean
    Dim missing As Collection

    Set missing = New Collection

    ValidateControlExists "txtLog", "TextBox", missing
    ValidateControlExists "lblProcessed", "Label", missing
    ValidateControlExists "lblRemaining", "Label", missing
    ValidateControlExists "lblPercentage", "Label", missing
    ValidateControlExists "lblElapsed", "Label", missing
    ValidateControlExists "lblETR", "Label", missing
    ValidateControlExists "lblProcessedBarFill", "Label", missing
    ValidateControlExists "btnPause", "CommandButton", missing
    ValidateControlExists "btnCancel", "CommandButton", missing
    ValidateControlExists "lblOAIS", "Label", missing

    EnsureRequiredControls = missing.Count = 0

    If Not EnsureRequiredControls Then
        modUIHelpers.ShowErrorMessage "AUTO_SPCL can't display the progress tracker because these controls are missing:" & _
                                      vbCrLf & " - " & JoinCollectionItems(missing, vbCrLf & " - ") & vbCrLf & _
                                      "Please contact the workbook administrator to restore them."
    End If
End Function

Private Sub ValidateControlExists(ByVal controlName As String, _
                                  ByVal expectedType As String, _
                                  ByRef missing As Collection)
    Dim ctrl As MSForms.Control

    Set ctrl = TryGetFormControl(controlName)
    If ctrl Is Nothing Then
        missing.Add controlName & " (" & expectedType & ")"
    ElseIf StrComp(TypeName(ctrl), expectedType, vbTextCompare) <> 0 Then
        missing.Add controlName & " (expected " & expectedType & ")"
    End If
End Sub

Private Function JoinCollectionItems(ByVal items As Collection, Optional ByVal delimiter As String = ", ") As String
    Dim entry As Variant
    Dim buffer As String

    If items Is Nothing Then Exit Function

    For Each entry In items
        If LenB(buffer) > 0 Then buffer = buffer & delimiter
        buffer = buffer & CStr(entry)
    Next entry

    JoinCollectionItems = buffer
End Function

Private Sub Class_Initialize()
    If Not EnsureRequiredControls() Then
        Err.Raise vbObjectError + 801, "ProgressForm.Class_Initialize", _
                  "Required controls are missing from ProgressForm."
    End If

    Me.txtLog.ControlSource = ""
    nextFormName = ""
    mTitleBarHidden = False
    timerId = 0
    timerBusy = False
    Paused = False
    Cancelled = False
    sessionStart = Now
    pauseStarted = 0
    pausedSeconds = 0#
    startTick = 0#
    lastUpdate = 0#
End Sub

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
    HMS = Format$(h, "00") & ":" & Format$(m, "00") & ":" & Format$(s, "00")
End Function

' Call once, right after showing modeless
Public Sub Init(totalCount As Long, Optional captionText As String = "Reviewing records")
    Me.Caption = captionText
    Me.lblProcessed.Caption = "0"
    Me.lblRemaining.Caption = CStr(totalCount)
    Me.lblPercentage.Caption = "0%"
    Me.lblElapsed.Caption = "00:00:00"
    Me.lblETR.Caption = "--:--:--"
    lblOAIS.Caption = ""

    CenterUserFormOnActiveMonitor Me
    
    ' Capture the design-time width of the bar as the maximum; then collapse to zero
    maxBarWidth = lblProcessedBarFill.Width
    lblProcessedBarFill.Width = 0

    Me.txtLog.ControlSource = ""
    Me.txtLog.Value = ""
    Me.txtLog.SelStart = 0

    Paused = False
    Cancelled = False
    modProgressUI.cancelled = False
    Me.btnPause.Caption = "Pause"
    Me.btnPause.Visible = True
    Me.btnCancel.Caption = "Cancel"
    Me.btnCancel.Enabled = True
    nextFormName = ""

    sessionStart = Now
    pauseStarted = 0
    pausedSeconds = 0#
    startTick = Timer
    lastUpdate = startTick
    Me.TotalCount = totalCount
    Me.CompletedCount = 0

    modReflectionsMonitor.PushCurrentStatus
    modProgressUI.Progress_StartTimer
End Sub

Friend Sub ShutdownTimer()
    modProgressUI.Progress_StopTimer
End Sub

Public Sub Tick_OneSecond()
    ' Lightweight 1 Hz timer hook: only touch elapsed/ETA labels (<10ms).
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    On Error GoTo HandleError

    Dim nowT As Double
    nowT = Timer
    If nowT < lastUpdate Then nowT = nowT + 86400#

    UpdateElapsedAndEta nowT
    lastUpdate = nowT

    Exit Sub

HandleError:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description

    If errNumber <> 0 Then
        Err.Clear
        Err.Raise errNumber, errSource, errDescription
    End If
End Sub

Friend Sub UpdateElapsedAndEta(ByVal nowT As Double, Optional ByVal currentTime As Date = 0)
    Dim elapsed As Double
    If currentTime = 0 Then
        elapsed = ActiveElapsedSeconds()
    Else
        elapsed = ActiveElapsedSeconds(currentTime)
    End If
    lblElapsed.Caption = HMS(elapsed)

    Dim pctComplete As Double
    If Me.TotalCount <= 0 Then
        pctComplete = 0#
    Else
        pctComplete = CompletedCount / Me.TotalCount
        pctComplete = Application.Max(Application.Min(pctComplete, 1#), 0#)
    End If

    Dim remain As Double
    Dim etrText As String
    If Me.TotalCount <= 0 Or pctComplete <= 0# Then
        etrText = "--:--:--"
    ElseIf pctComplete >= 1# Then
        etrText = "00:00:00"
    Else
        remain = elapsed * (1# - pctComplete) / pctComplete
        If remain < 0# Then remain = 0#
        etrText = HMS(remain)
    End If

    lblETR.Caption = etrText
End Sub

Private Sub RefreshTimingDisplays(Optional ByVal currentTime As Date = 0)
    Dim tickNow As Double
    tickNow = Timer
    If tickNow < lastUpdate Then tickNow = tickNow + 86400#

    UpdateElapsedAndEta tickNow, currentTime
    lastUpdate = tickNow
End Sub

Private Function ActiveElapsedSeconds(Optional ByVal currentTime As Date = 0) As Double
    Dim baseTime As Date
    If currentTime = 0 Then
        baseTime = Now
    Else
        baseTime = currentTime
    End If

    Dim elapsedSeconds As Double
    elapsedSeconds = (baseTime - sessionStart) * 86400#
    If elapsedSeconds < 0# Then elapsedSeconds = 0#

    Dim totalPaused As Double
    totalPaused = pausedSeconds
    If Paused And pauseStarted <> 0 Then
        totalPaused = totalPaused + (baseTime - pauseStarted) * 86400#
    End If

    elapsedSeconds = elapsedSeconds - totalPaused
    If elapsedSeconds < 0# Then elapsedSeconds = 0#

    ActiveElapsedSeconds = elapsedSeconds
End Function

Private Function GetTextBoxHwnd(ByVal tb As MSForms.TextBox) As LongPtr
#If VBA7 Then
    Dim previousFocus As LongPtr
    Dim tbHandle As LongPtr
#Else
    Dim previousFocus As Long
    Dim tbHandle As Long
#End If

    On Error Resume Next
    tbHandle = CallByName(tb, "hwnd", VbGet)
    If Err.Number = 0 And tbHandle <> 0 Then
        On Error GoTo 0
        GetTextBoxHwnd = tbHandle
        Exit Function
    End If
    Err.Clear

    previousFocus = GetFocus()

    tb.SetFocus
    On Error GoTo 0

    tbHandle = GetFocus()

    If previousFocus <> 0 Then
        On Error Resume Next
        SetFocusAPI previousFocus
        On Error GoTo 0
    End If

    GetTextBoxHwnd = tbHandle
End Function

Private Sub AppendIssueLine(ByVal target As MSForms.TextBox, ByVal lineText As String)
    target.Text = target.Text & lineText & vbCrLf

    On Error Resume Next
    target.SelStart = Len(target.Text)
    target.SelLength = 0
    On Error GoTo 0
End Sub

' Append a log line, optionally deduplicating identical consecutive entries.
Public Sub LogLine(ByVal lineText As String)
    Static lastLoggedLine As String

    If LenB(lineText) = 0 Then
        Exit Sub
    End If

    If lineText = lastLoggedLine Then
        Exit Sub
    End If

    AppendIssueLine Me.txtLog, CStr(lineText)
    lastLoggedLine = lineText
End Sub

' Update counters, percent, bar, elapsed, ETA. Call this once per record (or more).
Public Sub UpdateProgress(ByVal done As Long, ByVal totalCount As Long, Optional ByVal status As String = "")
    Dim currentTime As Date
    currentTime = Now

    ' Numbers
    lblProcessed.Caption = CStr(done)
    lblRemaining.Caption = CStr(Application.Max(totalCount - done, 0))

    ' Percent & bar
    Dim pct As Double: pct = 0#
    If totalCount > 0 Then pct = done / totalCount
    lblPercentage.Caption = Format$(pct, "0%")
    lblProcessedBarFill.Width = maxBarWidth * pct

    ' Optional status line no longer written to the log

    Dim isComplete As Boolean
    isComplete = (totalCount > 0 And done >= totalCount)

    Me.CompletedCount = done
    Me.TotalCount = totalCount

    RefreshTimingDisplays currentTime

    UpdateButtonStates isComplete

    DoEvents
End Sub

Private Sub UpdateButtonStates(ByVal isComplete As Boolean)
    modProgressUI.UpdateProgressButtonStates Me.btnCancel, Me.btnPause, Cancelled, isComplete
End Sub

Public Property Get ProgressComplete() As Boolean
    ProgressComplete = (Me.TotalCount > 0 And Me.CompletedCount >= Me.TotalCount)
End Property

' Blocks while paused; returns False if cancelled while waiting
Public Function WaitIfPaused() As Boolean
    Const SLICE_MS As Long = 25
#If DEBUG_PAUSE_WAIT Then
    Dim waitStart As Double
    Dim lastLogTick As Double
    waitStart = Timer
    lastLogTick = waitStart
#End If

    Do While Paused And Not Cancelled
        If pauseStarted = 0 Then
            pauseStarted = Now
        End If
        DoEvents
        Sleep SLICE_MS
#If DEBUG_PAUSE_WAIT Then
        Dim nowTick As Double
        nowTick = Timer
        If nowTick < lastLogTick Then
            lastLogTick = nowTick
            waitStart = nowTick
        End If
        If nowTick - lastLogTick >= 1# Then
            Debug.Print "WaitIfPaused running for", Format$(nowTick - waitStart, "0.0"), "seconds"
            lastLogTick = nowTick
        End If
#End If
    Loop
    If pauseStarted <> 0 Then
        pausedSeconds = pausedSeconds + (Now - pauseStarted) * 86400#
        pauseStarted = 0
    End If
    If Not Cancelled Then
        RefreshTimingDisplays
    End If
    WaitIfPaused = Not Cancelled
End Function

Private Sub btnPause_Click()
    Paused = Not Paused
    btnPause.Caption = IIf(Paused, "Resume", "Pause")

    Dim resumeTick As Double

    If Paused Then
        modProgressUI.LogRecordReviewPaused
        modProgressUI.Progress_StopTimer
    Else
        modProgressUI.LogRecordReviewResume
        resumeTick = Timer
        Dim adjustedTick As Double
        adjustedTick = resumeTick
        If adjustedTick < startTick Then adjustedTick = adjustedTick + 86400#
        startTick = resumeTick
        lastUpdate = adjustedTick
        modProgressUI.Progress_StartTimer
    End If
End Sub

Private Sub btnCancel_Click()
    If btnCancel.Caption = "Next" Then
        nextFormName = "EmailForm"
        Unload Me
        Exit Sub
    End If

    Cancelled = True
    modProgressUI.cancelled = True
    modProgressUI.Progress_StopTimer
    modProgressUI.LogRecordReviewStatus "Cancellation Triggered"
    modProgressUI.LogRecordReviewCancelled
    btnCancel.Enabled = False
    btnCancel.Caption = "Cancelling..."
    btnPause.Visible = False
    nextFormName = "StartupForm"
    ' Keep the form visible until the worker loop finishes and closes it.
End Sub

Private Sub bOAIS_Click()
    ' If not connected, try to connect; else toggle external frame if present.
    If lblOAIS.BackColor = vbRed Then
        EnsureReflectionsConnectionAlive True
        UpdateOAISStatusIndicator
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
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    On Error GoTo CleanFail

    SetCursorWait

    modProgressUI.Progress_ResetTimerState

    Me.txtLog.ControlSource = ""
    lblOAIS.Caption = ""

    modReflectionsMonitor.RegisterReflectionsListener Me.Name

    Dim isConnected As Boolean
    isConnected = EnsureReflectionsConnectionAlive(True)
    HandleReflectionsConnection isConnected

    If isConnected Then
        InitializeOAISSession lblOAIS, "", "", vbGreen, vbWhite
    Else
        UpdateOAISStatusIndicator
    End If

    startTick = Timer
    lastUpdate = startTick
    modProgressUI.Progress_StartTimer

    'A_Record_Review

CleanExit:
    SetCursorDefault
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
    Exit Sub

CleanFail:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    modProgressUI.Progress_StopTimer
    Resume CleanExit
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ProgressForm.MousePointer = fmMousePointerDefault
End Sub

Private Sub UserForm_Terminate()
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    On Error GoTo CleanFail

    modProgressUI.Progress_ResetTimerState

    SetCursorWait

    On Error Resume Next
    modReflectionsMonitor.UnregisterReflectionsListener Me.Name
    On Error GoTo CleanFail

    Dim targetForm As String
    targetForm = nextFormName
    nextFormName = ""

    Select Case targetForm
        Case "StartupForm"
            Dim allowStartup As Boolean
            allowStartup = modProgressUI.ProgressRunComplete()

            ' Allow navigation back to StartupForm when cancellation paths requested it
            If Not allowStartup Then
                allowStartup = (StrComp(targetForm, "StartupForm", vbTextCompare) = 0)
            End If

            ' Or when the module has flagged a user cancel
            If Not allowStartup Then
                allowStartup = modProgressUI.cancelled
            End If

            If allowStartup Then
                On Error Resume Next
                StartupForm.Show
                On Error GoTo CleanFail
            End If
        Case "EmailForm"
            If modProgressUI.ProgressRunComplete() Then
                On Error Resume Next
                EmailForm.Show
                On Error GoTo CleanFail
            End If
    End Select

CleanExit:
    SetCursorDefault
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
    Exit Sub

CleanFail:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    modProgressUI.Progress_ResetTimerState
    Resume CleanExit
End Sub

Public Sub HandleReflectionsConnection(ByVal isConnected As Boolean)
    lblOAIS.Caption = ""
    If isConnected Then
        If lblOAIS.ForeColor <> vbGreen Then
            lblOAIS.ForeColor = vbGreen
        End If
        If lblOAISCap.Caption <> "Connected to OAIS" Then
            lblOAISCap.Caption = "Connected to OAIS"
        End If
        lblOAIS.BackColor = vbGreen
    Else
        lblOAIS.ForeColor = vbWhite
        lblOAISCap.Caption = "OAIS is Disconnected= ""
        lblOAIS.BackColor = vbRed
    End If
End Sub

Private Sub UpdateOAISStatusIndicator()
    HandleReflectionsConnection Not (iCS Is Nothing)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    On Error GoTo CleanFail

    SetCursorWait

    ' X button behaves like Cancel to avoid orphaned background work
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        btnCancel_Click
    End If

CleanExit:
    SetCursorDefault
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
    Exit Sub

CleanFail:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Resume CleanExit
End Sub


Private Sub UserForm_Activate()
    modUIHelpers.HideUserFormTitleBar Me, mTitleBarHidden, "progress"
End Sub

