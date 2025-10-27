Option Explicit

#Const DEBUG_PAUSE_WAIT = False

#If VBA7 Then
    Private Declare PtrSafe Function SendMessageLongPtr Lib "user32" Alias "SendMessageA" ( _
        ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Private Declare PtrSafe Function SendMessageByRef Lib "user32" Alias "SendMessageA" ( _
        ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByRef lParam As Any) As LongPtr
    Private Declare PtrSafe Function GetFocus Lib "user32" () As LongPtr
    Private Declare PtrSafe Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongPtrA" ( _
        ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongPtrA" ( _
        ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Function SendMessageLongPtr Lib "user32" Alias "SendMessageA" ( _
        ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function SendMessageByRef Lib "user32" Alias "SendMessageA" ( _
        ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
    Private Declare Function GetFocus Lib "user32" () As Long
    Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
        ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
        ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
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
Private Const GWL_STYLE As Long = -16
Private Const WS_CAPTION As Long = &HC00000

'==== Private state ====
Private startTick As Double
Private lastUpdate As Double
Private emaSecPerItem As Double
Private Const SMOOTH As Double = 0.2   ' Exponential smoothing factor for ETA
Private maxBarWidth As Single          ' Captured from the design-time width
Private nextFormName As String
Private titleBarHidden As Boolean

Public TotalCount As Long
Public CompletedCount As Long

Public Paused As Boolean
Public Cancelled As Boolean

Private Sub Class_Initialize()
    Me.txtLog.ControlSource = ""
    nextFormName = ""
    titleBarHidden = False
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

    startTick = Timer
    lastUpdate = startTick
    emaSecPerItem = 0#
    TotalCount = totalCount
    CompletedCount = 0

    modReflectionsMonitor.PushCurrentStatus
End Sub

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

Private Function TextBoxIsScrolledToBottom(ByVal tb As MSForms.TextBox) As Boolean
#If VBA7 Then
    Dim hWndTB As LongPtr
#Else
    Dim hWndTB As Long
#End If
    hWndTB = GetTextBoxHwnd(tb)

    If hWndTB = 0 Then
        TextBoxIsScrolledToBottom = True
        Exit Function
    End If

    Dim totalLines As Long
    totalLines = CLng(SendMessageLongPtr(hWndTB, EM_GETLINECOUNT, 0&, 0&))
    If totalLines <= 1 Then
        TextBoxIsScrolledToBottom = True
        Exit Function
    End If

    Dim rc As RECT
    Call SendMessageByRef(hWndTB, EM_GETRECT, 0&, rc)

    Dim pt As POINTAPI
    pt.X = rc.Left + 1
    pt.Y = rc.Bottom - 1

    Dim charFromPos As Long
    charFromPos = CLng(SendMessageByRef(hWndTB, EM_CHARFROMPOS, 0&, pt))
    If charFromPos = -1 Then
        TextBoxIsScrolledToBottom = True
        Exit Function
    End If

    Dim lastVisibleLine As Long
    Dim lastVisibleChar As Long
    lastVisibleChar = charFromPos And &HFFFF&

    lastVisibleLine = (charFromPos And &HFFFF0000) \ &H10000
    If lastVisibleLine < 0 Then
        lastVisibleLine = lastVisibleLine + &H10000
    End If

    ' Fallback for any edge cases where the high-order extraction fails
    If lastVisibleLine < 0 Then
        lastVisibleLine = CLng(SendMessageLongPtr(hWndTB, EM_LINEFROMCHAR, lastVisibleChar, 0&))
    End If

    TextBoxIsScrolledToBottom = (lastVisibleLine >= totalLines - 1)
End Function

Private Function ApplyTextBoxSelection(ByVal tb As MSForms.TextBox, ByVal selStart As Long, ByVal selEnd As Long, Optional ByVal scrollCaret As Boolean = False) As Boolean
#If VBA7 Then
    Dim hWndTB As LongPtr
#Else
    Dim hWndTB As Long
#End If

    hWndTB = GetTextBoxHwnd(tb)

    If hWndTB = 0 Then
        Exit Function
    End If

    Call SendMessageLongPtr(hWndTB, EM_SETSEL, selStart, selEnd)

    If scrollCaret Then
        Call SendMessageLongPtr(hWndTB, EM_SCROLLCARET, 0&, 0&)
    End If

    ApplyTextBoxSelection = True
End Function

Private Sub RestoreSelection(ByVal tb As MSForms.TextBox, ByVal selStart As Long, ByVal selLength As Long)
    If ApplyTextBoxSelection(tb, selStart, selStart + selLength) Then
        Exit Sub
    End If

    On Error Resume Next
    tb.SelStart = selStart
    tb.SelLength = selLength
    On Error GoTo 0
End Sub

Private Sub ScrollTextBoxToBottom(ByVal tb As MSForms.TextBox)
    Dim textLen As Long
    textLen = Len(tb.Text)

    If ApplyTextBoxSelection(tb, textLen, textLen, True) Then
        Exit Sub
    End If

    On Error Resume Next
    tb.SelStart = textLen
    tb.SelLength = 0
    On Error GoTo 0
End Sub

Private Sub AppendLogEntry(ByVal target As MSForms.TextBox, ByVal newLine As String)
    Dim keepAtBottom As Boolean
    keepAtBottom = TextBoxIsScrolledToBottom(target)

    Dim originalSelStart As Long
    Dim originalSelLength As Long
    Dim originalLength As Long
    On Error Resume Next
    originalSelStart = target.SelStart
    originalSelLength = target.SelLength
    On Error GoTo 0

    originalLength = Len(target.Text)

    If originalLength > 0 Then
        target.Text = target.Text & vbCrLf & newLine
    Else
        target.Text = newLine
    End If

    Dim shouldStayAtBottom As Boolean
    shouldStayAtBottom = keepAtBottom Or (originalSelStart + originalSelLength >= originalLength)

    If shouldStayAtBottom Then
        ScrollTextBoxToBottom target
    Else
        RestoreSelection target, originalSelStart, originalSelLength
    End If
End Sub

' Append a time-stamped line to the log
Public Sub LogLine(ByVal lineText As String)
    Dim newLine As String
    newLine = Format$(Now, "hh:nn:ss") & "  " & CStr(lineText)

    AppendLogEntry Me.txtLog, newLine
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

    Dim isComplete As Boolean
    isComplete = (totalCount > 0 And done >= totalCount)

    CompletedCount = done
    TotalCount = totalCount

    If Cancelled Then
        btnCancel.Caption = "Cancelling..."
        btnCancel.Enabled = False
        btnPause.Visible = False
    ElseIf isComplete Then
        If btnCancel.Caption <> "Next" Then
            btnCancel.Caption = "Next"
        End If
        btnCancel.Enabled = True
        btnPause.Visible = False
    Else
        If btnCancel.Caption <> "Cancel" Then
            btnCancel.Caption = "Cancel"
        End If
        btnCancel.Enabled = True
        If Not btnPause.Visible Then
            btnPause.Visible = True
        End If
    End If

    DoEvents
End Sub

Public Property Get ProgressComplete() As Boolean
    ProgressComplete = (TotalCount > 0 And CompletedCount >= TotalCount)
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
    If btnCancel.Caption = "Next" Then
        nextFormName = "EmailForm"
        Unload Me
        Exit Sub
    End If

    Cancelled = True
    modProgressUI.cancelled = True
    btnCancel.Enabled = False
    btnCancel.Caption = "Cancelling..."
    btnPause.Visible = False
    LogLine "Cancel requested. Finishing current step"
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

    Me.MousePointer = fmMousePointerHourGlass

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

    'A_Record_Review

    Me.MousePointer = fmMousePointerDefault

End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ProgressForm.MousePointer = fmMousePointerDefault
End Sub

Private Sub UserForm_Terminate()
    On Error Resume Next
    modReflectionsMonitor.UnregisterReflectionsListener Me.Name
    On Error GoTo 0

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
                On Error GoTo 0
            End If
        Case "EmailForm"
            If modProgressUI.ProgressRunComplete() Then
                On Error Resume Next
                EmailForm.Show
                On Error GoTo 0
            End If
    End Select
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
    ' X button behaves like Cancel to avoid orphaned background work
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        btnCancel_Click
    End If
End Sub


Private Sub UserForm_Activate()
    If Not titleBarHidden Then
        HideTitleBar
    End If
End Sub

Private Sub HideTitleBar()
#If VBA7 Then
    Dim hWnd As LongPtr
    Dim currentStyle As LongPtr
    Dim newStyle As LongPtr
#Else
    Dim hWnd As Long
    Dim currentStyle As Long
    Dim newStyle As Long
#End If
    Dim originalCaption As String
    Dim tempCaption As String

    originalCaption = Me.Caption
    tempCaption = "progress-" & Hex$(ObjPtr(Me))
    Me.Caption = tempCaption

    hWnd = FindWindow("ThunderDFrame", tempCaption)
    Me.Caption = originalCaption

    If hWnd = 0 Then Exit Sub

    currentStyle = GetWindowLong(hWnd, GWL_STYLE)
    newStyle = currentStyle And (Not WS_CAPTION)
    SetWindowLong hWnd, GWL_STYLE, newStyle
    DrawMenuBar hWnd

    titleBarHidden = True
End Sub

