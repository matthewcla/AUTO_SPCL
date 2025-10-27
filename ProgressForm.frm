Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function SendMessageLongPtr Lib "user32" Alias "SendMessageA" ( _
        ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Private Declare PtrSafe Function SendMessageByRef Lib "user32" Alias "SendMessageA" ( _
        ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByRef lParam As Any) As LongPtr
#Else
    Private Declare Function SendMessageLongPtr Lib "user32" Alias "SendMessageA" ( _
        ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function SendMessageByRef Lib "user32" Alias "SendMessageA" ( _
        ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
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

'==== Private state ====
Private startTick As Double
Private lastUpdate As Double
Private emaSecPerItem As Double
Private Const SMOOTH As Double = 0.2   ' Exponential smoothing factor for ETA
Private maxBarWidth As Single          ' Captured from the design-time width

Private Sub Class_Initialize()
    Me.txtLog.ControlSource = ""
End Sub

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

    Me.txtLog.ControlSource = ""
    Me.txtLog.Value = ""
    Me.txtLog.SelStart = 0

    Paused = False
    Cancelled = False
    modProgressUI.cancelled = False
    Me.btnPause.Caption = "Pause"

    startTick = Timer
    lastUpdate = startTick
    emaSecPerItem = 0#
End Sub

Private Function TextBoxIsScrolledToBottom(ByVal tb As MSForms.TextBox) As Boolean
#If VBA7 Then
    Dim hWndTB As LongPtr
#Else
    Dim hWndTB As Long
#End If
    hWndTB = tb.hwnd

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

Private Sub RestoreSelection(ByVal tb As MSForms.TextBox, ByVal selStart As Long, ByVal selLength As Long)
    On Error Resume Next
    tb.SelStart = selStart
    tb.SelLength = selLength
    On Error GoTo 0
End Sub

Private Sub ScrollTextBoxToBottom(ByVal tb As MSForms.TextBox)
    On Error Resume Next
    tb.SelStart = Len(tb.Text)
    tb.SelLength = 0
    On Error GoTo 0
End Sub

Private Sub AppendLogEntry(ByVal target As MSForms.TextBox, ByVal newLine As String)
    Dim keepAtBottom As Boolean
    keepAtBottom = TextBoxIsScrolledToBottom(target)

    Dim originalSelStart As Long
    Dim originalSelLength As Long
    On Error Resume Next
    originalSelStart = target.SelStart
    originalSelLength = target.SelLength
    On Error GoTo 0

    If Len(target.Text) > 0 Then
        target.Text = target.Text & vbCrLf & newLine
    Else
        target.Text = newLine
    End If

    If keepAtBottom Then
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
    modProgressUI.cancelled = True
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

    Me.txtLog.ControlSource = ""

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


