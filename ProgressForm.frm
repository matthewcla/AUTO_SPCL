VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressForm 
   Caption         =   "OAIS Macro Progress Tracker"
   ClientHeight    =   11115
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10485
   Enabled         =   0   'False
   OleObjectBlob   =   "ProgressForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' --- Constants for screen-text detection (Reflection / OAIS banners) ---
Private Const TXT_DISA As String = "Defense Information Systems Agency"
Private Const TXT_SESSION_MENU As String = "CL/SuperSession"
Private Const TXT_OAIS_BANNER As String = "Officer Assignment Information System"
Private Const OAIS_MENU_CMD As String = "Start OAIS2"

' --- Tuning knobs ---
Private Const SMALL_WAIT_SEC As Double = 0.75
Private Const RETRY_WAIT_SEC As Double = 1.2
Private Const retries As Long = 3

'==== Private state ====
Private startTick As Double
Private lastUpdate As Double
Private emaSecPerItem As Double
Private Const SMOOTH As Double = 0.2   ' Exponential smoothing factor for ETA
Private maxBarWidth As Single          ' Captured from the design-time width

Public Paused As Boolean
Public Cancelled As Boolean

'— Utility: format seconds as h:mm:ss
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

'— Call once, right after showing modeless
Public Sub Init(totalCount As Long, Optional captionText As String = "Reviewing records…")
    Me.Caption = captionText
    progressform.lblProcessed.Caption = "0"
    progressform.lblRemaining.Caption = CStr(totalCount)
    progressform.lblPercentage.Caption = "0%"
    progressform.lblElapsed.Caption = "0:00:00"
    progressform.lblETR.Caption = "--:--:--"
    
    CenterUserFormOnActiveMonitor progressform
    
    ' Capture the design-time width of the bar as the maximum; then collapse to zero
    maxBarWidth = lblProcessedBarFill.Width
    lblProcessedBarFill.Width = 0

    progressform.txtLog.Value = vbNullString
    progressform.txtLog.SelStart = 0

    Paused = False
    Cancelled = False
    progressform.btnPause.Caption = "Pause"

    startTick = Timer
    lastUpdate = startTick
    emaSecPerItem = 0#
End Sub

'— Append a time-stamped line to the log
Public Sub LogLine(ByVal lineText As String)
    ' Ensure progressform is initialized
    If Not progressform Is Nothing Then
        With progressform.txtLog
            If Len(.Text) > 0 Then
                .SelStart = Len(.Text)  ' Move to the end of the text
                .SelText = vbCrLf & Format$(Now, "hh:nn:ss") & "  " & lineText  ' Append the line
            Else
                .Text = Format$(Now, "hh:nn:ss") & "  " & lineText  ' Start with the first line
            End If
            .SelStart = Len(.Text)  ' Scroll to the end
        End With
    Else
        MsgBox "ProgressForm is not initialized!"
    End If
    DoEvents
End Sub

'— Update counters, percent, bar, elapsed, ETA. Call this once per record (or more).
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

'— Blocks while paused; returns False if cancelled while waiting
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
    LogLine "Cancel requested. Finishing current step…"
End Sub

Private Sub bSettings_Click()
    ToggleThisWorkbookVisibility
End Sub

'--- Drive Reflection > Session menu > OAIS2 with light retries ---
Private Sub InitializeReflectionAndOAIS()
    ' (1) Reflection Workspace Intro Screen?
    If WaitForText(1, 1, 79, TXT_DISA, retries, RETRY_WAIT_SEC) Then
        HitEnter ' pass the splash / login handoff

        ' (2) Session selection menu?
        If WaitForText(3, 1, 79, TXT_SESSION_MENU, retries, RETRY_WAIT_SEC) Then
            entText 23, 15, OAIS_MENU_CMD

            ' (3) Wait for OAIS banner, allow one "enter" nudge if needed
            If Not WaitForText(2, 1, 79, TXT_OAIS_BANNER, 2, RETRY_WAIT_SEC) Then
                SafePause 0.6
                HitEnter
                ' final check
                Call WaitForText(2, 1, 79, TXT_OAIS_BANNER, 2, RETRY_WAIT_SEC)
            End If
        End If
    End If

    ' Refresh status light after attempts
    SetOAISStatus Not (iCS Is Nothing)
End Sub

Private Sub SetOAISStatus(ByVal isConnected As Boolean)
    If isConnected Then
        lblOAIS.BackColor = vbGreen
        lblOAIS.Caption = "Connected to OAIS"
        lblOAIS.ForeColor = vbWhite
    Else
        lblOAIS.BackColor = vbRed
        lblOAIS.Caption = "OAIS Not Connected"
        lblOAIS.ForeColor = vbWhite
    End If
End Sub

Private Sub bOAIS_Click()
    ' If not connected, try to connect; else toggle external frame if present.
    If lblOAIS.BackColor = vbRed Then
        ConnectToRunningOAIS
        SetOAISStatus Not (iCS Is Nothing)
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

InitializeReflectionAndOAIS

Set progressform = New progressform

'A_Record_Review

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' X button behaves like Cancel to avoid orphaned background work
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        btnCancel_Click
    End If
End Sub

Private Sub SafePause(ByVal seconds As Double)
    Dim t As Single: t = Timer
    Do While Timer - t < seconds
        DoEvents
    Loop
End Sub

' Polls for substring on the Reflection screen text with retries.
Private Function WaitForText(ByVal row As Long, ByVal col As Long, ByVal nChars As Long, _
                             ByVal needle As String, ByVal retries As Long, ByVal waitSec As Double) As Boolean
    Dim i As Long, hay As String
    On Error Resume Next
    For i = 1 To retries
        hay = iCS.GetText(row, col, nChars)
        If InStr(1, hay, needle, vbTextCompare) > 0 Then
            WaitForText = True
            Exit Function
        End If
        SafePause waitSec
    Next i
    WaitForText = False
End Function
