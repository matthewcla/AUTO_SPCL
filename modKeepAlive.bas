Attribute VB_Name = "modKeepAlive"
' ===========================
'  Reflections Keep-Alive (VBA)
'  - Nudge the session every N seconds (PER1 -> F3)
'  - Auto-suspend while user runs any priority proc
'  - Auto-resume when that proc completes
' ===========================

Option Explicit

' === CONFIG ===
Private Const KEEPALIVE_PASSWORD As String = "666826"
Private Const LOCKED_OUT As String = "unlock word"
Private Const KEEPALIVE_SECONDS As Long = 30 '120   ' every 2 minutes
Private Const ENTER_DELAY_MS     As Long = 250  ' small delay between actions

' === STATE ===
Private mNextFire      As Date
Private mEnabled       As Boolean
Private mPaused        As Boolean
Private mInTick        As Boolean
Private mStarted       As Boolean

' ========= Public API =========

' Start the background keep-alive loop (call once, e.g., in Workbook_Open)
Public Sub KeepAlive_Start()
    mEnabled = True
    mPaused = False
    mStarted = True
    KeepAlive_ScheduleNext
End Sub

' Stop the background loop entirely (won’t auto-resume)
Public Sub KeepAlive_Stop()
    On Error Resume Next
    If mNextFire <> 0 Then
        Application.OnTime mNextFire, "modKeepAlive.KeepAlive_Tick", , False
    End If
    mEnabled = False
    mPaused = False
    mInTick = False
    mNextFire = 0
End Sub

' Suspend the loop during a priority task; call KeepAlive_Resume afterward
Public Sub KeepAlive_Suspend()
    mPaused = True
    KeepAlive_CancelNext
End Sub

' Resume after a priority task completes (only if previously started)
Public Sub KeepAlive_Resume()
    If Not mStarted Then Exit Sub
    mPaused = False
    If mEnabled Then KeepAlive_ScheduleNext
End Sub

' Convenience wrapper to run any procedure "as priority"
' Example: RunPriority "DoBigSync" or RunPriority("DoThingWithArgs", arg1, arg2)
Public Sub RunPriority(ByVal procName As String, ParamArray args() As Variant)
    Dim capturedErrNumber As Long
    Dim capturedErrDescription As String

    On Error GoTo HandleError
    KeepAlive_Suspend
    Select Case UBound(args)
        Case -1: Application.Run procName
        Case 0:  Application.Run procName, args(0)
        Case 1:  Application.Run procName, args(0), args(1)
        Case 2:  Application.Run procName, args(0), args(1), args(2)
        Case 3:  Application.Run procName, args(0), args(1), args(2), args(3)
        Case 4:  Application.Run procName, args(0), args(1), args(2), args(3), args(4)
        Case Else
            ' Add more cases if you need >5 args
            Application.Run procName, args(0), args(1), args(2), args(3), args(4)
    End Select
    KeepAlive_Resume
    Exit Sub

HandleError:
    capturedErrNumber = Err.Number
    capturedErrDescription = Err.Description
    KeepAlive_Resume
    If capturedErrNumber <> 0 Then
        Err.Clear
        Err.Raise capturedErrNumber, , capturedErrDescription
    End If
End Sub

' ========= Internals =========

Private Sub KeepAlive_ScheduleNext()
    If Not mEnabled Or mPaused Then Exit Sub
    KeepAlive_CancelNext
    mNextFire = Now + TimeSerial(0, 0, KEEPALIVE_SECONDS)
    On Error Resume Next
    Application.OnTime mNextFire, "modKeepAlive.KeepAlive_Tick", , True
End Sub

Private Sub KeepAlive_CancelNext()
    On Error Resume Next
    If mNextFire <> 0 Then
        Application.OnTime mNextFire, "modKeepAlive.KeepAlive_Tick", , False
        mNextFire = 0
    End If
End Sub

' The heartbeat that performs the nudge, then reschedules itself
Public Sub KeepAlive_Tick()
    On Error GoTo Resched
    If Not mEnabled Or mPaused Then GoTo Resched
    If mInTick Then GoTo Resched              ' re-entrancy guard
    mInTick = True

    ' --- Do the minimal "nudge": PER1 then F3 back to menu ---
    SafeNudge

Resched:
    mInTick = False
    KeepAlive_ScheduleNext
End Sub

' Encapsulate your two operations with small delay & error swallow
Private Sub SafeNudge()
    On Error Resume Next
    If InStr(iCS.GetText(11, 1, 79), LOCKED_OUT) > 0 Then entText 11, 36, KEEPALIVE_PASSWORD
    entText 19, 11, "PER1"         ' go to PER1
    TinyDelay ENTER_DELAY_MS
    HitF3                           ' back to menu
    On Error GoTo 0
End Sub

' Simple millisecond delay without freezing Excel
Private Sub TinyDelay(ByVal ms As Long)
    Dim t As Single: t = Timer
    Do While (Timer - t) * 1000 < ms
        DoEvents
    Loop
End Sub

' ========= Optional boot hook =========
' Call this from ThisWorkbook.Workbook_Open to auto-start:
' Private Sub Workbook_Open()
'     modKeepAlive.KeepAlive_Start
' End Sub

' ========= Usage Patterns =========
' 1) Start once (e.g., on workbook open):
'       KeepAlive_Start
'
' 2) Wrap any user-invoked/priority macro:
'       Sub RunMyBigJob()
'           RunPriority "MyBigJob"            ' or RunPriority "MyJobWithArgs", arg1, arg2
'       End Sub
'
'       Sub MyBigJob()
'           ' ... your long-running work ...
'       End Sub
'
'    (Keep-alive suspends before MyBigJob runs and resumes afterward.)
'
' 3) Manually control if you prefer:
'       KeepAlive_Suspend
'       ' do stuff...
'       KeepAlive_Resume
'
' 4) Stop entirely (e.g., when closing):
'       KeepAlive_Stop


