Attribute VB_Name = "modKeepAlive"
' ===========================
'  Reflections Keep-Alive & Monitor (VBA)
'  - Nudge the session every N seconds (PER1 -> F3)
'  - Auto-suspend while user runs any priority proc
'  - Auto-resume when that proc completes
'  - Surface connection status changes to interested forms
' ===========================

Option Explicit

' === CONFIG ===
Private Const KEEPALIVE_PASSWORD As String = "666826"
Private Const LOCKED_OUT As String = "unlock word"
Private Const KEEPALIVE_SECONDS As Long = 30 '120   ' every 2 minutes
Private Const ENTER_DELAY_MS     As Long = 250  ' small delay between actions

' === STATE ===
Private mNextFire          As Date
Private mEnabled           As Boolean
Private mPaused            As Boolean
Private mInTick            As Boolean
Private mStarted           As Boolean
Private mRegisteredForms   As Collection
Private mLastKnownStatus   As Variant
Private mLossMessageShown  As Boolean

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
    ResetConnectionState
    Set mRegisteredForms = Nothing
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

' Register a form (by type/name) for periodic connection updates
Public Sub RegisterReflectionsListener(ByVal formName As String)
    Dim key As String
    key = NormalizeKey(formName)
    If Len(key) = 0 Then Exit Sub

    If mRegisteredForms Is Nothing Then
        Set mRegisteredForms = New Collection
    End If

    On Error Resume Next
    mRegisteredForms.Add key, key
    On Error GoTo 0

    RefreshConnectionState True, key
End Sub

' Remove a form from monitoring
Public Sub UnregisterReflectionsListener(ByVal formName As String)
    If mRegisteredForms Is Nothing Then Exit Sub

    Dim key As String
    key = NormalizeKey(formName)
    If Len(key) = 0 Then Exit Sub

    On Error Resume Next
    mRegisteredForms.Remove key
    On Error GoTo 0

    If Not HasRegisteredForms Then
        ResetConnectionState
    End If
End Sub

' Force an immediate status push to listeners (if any)
Public Sub PushCurrentStatus()
    If Not HasRegisteredForms Then Exit Sub
    RefreshConnectionState False, "", True
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

    Dim isConnected As Boolean
    isConnected = RefreshConnectionState(True)

    If isConnected Then
        ' --- Do the minimal "nudge": PER1 then F3 back to menu ---
        SafeNudge
    End If

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

' ===== Connection monitoring helpers =====

Private Function RefreshConnectionState(Optional ByVal showLossMessage As Boolean = True, _
                                        Optional ByVal notifyKey As String = "", _
                                        Optional ByVal forceNotify As Boolean = False) As Boolean
    Dim isConnected As Boolean
    isConnected = EnsureReflectionsConnectionAlive(True)

    Dim haveListeners As Boolean
    haveListeners = HasRegisteredForms

    If isConnected Then
        mLossMessageShown = False
    ElseIf showLossMessage And (Len(notifyKey) > 0 Or haveListeners) Then
        If Not mLossMessageShown Then
            ShowConnectionLostMessage
            mLossMessageShown = True
        End If
    End If

    If Len(notifyKey) > 0 Then
        NotifySpecific notifyKey, isConnected
    ElseIf haveListeners Then
        Dim statusChanged As Boolean
        If VarType(mLastKnownStatus) = vbBoolean Then
            statusChanged = (CBool(mLastKnownStatus) <> isConnected)
        Else
            statusChanged = True
        End If

        If forceNotify Or statusChanged Then
            NotifyAll isConnected
        End If
    End If

    mLastKnownStatus = isConnected
    RefreshConnectionState = isConnected
End Function

Private Sub ResetConnectionState()
    mLastKnownStatus = Empty
    mLossMessageShown = False
End Sub

Private Function HasRegisteredForms() As Boolean
    HasRegisteredForms = Not (mRegisteredForms Is Nothing) And mRegisteredForms.Count > 0
End Function

Private Sub NotifyAll(ByVal isConnected As Boolean)
    Dim formObj As Object
    Dim key As String

    For Each formObj In VBA.UserForms
        key = NormalizeKey(TypeName(formObj))
        If ListenerRegistered(key) Then
            On Error Resume Next
            CallByName formObj, "HandleReflectionsConnection", VbMethod, isConnected
            On Error GoTo 0
        End If
    Next formObj
End Sub

Private Sub NotifySpecific(ByVal key As String, ByVal isConnected As Boolean)
    Dim formObj As Object
    For Each formObj In VBA.UserForms
        If NormalizeKey(TypeName(formObj)) = key Then
            On Error Resume Next
            CallByName formObj, "HandleReflectionsConnection", VbMethod, isConnected
            On Error GoTo 0
            Exit For
        End If
    Next formObj
End Sub

Private Function ListenerRegistered(ByVal key As String) As Boolean
    On Error GoTo NotFound
    Dim tmp As String
    tmp = mRegisteredForms(key)
    ListenerRegistered = True
    Exit Function
NotFound:
    ListenerRegistered = False
    Err.Clear
End Function

Private Function NormalizeKey(ByVal formName As String) As String
    NormalizeKey = UCase$(Trim$(formName))
End Function

Private Sub ShowConnectionLostMessage()
    Dim prompt As String
    prompt = "Excel is unable to detect an active Reflections/OAIS session." & vbCrLf & _
             "Please make sure the Reflections window is running and connected." & vbCrLf & _
             "If the session is restored, the form will reconnect automatically."
    modUIHelpers.ShowWarningMessage prompt, "Reflections Connection"
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


