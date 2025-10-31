Attribute VB_Name = "modReflectionsMonitor"
Option Explicit

' ==============================================================
'  Module: modReflectionsMonitor
'  Purpose: Periodically ensure the Reflections session is alive
'           and notify interested forms without blocking the UI.
' ==============================================================

Private Const CHECK_INTERVAL_SECONDS As Long = 30

Private mNextCheck As Date
Private mRegisteredForms As Collection
Private mLastKnownStatus As Variant
Private mLossMessageShown As Boolean

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

    Dim isConnected As Boolean
    isConnected = EnsureReflectionsConnectionAlive(True)

    NotifySpecific key, isConnected

    mLastKnownStatus = isConnected
    If isConnected Then
        mLossMessageShown = False
    Else
        ShowConnectionLostMessage
        mLossMessageShown = True
    End If

    EnsureTimerScheduled
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

    If mRegisteredForms.Count = 0 Then
        CancelScheduledTick
        mLastKnownStatus = Empty
        mLossMessageShown = False
    End If
End Sub

' Timer callback executed by Application.OnTime
Public Sub ReflectionsMonitor_Tick()
    On Error GoTo CleanExit

    mNextCheck = 0

    If Not HasRegisteredForms Then GoTo CleanExit

    Dim isConnected As Boolean
    isConnected = EnsureReflectionsConnectionAlive(True)

    If isConnected Then
        mLossMessageShown = False
    Else
        If Not mLossMessageShown Then
            ShowConnectionLostMessage
            mLossMessageShown = True
        End If
    End If

    mLastKnownStatus = isConnected
    NotifyAll isConnected

CleanExit:
    If HasRegisteredForms Then
        ScheduleNextTick
    Else
        CancelScheduledTick
    End If
End Sub

' Force an immediate status push to listeners (if any)
Public Sub PushCurrentStatus()
    If Not HasRegisteredForms Then Exit Sub
    Dim isConnected As Boolean
    isConnected = EnsureReflectionsConnectionAlive(True)
    mLastKnownStatus = isConnected
    NotifyAll isConnected
End Sub

' ===== Helper routines =====

Private Function HasRegisteredForms() As Boolean
    HasRegisteredForms = Not (mRegisteredForms Is Nothing) And mRegisteredForms.Count > 0
End Function

Private Sub EnsureTimerScheduled()
    If HasRegisteredForms Then
        If mNextCheck = 0 Then
            ScheduleNextTick
        End If
    End If
End Sub

Private Sub ScheduleNextTick()
    CancelScheduledTick
    mNextCheck = Now + TimeSerial(0, 0, CHECK_INTERVAL_SECONDS)
    On Error Resume Next
    Application.OnTime mNextCheck, "modReflectionsMonitor.ReflectionsMonitor_Tick", , True
    On Error GoTo 0
End Sub

Private Sub CancelScheduledTick()
    On Error Resume Next
    If mNextCheck <> 0 Then
        Application.OnTime mNextCheck, "modReflectionsMonitor.ReflectionsMonitor_Tick", , False
        mNextCheck = 0
    End If
    On Error GoTo 0
End Sub

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
