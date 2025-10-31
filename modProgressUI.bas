Attribute VB_Name = "modProgressUI"
Option Explicit

' Public constants used to control progress logging
Public Const PROGRESS_LOG_STARTED As String = "Record Review Started"
Public Const PROGRESS_LOG_CONCLUDED As String = "Record Review Concluded"

Public ProgressFormView As ProgressForm
Public IsCancellationRequested As Boolean
Public CurrentRecordName As String
Public CurrentRecordSSN As String

Private mProgressRunComplete As Boolean
Private mTotalCount As Long
Private mCompletedCount As Long
Private mInTick As Boolean
Private Const PROGRESS_TIMER_INTERVAL As Double = 1# / (24# * 60# * 60#)
Private mTimerEnabled As Boolean
Private mTimerScheduled As Boolean
Private mNextTick As Date

Public Sub UpdateProgressButtonStates(ByRef btnCancel As MSForms.CommandButton, _
                                      ByRef btnPause As MSForms.CommandButton, _
                                      ByVal isCancellationRequested As Boolean, _
                                      ByVal isComplete As Boolean)

    If btnCancel Is Nothing Then Exit Sub

    If isCancellationRequested Then
        ' Lock the UI while cancellation completes so the user cannot trigger duplicate requests.
        SetButtonCaptionIfDifferent btnCancel, "Cancelling..."
        modUIHelpers.SetControlsEnabled btnCancel, False
        modUIHelpers.SetControlsVisible btnPause, False
        Exit Sub
    End If

    modUIHelpers.SetControlsEnabled btnCancel, True

    If isComplete Then
        ' Once work is finished, convert the cancel button into "Next" to advance flows (e.g. EmailForm).
        SetButtonCaptionIfDifferent btnCancel, "Next"
        modUIHelpers.SetControlsVisible btnPause, False
    Else
        ' Keep pause visible only during active processing to avoid confusing idle users.
        SetButtonCaptionIfDifferent btnCancel, "Cancel"
        modUIHelpers.SetControlsVisible btnPause, True
    End If
End Sub

Private Sub SetButtonCaptionIfDifferent(ByRef target As MSForms.CommandButton, _
                                        ByVal captionText As String)
    If target Is Nothing Then Exit Sub

    On Error Resume Next
    If StrComp(CStr(target.Caption), captionText, vbBinaryCompare) <> 0 Then
        target.Caption = captionText
    End If
    On Error GoTo 0
End Sub

Public Function IsFormLoaded(ByVal formName As String) As Boolean
    Dim frm As Object

    For Each frm In VBA.UserForms
        If StrComp(frm.Name, formName, vbTextCompare) = 0 Then
            IsFormLoaded = True
            Exit Function
        End If
    Next frm
End Function

Public Function ProgressRunComplete() As Boolean
    ProgressRunComplete = mProgressRunComplete
End Function

Public Sub Progress_Show(ByVal totalCount As Long, Optional ByVal title As String = "")
    On Error GoTo HandleError

    Progress_ResetTimerState
    Set ProgressFormView = New ProgressForm
    IsCancellationRequested = False
    CurrentRecordName = vbNullString
    CurrentRecordSSN = vbNullString
    mTotalCount = totalCount
    mCompletedCount = 0
    mProgressRunComplete = (mTotalCount > 0 And mCompletedCount >= mTotalCount)
    ProgressFormView.Show vbModeless
    ProgressFormView.Init totalCount, title
    Exit Sub

HandleError:
    Progress_Close
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub Progress_Log(ByVal msg As String)
    If Not ProgressFormView Is Nothing Then
        ProgressFormView.LogLine msg
    End If
End Sub

Public Sub LogRecordReviewStart(ByVal recordName As String, ByVal recordSSN As String)
    CurrentRecordName = Trim$(recordName)
    CurrentRecordSSN = Trim$(recordSSN)
    Progress_Log FormatRecordReviewMessage("Start", CurrentRecordName, CurrentRecordSSN)
End Sub

Public Sub LogRecordReviewResume()
    If LenB(CurrentRecordName) = 0 And LenB(CurrentRecordSSN) = 0 Then Exit Sub
    Progress_Log FormatRecordReviewMessage("Start", CurrentRecordName, CurrentRecordSSN, "Resumed")
End Sub

Public Sub LogRecordReviewPaused()
    If LenB(CurrentRecordName) = 0 And LenB(CurrentRecordSSN) = 0 Then Exit Sub
    Progress_Log FormatRecordReviewMessage("End", CurrentRecordName, CurrentRecordSSN, "Paused")
End Sub

Public Sub LogRecordReviewCancelled()
    If LenB(CurrentRecordName) = 0 And LenB(CurrentRecordSSN) = 0 Then Exit Sub
    Progress_Log FormatRecordReviewMessage("End", CurrentRecordName, CurrentRecordSSN, "Cancelled")
    CurrentRecordName = vbNullString
    CurrentRecordSSN = vbNullString
End Sub

Public Sub LogRecordReviewCompleted()
    If LenB(CurrentRecordName) = 0 And LenB(CurrentRecordSSN) = 0 Then Exit Sub
    Progress_Log FormatRecordReviewMessage("End", CurrentRecordName, CurrentRecordSSN, "Completed")
    CurrentRecordName = vbNullString
    CurrentRecordSSN = vbNullString
End Sub

Public Sub LogRecordReviewStatus(ByVal statusText As String)
    Progress_Log FormatRecordReviewMessage("Status", CurrentRecordName, CurrentRecordSSN, statusText)
End Sub

Private Function FormatRecordReviewMessage( _
    ByVal actionText As String, _
    ByVal recordName As String, _
    ByVal recordSSN As String, _
    Optional ByVal statusText As String = "") As String

    Dim timestamp As String
    Dim formattedAction As String
    Dim suffix As String

    recordName = Trim$(recordName)
    recordSSN = Trim$(recordSSN)

    If Len(recordName) = 0 Then
        recordName = "<unknown>"
    End If

    If Len(recordSSN) = 0 Then
        recordSSN = "<unknown>"
    End If

    timestamp = Format$(Now, "hh:nn:ss")

    Select Case LCase$(actionText)
        Case "start"
            formattedAction = "Start"
        Case "end"
            formattedAction = "End"
        Case "status"
            formattedAction = "Status"
        Case Else
            formattedAction = actionText
    End Select

    If Len(statusText) > 0 Then
        If LCase$(formattedAction) = "status" Then
            suffix = " - " & statusText
        Else
            suffix = " (" & statusText & ")"
        End If
    Else
        suffix = ""
    End If

    FormatRecordReviewMessage = "[Time: " & timestamp & "] Record Review " & formattedAction & suffix & ": " & recordName & " | SSN: " & recordSSN
End Function

Public Sub Progress_TimerTick()
    mTimerScheduled = False
    mNextTick = 0

    If Not mTimerEnabled Then Exit Sub
    If mInTick Then
        Progress_ScheduleNextTick
        Exit Sub
    End If

    mInTick = True
    On Error GoTo CleanExit

    If Not IsFormLoaded("ProgressForm") Then GoTo CleanExit
    If ProgressFormView Is Nothing Then GoTo CleanExit

    ProgressFormView.Tick_OneSecond

CleanExit:
    mInTick = False
    If mTimerEnabled Then
        Progress_ScheduleNextTick
    End If
End Sub

Public Sub ProgressForm_TimerTick()
    On Error Resume Next
    If Not ProgressFormView Is Nothing Then
        ProgressFormView.Tick_OneSecond
    End If

    mTimerScheduled = False
    mNextTick = 0
End Sub

Public Sub Progress_Pulse()
    If ProgressFormView Is Nothing Then Exit Sub
    If Not IsFormLoaded("ProgressForm") Then Exit Sub
    If mInTick Then Exit Sub

    mInTick = True
    On Error GoTo CleanExit

    ProgressFormView.Tick_OneSecond

CleanExit:
    mInTick = False
End Sub

Public Sub Progress_StartTimer()
    mTimerEnabled = True
    mInTick = False

    If Not mTimerScheduled Then
        Progress_ScheduleNextTick
    End If
End Sub

Public Sub Progress_StopTimer()
    Dim scheduledTick As Date

    If mTimerScheduled Then
        scheduledTick = mNextTick
        If scheduledTick = 0 Then
            scheduledTick = Now
        End If
        On Error Resume Next
        Application.OnTime EarliestTime:=scheduledTick, Procedure:="modProgressUI.Progress_TimerTick", Schedule:=False
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
    End If

    mTimerEnabled = False
    mTimerScheduled = False
    mNextTick = 0
    mInTick = False
End Sub

' Reset any pending timer callbacks and mark the timer as disabled.
Public Sub Progress_ResetTimerState()
    Progress_StopTimer
End Sub

Private Sub Progress_ScheduleNextTick(Optional ByVal referenceTime As Date)
    If Not mTimerEnabled Then Exit Sub

    If referenceTime = 0 Then
        referenceTime = Now
    End If

    mNextTick = referenceTime + PROGRESS_TIMER_INTERVAL

    On Error GoTo FailedSchedule
    Application.OnTime EarliestTime:=mNextTick, Procedure:="modProgressUI.Progress_TimerTick", Schedule:=True
    mTimerScheduled = True
    Exit Sub

FailedSchedule:
    mTimerScheduled = False
    mTimerEnabled = False
    mNextTick = 0
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub Progress_Update(ByVal recordsCompleted As Long, _
                          ByVal recordsTotal As Long, _
                          Optional ByVal statusMessage As String = "")
    If Not ProgressFormView Is Nothing Then
        ProgressFormView.UpdateProgress recordsCompleted, recordsTotal, statusMessage
    End If

    mCompletedCount = recordsCompleted
    mTotalCount = recordsTotal
    mProgressRunComplete = (mTotalCount > 0 And mCompletedCount >= mTotalCount)
End Sub

Public Function Progress_WaitIfPaused() As Boolean
    If Not ProgressFormView Is Nothing Then
        Progress_WaitIfPaused = ProgressFormView.WaitIfPaused
    Else
        Progress_WaitIfPaused = False
    End If
End Function

Public Function Progress_Cancelled() As Boolean
    If ProgressFormView Is Nothing Then
        Progress_Cancelled = IsCancellationRequested
        Exit Function
    End If

    On Error Resume Next
    Progress_Cancelled = ProgressFormView.IsCancelled
    If Err.Number <> 0 Then
        Err.Clear
        Progress_Cancelled = IsCancellationRequested
    Else
        IsCancellationRequested = Progress_Cancelled
    End If
    On Error GoTo 0
End Function

Public Sub Progress_Close(Optional ByVal finalNote As String = "", Optional ByVal keepOpen As Boolean = False)
    Progress_StopTimer

    If Not ProgressFormView Is Nothing Then
        On Error Resume Next
        mCompletedCount = ProgressFormView.CompletedCount
        mTotalCount = ProgressFormView.TotalCount
        mProgressRunComplete = ProgressFormView.ProgressComplete
        If Err.Number <> 0 Then
            Err.Clear
            mProgressRunComplete = (mTotalCount > 0 And mCompletedCount >= mTotalCount)
        End If
        On Error GoTo 0
    Else
        mProgressRunComplete = (mTotalCount > 0 And mCompletedCount >= mTotalCount)
    End If

    If IsCancellationRequested Then
        mProgressRunComplete = True
    End If

    If Not ProgressFormView Is Nothing Then
        On Error Resume Next
        If Len(finalNote) > 0 Then
            ProgressFormView.LogLine finalNote
        End If
        ProgressFormView.ShutdownTimer
        If Not keepOpen Then
            Unload ProgressFormView
            Set ProgressFormView = Nothing
            Progress_ResetTimerState
        End If
        On Error GoTo 0
    End If

    CurrentRecordName = vbNullString
    CurrentRecordSSN = vbNullString
End Sub
