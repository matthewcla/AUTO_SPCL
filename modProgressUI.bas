Attribute VB_Name = "modProgressUI"
Option Explicit

' Public constants used to control progress logging
Public Const PROGRESS_LOG_STARTED As String = "Record Review Started"
Public Const PROGRESS_LOG_CONCLUDED As String = "Record Review Concluded"

Public progressForm As progressForm
Public paused As Boolean
Public cancelled As Boolean
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
    Set progressForm = New ProgressForm
    cancelled = False
    CurrentRecordName = vbNullString
    CurrentRecordSSN = vbNullString
    mTotalCount = totalCount
    mCompletedCount = 0
    mProgressRunComplete = (mTotalCount > 0 And mCompletedCount >= mTotalCount)
    progressForm.Show vbModeless
    progressForm.Init totalCount, title
    Exit Sub

HandleError:
    Progress_Close
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub Progress_Log(ByVal msg As String)
    If Not progressForm Is Nothing Then
        progressForm.LogLine msg
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
    If progressForm Is Nothing Then GoTo CleanExit

    progressForm.Tick_OneSecond

CleanExit:
    mInTick = False
    If mTimerEnabled Then
        Progress_ScheduleNextTick
    End If
End Sub

Public Sub ProgressForm_TimerTick()
    On Error Resume Next
    If Not progressForm Is Nothing Then
        progressForm.Tick_OneSecond
    End If

    mTimerScheduled = False
    mNextTick = 0
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

Public Sub Progress_Update(ByVal done As Long, ByVal totalCount As Long, Optional ByVal status As String = "")
    If Not progressForm Is Nothing Then
        progressForm.UpdateProgress done, totalCount, status
    End If

    mCompletedCount = done
    mTotalCount = totalCount
    mProgressRunComplete = (mTotalCount > 0 And mCompletedCount >= mTotalCount)
End Sub

Public Function Progress_WaitIfPaused() As Boolean
    If Not progressForm Is Nothing Then
        Progress_WaitIfPaused = progressForm.WaitIfPaused
    Else
        Progress_WaitIfPaused = False
    End If
End Function

Public Function Progress_Cancelled() As Boolean
    If progressForm Is Nothing Then
        Progress_Cancelled = cancelled
        Exit Function
    End If

    On Error Resume Next
    Progress_Cancelled = progressForm.Cancelled
    If Err.Number <> 0 Then
        Err.Clear
        Progress_Cancelled = cancelled
    Else
        cancelled = Progress_Cancelled
    End If
    On Error GoTo 0
End Function

Public Sub Progress_Close(Optional ByVal finalNote As String = "", Optional ByVal keepOpen As Boolean = False)
    Progress_StopTimer

    If Not progressForm Is Nothing Then
        On Error Resume Next
        mCompletedCount = progressForm.CompletedCount
        mTotalCount = progressForm.TotalCount
        mProgressRunComplete = progressForm.ProgressComplete
        If Err.Number <> 0 Then
            Err.Clear
            mProgressRunComplete = (mTotalCount > 0 And mCompletedCount >= mTotalCount)
        End If
        On Error GoTo 0
    Else
        mProgressRunComplete = (mTotalCount > 0 And mCompletedCount >= mTotalCount)
    End If

    If cancelled Then
        mProgressRunComplete = True
    End If

    If Not progressForm Is Nothing Then
        On Error Resume Next
        If Len(finalNote) > 0 Then
            progressForm.LogLine finalNote
        End If
        progressForm.ShutdownTimer
        If Not keepOpen Then
            Unload progressForm
            Set progressForm = Nothing
            Progress_ResetTimerState
        End If
        On Error GoTo 0
    End If

    CurrentRecordName = vbNullString
    CurrentRecordSSN = vbNullString
End Sub
