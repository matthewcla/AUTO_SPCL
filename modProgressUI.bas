Attribute VB_Name = "modProgressUI"
Option Explicit

#If VBA7 Then
Public Declare PtrSafe Function SetTimer Lib "user32" ( _
    ByVal hWnd As LongPtr, _
    ByVal nIDEvent As LongPtr, _
    ByVal uElapse As Long, _
    ByVal lpTimerFunc As LongPtr) As LongPtr
Public Declare PtrSafe Function KillTimer Lib "user32" ( _
    ByVal hWnd As LongPtr, _
    ByVal uIDEvent As LongPtr) As Long
#Else
Public Declare Function SetTimer Lib "user32" ( _
    ByVal hWnd As Long, _
    ByVal nIDEvent As Long, _
    ByVal uElapse As Long, _
    ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" ( _
    ByVal hWnd As Long, _
    ByVal uIDEvent As Long) As Long
#End If

Public progressForm As progressForm
Public paused As Boolean
Public cancelled As Boolean

Private mProgressRunComplete As Boolean
Private mTotalCount As Long
Private mCompletedCount As Long

Public Function ProgressRunComplete() As Boolean
    ProgressRunComplete = mProgressRunComplete
End Function

Public Sub Progress_Show(ByVal totalCount As Long, Optional ByVal title As String = "")
    On Error GoTo HandleError

    Set progressForm = New ProgressForm
    cancelled = False
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

Public Sub ProgressForm_TimerTick()
    On Error Resume Next
    If Not progressForm Is Nothing Then
        progressForm.Tick_OneSecond
    End If
    On Error GoTo 0
End Sub

#If VBA7 Then
Public Sub ProgressForm_TimerProc(ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal idEvent As LongPtr, ByVal dwTime As Long)
#Else
Public Sub ProgressForm_TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
#End If
    On Error Resume Next
    ProgressForm_TimerTick
    On Error GoTo 0
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
        If Not keepOpen Then
            progressForm.ShutdownTimer
            Unload progressForm
            Set progressForm = Nothing
        End If
        On Error GoTo 0
    End If
End Sub
