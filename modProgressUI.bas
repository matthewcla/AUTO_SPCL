Attribute VB_Name = "modProgressUI"
Option Explicit

Public progressForm As progressForm
Public paused As Boolean
Public cancelled As Boolean

Private mProgressRunComplete As Boolean

Public Property Get ProgressRunComplete() As Boolean
    ProgressRunComplete = mProgressRunComplete
End Property

Public Property Let ProgressRunComplete(ByVal value As Boolean)
    mProgressRunComplete = value
End Property

Public Sub Progress_Show(ByVal totalCount As Long, Optional ByVal title As String = "")
    On Error GoTo HandleError

    Set progressForm = New ProgressForm
    cancelled = False
    ProgressRunComplete = False
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

Public Sub Progress_Update(ByVal done As Long, ByVal totalCount As Long, Optional ByVal status As String = "")
    If Not progressForm Is Nothing Then
        progressForm.UpdateProgress done, totalCount, status
    End If
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
    ProgressRunComplete = True

    If Not progressForm Is Nothing Then
        On Error Resume Next
        If Len(finalNote) > 0 Then
            progressForm.LogLine finalNote
        End If
        If Not keepOpen Then
            Unload progressForm
            Set progressForm = Nothing
        End If
        On Error GoTo 0
    End If
End Sub
