Attribute VB_Name = "modProgressUI"
Option Explicit

Public progressForm As ProgressForm

Public Sub Progress_Show(ByVal totalCount As Long, Optional ByVal title As String = "")
    On Error GoTo HandleError

    Set progressForm = New ProgressForm
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
    If Not progressForm Is Nothing Then
        Progress_Cancelled = progressForm.Cancelled
    Else
        Progress_Cancelled = False
    End If
End Function

Public Sub Progress_Close(Optional ByVal finalNote As String = "")
    If Not progressForm Is Nothing Then
        On Error Resume Next
        If Len(finalNote) > 0 Then
            progressForm.LogLine finalNote
        End If
        Unload progressForm
        Set progressForm = Nothing
        On Error GoTo 0
    End If
End Sub
