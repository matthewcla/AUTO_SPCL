Attribute VB_Name = "modProgressUI"
Option Explicit


Public progressform As ProgressForm ' Strongly typed reference for compile-time safety

Public Sub Progress_Show(ByVal totalCount As Long, Optional ByVal title As String = "Reviewing records")
    If Not progressform Is Nothing Then
        Unload progressform
        Set progressform = Nothing
    End If

    Set progressform = New ProgressForm
    progressform.Show vbModeless
    progressform.Init totalCount, title
End Sub

Public Sub Progress_Log(ByVal msg As String)
    If Not progressform Is Nothing Then
        progressform.LogLine msg
    End If
End Sub

Public Sub Progress_Update(ByVal done As Long, ByVal totalCount As Long, Optional ByVal status As String = "")
    If Not progressform Is Nothing Then
        progressform.UpdateProgress done, totalCount, status
    End If
End Sub

Public Function Progress_WaitIfPaused() As Boolean
    If Not progressform Is Nothing Then
        Progress_WaitIfPaused = progressform.WaitIfPaused
    Else
        Progress_WaitIfPaused = True
    End If
End Function

Public Function Progress_Cancelled() As Boolean
    If Not progressform Is Nothing Then
        Progress_Cancelled = progressform.Cancelled
    Else
        Progress_Cancelled = False
    End If
End Function

Public Sub Progress_Close(Optional ByVal finalNote As String = "")
    If Not progressform Is Nothing Then
        If Len(finalNote) > 0 Then
            progressform.LogLine finalNote
        End If
        Unload progressform
        Set progressform = Nothing
    End If
End Sub
