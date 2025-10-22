Attribute VB_Name = "modProgressUI"
Option Explicit


Public progressform As Object ' Declare as the actual UserForm type

' Safe wrappers so other modules don’t need to know the form’s internals.

Public Sub Progress_Show(ByVal totalCount As Long, Optional ByVal title As String = "Record Review")
    On Error Resume Next
    If progressform Is Nothing Then
        ' no-op - compiled reference
    End If
    On Error GoTo 0
    
    progressform.Show vbModeless
    progressform.Init totalCount, title
End Sub

Public Sub Progress_Log(ByVal msg As String)
    If Not progressform Is Nothing Then progressform.LogLine msg
End Sub

Public Sub Progress_Update(ByVal done As Long, ByVal totalCount As Long, Optional ByVal status As String = "")
    If progressform Is Nothing Then progressform.UpdateProgress done, totalCount, status
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
        Progress_Cancelled = False
    Else
        Progress_Cancelled = progressform.Cancelled
    End If
End Function

Public Sub Progress_Close(Optional ByVal finalNote As String = "")
    If Not progressform Is Nothing Then
        If Len(finalNote) > 0 Then progressform.LogLine finalNote
        Unload progressform ' Unload the progressform to free memory
        Set progressform = Nothing ' Clean up the reference
    End If
End Sub

