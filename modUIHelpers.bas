Attribute VB_Name = "modUIHelpers"
Option Explicit

Private mWaitCursorDepth As Long

Public Sub SetCursorWait()
    On Error Resume Next
    mWaitCursorDepth = mWaitCursorDepth + 1
    If mWaitCursorDepth = 1 Then
        Application.Cursor = xlWait
    End If
    On Error GoTo 0
End Sub

Public Sub SetCursorDefault()
    On Error Resume Next
    If mWaitCursorDepth > 0 Then
        mWaitCursorDepth = mWaitCursorDepth - 1
    End If
    If mWaitCursorDepth = 0 Then
        Application.Cursor = xlDefault
    End If
    On Error GoTo 0
End Sub
