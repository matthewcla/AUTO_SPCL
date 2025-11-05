Attribute VB_Name = "modEmailFormDiagnostics"
Option Explicit

Public Sub ReportMissingControls(ByVal missing As Collection)
    Dim message As String
    Dim entry As Variant

    message = "Missing EmailForm controls:" 

    If missing Is Nothing Then
        Debug.Print "[EmailFormDiagnostics] " & message & " <none provided>"
        Exit Sub
    End If

    If missing.Count = 0 Then
        Debug.Print "[EmailFormDiagnostics] " & message & " <none>"
        Exit Sub
    End If

    For Each entry In missing
        message = message & vbCrLf & " - " & CStr(entry)
    Next entry

    Debug.Print "[EmailFormDiagnostics] " & message
End Sub
