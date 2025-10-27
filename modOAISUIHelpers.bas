Attribute VB_Name = "modOAISUIHelpers"
Option Explicit

' --- Constants for screen-text detection (Reflection / OAIS banners) ---
Public Const TXT_DISA As String = "Defense Information Systems Agency"
Public Const TXT_SESSION_MENU As String = "CL/SuperSession"
Public Const TXT_OAIS_BANNER As String = "Officer Assignment Information System"
Public Const OAIS_MENU_CMD As String = "Start OAIS2"

' --- Tuning knobs ---
Private Const RETRY_WAIT_SEC As Double = 1.2
Private Const RETRIES As Long = 3

' Drive Reflection > Session menu > OAIS2 with light retries
Public Sub InitializeOAISSession(ByVal statusControl As Object, _
                                 Optional ByVal connectedText As String = "Connected to OAIS", _
                                 Optional ByVal disconnectedText As String = "OAIS Not Connected", _
                                 Optional ByVal connectedForeColor As Variant, _
                                 Optional ByVal disconnectedForeColor As Variant)
    If WaitForText(1, 1, 79, TXT_DISA, RETRIES, RETRY_WAIT_SEC) Then
        HitEnter ' pass the splash / login handoff

        If WaitForText(3, 1, 79, TXT_SESSION_MENU, RETRIES, RETRY_WAIT_SEC) Then
            entText 23, 15, OAIS_MENU_CMD

            If Not WaitForText(2, 1, 79, TXT_OAIS_BANNER, 2, RETRY_WAIT_SEC) Then
                SafePause 0.6
                HitEnter
                Call WaitForText(2, 1, 79, TXT_OAIS_BANNER, 2, RETRY_WAIT_SEC)
            End If
        End If
    End If

    SetOAISStatus statusControl, Not (iCS Is Nothing), connectedText, disconnectedText, connectedForeColor, disconnectedForeColor
End Sub

Public Sub SetOAISStatus(ByVal statusControl As Object, ByVal isConnected As Boolean, _
                         Optional ByVal connectedText As String = "Connected to OAIS", _
                         Optional ByVal disconnectedText As String = "OAIS Not Connected", _
                         Optional ByVal connectedForeColor As Variant, _
                         Optional ByVal disconnectedForeColor As Variant)
    If statusControl Is Nothing Then Exit Sub

    Dim controlName As String
    On Error Resume Next
    controlName = UCase$(CStr(CallByName(statusControl, "Name", VbGet)))
    On Error GoTo 0

    Dim effectiveConnectedText As String
    Dim effectiveDisconnectedText As String
    effectiveConnectedText = connectedText
    effectiveDisconnectedText = disconnectedText

    If controlName = "LBLOAIS" Then
        effectiveConnectedText = ""
        effectiveDisconnectedText = ""
    End If

    On Error Resume Next
    If isConnected Then
        CallByName statusControl, "Caption", VbLet, effectiveConnectedText
        CallByName statusControl, "BackColor", VbLet, vbGreen
        If Not IsMissing(connectedForeColor) Then
            CallByName statusControl, "ForeColor", VbLet, connectedForeColor
        End If
    Else
        CallByName statusControl, "Caption", VbLet, effectiveDisconnectedText
        CallByName statusControl, "BackColor", VbLet, vbRed
        If Not IsMissing(disconnectedForeColor) Then
            CallByName statusControl, "ForeColor", VbLet, disconnectedForeColor
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub SafePause(ByVal seconds As Double)
    Dim t As Single: t = Timer
    Do While Timer - t < seconds
        DoEvents
    Loop
End Sub

Private Function WaitForText(ByVal row As Long, ByVal col As Long, ByVal nChars As Long, _
                             ByVal needle As String, ByVal attempts As Long, ByVal waitSec As Double) As Boolean
    Dim i As Long
    Dim hay As String

    On Error Resume Next
    For i = 1 To attempts
        hay = iCS.GetText(row, col, nChars)
        If InStr(1, hay, needle, vbTextCompare) > 0 Then
            WaitForText = True
            Exit Function
        End If
        SafePause waitSec
    Next i
    On Error GoTo 0

    WaitForText = False
End Function
