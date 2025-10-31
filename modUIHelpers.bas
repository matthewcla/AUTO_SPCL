Attribute VB_Name = "modUIHelpers"
Option Explicit

Private mWaitCursorDepth As Long

#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongPtrA" ( _
        ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongPtrA" ( _
        ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As LongPtr) As Long
#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
        ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
        ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
#End If

Private Const GWL_STYLE As Long = -16
Private Const WS_CAPTION As Long = &HC00000

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

Public Sub HideUserFormTitleBar(ByVal targetForm As Object, _
                                ByRef titleBarHiddenFlag As Boolean, _
                                Optional ByVal captionPrefix As String = "form")
#If VBA7 Then
    Dim hWnd As LongPtr
    Dim currentStyle As LongPtr
    Dim newStyle As LongPtr
#Else
    Dim hWnd As Long
    Dim currentStyle As Long
    Dim newStyle As Long
#End If
    Dim originalCaption As String
    Dim tempCaption As String

    If targetForm Is Nothing Then Exit Sub
    If titleBarHiddenFlag Then Exit Sub

    On Error Resume Next
    originalCaption = CStr(CallByName(targetForm, "Caption", VbGet))
    tempCaption = captionPrefix & "-" & Hex$(ObjPtr(targetForm))
    CallByName(targetForm, "Caption", VbLet, tempCaption)
    hWnd = FindWindow("ThunderDFrame", tempCaption)
    CallByName(targetForm, "Caption", VbLet, originalCaption)
    On Error GoTo 0

    If hWnd = 0 Then Exit Sub

    currentStyle = GetWindowLong(hWnd, GWL_STYLE)
    newStyle = currentStyle And (Not WS_CAPTION)
    SetWindowLong hWnd, GWL_STYLE, newStyle
    DrawMenuBar hWnd

    titleBarHiddenFlag = True
End Sub

Public Sub SetControlsEnabled(ByVal controls As Variant, ByVal enabled As Boolean)
    ApplyControlBooleanProperty controls, "Enabled", enabled
End Sub

Public Sub SetControlsVisible(ByVal controls As Variant, ByVal isVisible As Boolean)
    ApplyControlBooleanProperty controls, "Visible", isVisible
End Sub

Private Sub ApplyControlBooleanProperty(ByVal controls As Variant, _
                                        ByVal propertyName As String, _
                                        ByVal propertyValue As Boolean)
    Dim control As Variant

    If IsObject(controls) Then
        If TypeOf controls Is Collection Then
            For Each control In controls
                ApplySingleControlProperty control, propertyName, propertyValue
            Next control
        Else
            ApplySingleControlProperty controls, propertyName, propertyValue
        End If
    ElseIf IsArray(controls) Then
        For Each control In controls
            ApplySingleControlProperty control, propertyName, propertyValue
        Next control
    Else
        ApplySingleControlProperty controls, propertyName, propertyValue
    End If
End Sub

Private Sub ApplySingleControlProperty(ByVal control As Variant, _
                                       ByVal propertyName As String, _
                                       ByVal propertyValue As Boolean)
    Dim target As Object

    If Not IsObject(control) Then Exit Sub

    Set target = control
    If target Is Nothing Then Exit Sub

    On Error Resume Next
    CallByName target, propertyName, VbLet, propertyValue
    On Error GoTo 0
End Sub
