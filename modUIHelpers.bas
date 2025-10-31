Attribute VB_Name = "modUIHelpers"
Option Explicit

Private mWaitCursorDepth As Long
Private Const DEFAULT_MESSAGE_TITLE As String = "AUTO_SPCL"

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

'-------------------------------------------------------------------------------
' Procedure: SetCursorWait
' Purpose  : Display the wait cursor during long-running operations, nesting safely.
' Parameters: None.
' Returns  : None.
' Side Effects:
'   Increments the wait cursor depth counter and switches Application.Cursor to xlWait
'   when transitioning from idle to busy state.
'-------------------------------------------------------------------------------
Public Sub SetCursorWait()
    On Error Resume Next
    mWaitCursorDepth = mWaitCursorDepth + 1
    If mWaitCursorDepth = 1 Then
        Application.Cursor = xlWait
    End If
    On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------
' Procedure: SetCursorDefault
' Purpose  : Restore the standard cursor once all nested wait scopes have completed.
' Parameters: None.
' Returns  : None.
' Side Effects:
'   Decrements the depth counter and resets Application.Cursor to xlDefault when depth
'   reaches zero.
'-------------------------------------------------------------------------------
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

'-------------------------------------------------------------------------------
' Procedure: ShowInfoMessage
' Purpose  : Present a consistent informational message dialog to the user.
'-------------------------------------------------------------------------------
Public Function ShowInfoMessage(ByVal message As String, _
                                Optional ByVal title As String = DEFAULT_MESSAGE_TITLE) As VbMsgBoxResult
    ShowInfoMessage = MsgBox(message, vbInformation + vbOKOnly, title)
End Function

'-------------------------------------------------------------------------------
' Procedure: ShowWarningMessage
' Purpose  : Present a consistent warning dialog to the user.
'-------------------------------------------------------------------------------
Public Function ShowWarningMessage(ByVal message As String, _
                                   Optional ByVal title As String = DEFAULT_MESSAGE_TITLE) As VbMsgBoxResult
    ShowWarningMessage = MsgBox(message, vbExclamation + vbOKOnly, title)
End Function

'-------------------------------------------------------------------------------
' Procedure: ShowErrorMessage
' Purpose  : Present a consistent critical-error dialog to the user.
'-------------------------------------------------------------------------------
Public Function ShowErrorMessage(ByVal message As String, _
                                 Optional ByVal title As String = DEFAULT_MESSAGE_TITLE) As VbMsgBoxResult
    ShowErrorMessage = MsgBox(message, vbCritical + vbOKOnly, title)
End Function

'-------------------------------------------------------------------------------
' Procedure: ShowDecisionMessage
' Purpose  : Present a consistent yes/no style decision dialog to the user.
'-------------------------------------------------------------------------------
Public Function ShowDecisionMessage(ByVal message As String, _
                                    Optional ByVal buttons As VbMsgBoxStyle = vbYesNo, _
                                    Optional ByVal title As String = DEFAULT_MESSAGE_TITLE) As VbMsgBoxResult
    ShowDecisionMessage = MsgBox(message, buttons, title)
End Function

'-------------------------------------------------------------------------------
' Procedure: FocusControl
' Purpose  : Safely attempt to move focus to the supplied control.
'-------------------------------------------------------------------------------
Public Sub FocusControl(ByVal control As Object)
    If control Is Nothing Then Exit Sub

    On Error Resume Next
    CallByName control, "SetFocus", VbMethod
    If Err.Number <> 0 Then
        Err.Clear
        CallByName control, "Activate", VbMethod
    End If
    On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------
' Procedure: EnsureFormFocus
' Purpose  : Bring a form back to the foreground, optionally targeting a control.
'-------------------------------------------------------------------------------
Public Sub EnsureFormFocus(ByVal form As Object, Optional ByVal fallbackControl As Object)
    If Not fallbackControl Is Nothing Then
        FocusControl fallbackControl
    End If

    If form Is Nothing Then Exit Sub

    On Error Resume Next
    CallByName form, "SetFocus", VbMethod
    If Err.Number <> 0 Then
        Err.Clear
        CallByName form, "Activate", VbMethod
    End If
    On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------
' Procedure: HideUserFormTitleBar
' Purpose  : Remove the standard Windows title bar from a userform to present a clean UI.
' Parameters:
'   targetForm - The userform instance to modify.
'   titleBarHiddenFlag - Boolean flag that tracks whether the form has already been hidden.
'   captionPrefix - Optional prefix used to reliably identify the window handle.
' Returns  : None.
' Side Effects:
'   Adjusts the form's window style using Win32 APIs and updates the tracking flag.
'-------------------------------------------------------------------------------
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
    originalCaption = CStr(CallByName(targetForm, "Caption", vbGet))
    tempCaption = captionPrefix & "-" & Hex$(ObjPtr(targetForm))
    CallByName targetForm, "Caption", vbLet, tempCaption
    hWnd = FindWindow("ThunderDFrame", tempCaption)
    CallByName targetForm, "Caption", vbLet, originalCaption
    On Error GoTo 0

    If hWnd = 0 Then Exit Sub

    currentStyle = GetWindowLong(hWnd, GWL_STYLE)
    newStyle = currentStyle And (Not WS_CAPTION)
    SetWindowLong hWnd, GWL_STYLE, newStyle
    DrawMenuBar hWnd

    titleBarHiddenFlag = True
End Sub

'-------------------------------------------------------------------------------
' Procedure: SetControlsEnabled
' Purpose  : Apply a consistent enabled/disabled state to one or more controls.
' Parameters:
'   controls - Individual control, array, or collection of controls to update.
'   enabled - Target Boolean enabled state.
' Returns  : None.
' Side Effects:
'   Calls CallByName on each control to set the Enabled property.
'-------------------------------------------------------------------------------
Public Sub SetControlsEnabled(ByVal controls As Variant, ByVal enabled As Boolean)
    ApplyControlBooleanProperty controls, "Enabled", enabled
End Sub

'-------------------------------------------------------------------------------
' Procedure: SetControlsVisible
' Purpose  : Apply a consistent visibility state to one or more controls.
' Parameters:
'   controls - Individual control, array, or collection of controls to update.
'   isVisible - Target Boolean visibility state.
' Returns  : None.
' Side Effects:
'   Calls CallByName on each control to set the Visible property.
'-------------------------------------------------------------------------------
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
