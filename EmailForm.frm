VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EmailForm
   ClientHeight    =   12285
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14355
   OleObjectBlob   =   "EmailForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EmailForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

Private titleBarHidden As Boolean

Private Sub UserForm_Initialize()
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    On Error GoTo CleanFail

    SetCursorWait

    titleBarHidden = False

CleanExit:
    SetCursorDefault
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
    Exit Sub

CleanFail:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Resume CleanExit
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    On Error GoTo CleanFail

    SetCursorWait

CleanExit:
    SetCursorDefault
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
    Exit Sub

CleanFail:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Resume CleanExit
End Sub

Private Sub UserForm_Activate()
    If Not titleBarHidden Then
        HideTitleBar
    End If
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub HideTitleBar()
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

    originalCaption = Me.Caption
    tempCaption = "email-" & Hex$(ObjPtr(Me))
    Me.Caption = tempCaption

    hWnd = FindWindow("ThunderDFrame", tempCaption)
    Me.Caption = originalCaption

    If hWnd = 0 Then Exit Sub

    currentStyle = GetWindowLong(hWnd, GWL_STYLE)
    newStyle = currentStyle And (Not WS_CAPTION)
    SetWindowLong hWnd, GWL_STYLE, newStyle
    DrawMenuBar hWnd

    titleBarHidden = True
End Sub
