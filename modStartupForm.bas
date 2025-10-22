Attribute VB_Name = "modStartupForm"
'======================
' Module: modStartupForm
'======================
Option Explicit


#If VBA7 Then
    Private Declare PtrSafe Function FindWindowA Lib "user32" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" _
        (ByVal hWnd As LongPtr) As Long
#Else
    Private Declare Function FindWindowA Lib "user32" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function SetForegroundWindow Lib "user32" _
        (ByVal hWnd As Long) As Long
#End If

' Prevents auto-show during visibility flips / activation churn
Public g_SuspendStartupAutoShow As Boolean

' Tracks whether we already auto-showed once this session
Public m_StartupShownOnce As Boolean

Public startup As Object

' === Visibility control ===
Public Sub SetWorkbookVisibility(ByVal wb As Workbook, ByVal makeVisible As Boolean)
    Dim i As Long
    On Error GoTo CleanFail

    ' Avoid Activate/Deactivate-triggered autoshow while we change visibility
    BeginSuspendAutoShow

    If wb Is Nothing Then GoTo CleanExit
    If wb.Windows.Count = 0 Then wb.Activate

    For i = 1 To wb.Windows.Count
        wb.Windows(i).Visible = makeVisible
    Next i

    BringStartupToFrontIfLoaded

CleanExit:
    EndSuspendAutoShow
    Exit Sub
CleanFail:
    EndSuspendAutoShow
End Sub

Public Function IsWorkbookVisible(ByVal wb As Workbook) As Boolean
    Dim i As Long
    On Error Resume Next
    If wb Is Nothing Then Exit Function
    For i = 1 To wb.Windows.Count
        If wb.Windows(i).Visible Then IsWorkbookVisible = True: Exit Function
    Next i
End Function

' === StartupForm single-instance helpers ===
Public Sub ShowStartupFormOnce()
    Dim uf As Object
    If Not ThisWorkbookIsFrontCandidate() Then Exit Sub

    ' If already up, just bring to front
    If StartupFormLoaded() Then
        Set uf = GetLoadedStartupForm
        If Not uf Is Nothing Then
            If uf.Visible = False Then uf.Show vbModeless
            uf.ZOrder 0
        End If
        Exit Sub
    End If

    ' Not loaded: show modeless one time
    StartupForm.Show vbModeless
End Sub

Public Sub HideAndReleaseStartupForm()
    Dim uf As Object
    If StartupFormLoaded() Then
        Set uf = GetLoadedStartupForm
        If Not uf Is Nothing Then Unload uf
    End If
End Sub

Public Function StartupFormLoaded() As Boolean
    Dim uf As Object
    For Each uf In VBA.UserForms
        If StrComp(TypeName(uf), "StartupForm", vbTextCompare) = 0 Then
            StartupFormLoaded = True
            Exit Function
        End If
    Next uf
End Function

' === Safe toggle for THIS workbook ===
Public Sub ToggleThisWorkbookVisibility(Optional ByVal keepFormOnTop As Boolean = True)
    Dim makeVisible As Boolean
    makeVisible = Not IsWorkbookVisible(ThisWorkbook)
    SetWorkbookVisibility ThisWorkbook, makeVisible

    If keepFormOnTop Then BringStartupToFrontIfLoaded
End Sub

' Return the live StartupForm instance (or Nothing) as the correct type
Private Function GetLoadedStartupForm() As StartupForm
    Dim f As Object
    For Each f In VBA.UserForms
        If TypeOf f Is StartupForm Then
            Set GetLoadedStartupForm = f
            Exit Function
        End If
    Next f
End Function


Private Sub BringStartupToFrontIfLoaded()
    Dim uf As Object, hWndForm As LongPtr
    Set uf = GetLoadedStartupForm()
    If Not uf Is Nothing Then
        If uf.Visible Then
            hWndForm = FindWindowA(vbNullString, uf.Caption)
            If hWndForm <> 0 Then SetForegroundWindow hWndForm
        End If
    End If
End Sub

Private Function ThisWorkbookIsFrontCandidate() As Boolean
    ' Only auto-show if this workbook actually has a visible window (not hidden/add-in)
    ThisWorkbookIsFrontCandidate = IsWorkbookVisible(ThisWorkbook)
End Function

' === AutoShow suspend guards (useful while changing visibility) ===
Public Sub BeginSuspendAutoShow()
    On Error Resume Next
    g_SuspendStartupAutoShow = True
End Sub

Public Sub EndSuspendAutoShow()
    On Error Resume Next
    g_SuspendStartupAutoShow = False
End Sub

