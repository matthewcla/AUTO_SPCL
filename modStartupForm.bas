Attribute VB_Name = "modStartupForm"
'======================
' Module: modStartupForm
'======================
Option Explicit

Private Const STARTUP_FORM_NAME As String = "StartupForm"

' Prevents auto-show during visibility flips / activation churn
Public g_SuspendStartupAutoShow As Boolean

' Tracks whether we already auto-showed once this session
Public m_StartupShownOnce As Boolean

' === Visibility control ===
Public Sub SetStartupFormVisibility(ByVal makeVisible As Boolean)
    Dim frm As Access.Form
    On Error GoTo CleanExit

    ' Avoid auto-show loops while changing visibility
    BeginSuspendAutoShow

    If makeVisible Then
        If StartupFormLoaded() Then
            Set frm = Forms(STARTUP_FORM_NAME)
            frm.Visible = True
            modUIHelpers.EnsureFormFocus frm
        Else
            DoCmd.OpenForm STARTUP_FORM_NAME, WindowMode:=acWindowNormal
        End If
    Else
        If StartupFormLoaded() Then
            Forms(STARTUP_FORM_NAME).Visible = False
        End If
    End If

    BringStartupToFrontIfLoaded

CleanExit:
    EndSuspendAutoShow
End Sub

Public Function IsStartupFormVisible() As Boolean
    On Error Resume Next
    If StartupFormLoaded() Then
        IsStartupFormVisible = Forms(STARTUP_FORM_NAME).Visible
    End If
End Function

' === StartupForm single-instance helpers ===
Public Sub HandleSplashComplete()
    If m_StartupShownOnce Then Exit Sub

    m_StartupShownOnce = True
    ShowStartupFormOnce True
End Sub

Public Sub ShowStartupFormOnce(Optional ByVal forceShow As Boolean = False)
    Dim frm As Access.Form
    If g_SuspendStartupAutoShow Then Exit Sub
    If m_StartupShownOnce And Not forceShow Then Exit Sub

    ' If already up, just bring to front
    If StartupFormLoaded() Then
        Set frm = Forms(STARTUP_FORM_NAME)
        If Not frm.Visible Then frm.Visible = True
        modUIHelpers.EnsureFormFocus frm
        Exit Sub
    End If

    ' Not loaded: open the Access form modelessly
    DoCmd.OpenForm STARTUP_FORM_NAME, WindowMode:=acWindowNormal
    m_StartupShownOnce = True
End Sub

Public Sub HideAndReleaseStartupForm()
    If StartupFormLoaded() Then
        DoCmd.Close acForm, STARTUP_FORM_NAME, acSaveNo
    End If
End Sub

Public Function StartupFormLoaded() As Boolean
    On Error GoTo ExitCheck
    StartupFormLoaded = (SysCmd(acSysCmdGetObjectState, acForm, STARTUP_FORM_NAME) <> 0)
ExitCheck:
End Function

' === Safe toggle for startup form ===
Public Sub ToggleStartupFormVisibility(Optional ByVal keepFormOnTop As Boolean = True)
    Dim makeVisible As Boolean
    makeVisible = Not IsStartupFormVisible()
    SetStartupFormVisibility makeVisible

    If keepFormOnTop Then BringStartupToFrontIfLoaded
End Sub

Private Sub BringStartupToFrontIfLoaded()
    If StartupFormLoaded() Then
        On Error Resume Next
        DoCmd.SelectObject acForm, STARTUP_FORM_NAME, False
        Forms(STARTUP_FORM_NAME).SetFocus
        On Error GoTo 0
    End If
End Sub

' === AutoShow suspend guards (useful while changing visibility) ===
Public Sub BeginSuspendAutoShow()
    On Error Resume Next
    g_SuspendStartupAutoShow = True
End Sub

Public Sub EndSuspendAutoShow()
    On Error Resume Next
    g_SuspendStartupAutoShow = False
End Sub
