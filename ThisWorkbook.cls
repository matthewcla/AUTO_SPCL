'==== Access Startup Bootstrap ====
Option Explicit

Private m_SplashScreenShown As Boolean
Private m_IsShuttingDown As Boolean

' Entry point for Access AutoExec macro
Public Sub AutoExec_Startup()
    If m_IsShuttingDown Then Exit Sub

    ShowSplashScreenWhenVisible
    ShowStartupFormOnce True
End Sub

' Hook for Access form load to ensure startup UI is present
Public Sub OnStartupFormLoad()
    If m_IsShuttingDown Then Exit Sub
    HandleSplashComplete
End Sub

' Gracefully close any startup UI when Access shuts down
Public Sub HandleApplicationShutdown()
    If m_IsShuttingDown Then Exit Sub

    m_IsShuttingDown = True
    HideAndReleaseStartupForm
    m_SplashScreenShown = False
End Sub

Friend Property Get IsShuttingDown() As Boolean
    IsShuttingDown = m_IsShuttingDown
End Property

' --- Splash handling for Access UI ---
Private Sub ShowSplashScreenWhenVisible()
    If m_SplashScreenShown Then Exit Sub

    On Error Resume Next
    ufSplashScreen.Show vbModeless
    On Error GoTo 0
    m_SplashScreenShown = True
End Sub

Private Function SplashScreenLoaded() As Boolean
    Dim uf As Object
    For Each uf In VBA.UserForms
        If TypeOf uf Is ufSplashScreen Then
            SplashScreenLoaded = True
            Exit Function
        End If
    Next uf
End Function
