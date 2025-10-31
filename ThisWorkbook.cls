'==== ThisWorkbook ====
Option Explicit

Private m_SavedWindowState As XlWindowState
Private m_HasSavedWindowState As Boolean
Private m_SplashScreenShown As Boolean

Private Sub Workbook_Open()
    Dim routineSucceeded As Boolean

    On Error GoTo EH2

    Set iApp = GetObject(, "Attachmate_Reflection_Objects_Framework.ApplicationObject")
    Set iFrame = iApp.GetObject("Frame")
    Set iCT = iFrame.SelectedView.control
    Set iCS = iCT.screen

    On Error GoTo EH
    
    ' Make sure THIS workbook is visible before any UI/setup
    SetWorkbookVisibility ThisWorkbook, True

    Application.DisplayFullScreen = True
    Application.DisplayFormulaBar = False
    If Worksheets("Eligibles RED Board").Visible = False Then Worksheets("Eligibles RED Board").Visible = True
    Worksheets("Eligibles Status Board").Select
    Application.ScreenUpdating = False
    Application.EnableCancelKey = xlErrorHandler
    ActiveWindow.DisplayWorkbookTabs = False
    m_SavedWindowState = Application.WindowState
    m_HasSavedWindowState = True
    Application.WindowState = xlMinimized
       
    ' Only show the splash UI if THIS workbook is visible
    ShowSplashScreenWhenVisible

    ToggleThisWorkbookVisibility
    ' 1) Load Board Type safely
    BoardType = SafeReadBoardType("ID", "H5")

    ' 2) Start keep-alive (best-effort)
    On Error Resume Next
    KeepAlive_Start
    On Error GoTo EH

    ' 3) Ensure the form is present (modeless) once the splash is gone
    If Not SplashScreenLoaded() Then
        ShowStartupFormOnce True
    End If
    routineSucceeded = True

CleanUp:
    If m_HasSavedWindowState Then
        Application.WindowState = m_SavedWindowState
    Else
        Application.WindowState = xlMaximized
    End If
    If routineSucceeded Then
        m_HasSavedWindowState = False
    End If
    Application.ScreenUpdating = True
    Exit Sub
EH:
    Debug.Print "Workbook_Open error: "; Err.Number; Err.Description
    Resume CleanUp
EH2:
    MsgBox "OAIS Session not found. Open OAIS and try opening AUTO_SPCL again.", vbCritical
    Resume CleanUp
    ThisWorkbook.Close
End Sub

Private Sub MaintenanceMode()
    Dim cancelRequest As Boolean

    HandleWorkbookShutdown cancelRequest
End Sub

Private Sub Workbook_Activate()
    ShowSplashScreenWhenVisible
    If Not g_SuspendStartupAutoShow Then
        ShowStartupFormOnce
    End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    HandleWorkbookShutdown Cancel
End Sub

Private Sub HandleWorkbookShutdown(ByRef cancelClose As Boolean)
    Dim response As VbMsgBoxResult
    Dim frm As Object

    On Error GoTo Fail

    Application.ScreenUpdating = False

    ' Prompt to save changes (allow cancel)
    If Not ThisWorkbook.Saved Then
        response = MsgBox("Do you want to save changes before closing?", vbYesNoCancel + vbQuestion, ThisWorkbook.Name)
        Select Case response
            Case vbYes
                On Error GoTo SaveError
                ThisWorkbook.Save
                On Error GoTo Fail
            Case vbNo
                ' continue with close
            Case Else
                cancelClose = True
                GoTo CleanExit
        End Select
    End If

    ' Stop background tasks and clear temporary state
    On Error Resume Next
    KeepAlive_Stop
    modProgressUI.Progress_Close
    ResetRunState
    BoardType = ""
    processed = 0
    total = 0
    modProgressUI.IsCancellationRequested = False
    m_StartupShownOnce = False
    g_SuspendStartupAutoShow = False
    m_SplashScreenShown = False
    Set iCS = Nothing
    Set iCT = Nothing
    Set iFrame = Nothing
    Set iApp = Nothing
    On Error GoTo Fail

    ' Close any userforms that might still be open
    HideAndReleaseStartupForm
    For Each frm In VBA.UserForms
        Unload frm
    Next frm

    ' Restore workbook and application UI defaults
    SetWorkbookVisibility ThisWorkbook, True
    Application.DisplayFullScreen = False
    Application.DisplayFormulaBar = True
    On Error Resume Next
    ActiveWindow.DisplayWorkbookTabs = True
    On Error GoTo Fail

CleanExit:
    Application.ScreenUpdating = True
    Exit Sub

SaveError:
    MsgBox "Unable to save changes: " & Err.Description, vbCritical
    cancelClose = True
    On Error GoTo Fail
    GoTo CleanExit

Fail:
    Application.ScreenUpdating = True
    MsgBox "An error occurred while closing the workbook: " & Err.Description, vbCritical
    cancelClose = True
End Sub

' --- Called from Workbook_Open to guard one-time autoshow intent ---
Private Sub ShowStartupOnceOnOpen()
    If m_StartupShownOnce Then Exit Sub
    If Not IsWorkbookVisible(ThisWorkbook) Then Exit Sub
    ufSplashScreen.Show vbModeless
End Sub
' --- Called from Workbook_Open to guard one-time splash intent ---
Private Sub ShowSplashScreenWhenVisible()
    If m_SplashScreenShown Then Exit Sub
    If Not IsWorkbookVisible(ThisWorkbook) Then Exit Sub

    ufSplashScreen.Show vbModeless
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

Private Function SafeReadBoardType(ByVal sheetName As String, ByVal addr As String) As String
    On Error GoTo Fallback
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    SafeReadBoardType = CStr(ws.Range(addr).Value2)
    Exit Function
Fallback:
    SafeReadBoardType = ""
End Function

