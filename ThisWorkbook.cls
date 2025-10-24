'==== ThisWorkbook ====
Option Explicit

Private m_SavedWindowState As XlWindowState
Private m_HasSavedWindowState As Boolean

Private Sub Workbook_Open()
    Dim routineSucceeded As Boolean

    On Error GoTo EH2

    Set iApp = GetObject(, "Attachmate_Reflection_Objects_Framework.ApplicationObject")
    Set iFrame = iApp.GetObject("Frame")
    Set iCT = iFrame.SelectedView.Control
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
       
    ' Only show the startup UI if THIS workbook is visible
    ShowStartupOnceOnOpen

    ToggleThisWorkbookVisibility
    ' 1) Load Board Type safely
    BoardType = SafeReadBoardType("ID", "H5")

    ' 2) Start keep-alive (best-effort)
    On Error Resume Next
    KeepAlive_Start
    On Error GoTo EH

    ' 3) Ensure the form is present (modeless)
    ShowStartupFormOnce
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

Private Sub Workbook_Activate()
    If Not g_SuspendStartupAutoShow Then
        ShowStartupFormOnce
    End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Always restore workbook visibility on exit
    SetWorkbookVisibility ThisWorkbook, True

    HideAndReleaseStartupForm
    Application.DisplayFullScreen = False
    Application.DisplayFormulaBar = True
    ActiveWindow.DisplayWorkbookTabs = True
End Sub

' --- Called from Workbook_Open to guard one-time autoshow intent ---
Private Sub ShowStartupOnceOnOpen()
    If m_StartupShownOnce Then Exit Sub
    If Not IsWorkbookVisible(ThisWorkbook) Then Exit Sub
    ShowStartupFormOnce
    m_StartupShownOnce = True
End Sub

Private Function SafeReadBoardType(ByVal sheetName As String, ByVal addr As String) As String
    On Error GoTo Fallback
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    SafeReadBoardType = CStr(ws.Range(addr).Value2)
    Exit Function
Fallback:
    SafeReadBoardType = ""
End Function

