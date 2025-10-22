VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'==== ThisWorkbook ====
Option Explicit

Private Sub Workbook_Open()
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
    Application.WindowState = xlMaximized

CleanUp:
    Application.ScreenUpdating = True
    Exit Sub
EH:
    Debug.Print "Workbook_Open error: "; Err.Number; Err.Description
    Resume CleanUp
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

