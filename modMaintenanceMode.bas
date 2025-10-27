Attribute VB_Name = "modMaintenanceMode"
Option Explicit

Private Type SheetState
    Name As String
    WasVisible As XlSheetVisibility
    WasProtected As Boolean
    EnableSelection As XlEnableSelection
End Type

Private Type MaintenanceAppState
    DisplayFormulaBar As Boolean
    DisplayFullScreen As Boolean
    ScreenUpdating As Boolean
    DisplayWorkbookTabs As Boolean
    HasDisplayTabs As Boolean
    WorkbookStructureProtected As Boolean
    WorkbookWindowsProtected As Boolean
End Type

Private m_SheetStates() As SheetState
Private m_SheetStateCount As Long
Private m_AppState As MaintenanceAppState
Private m_AppStateSaved As Boolean
Private m_MaintenanceModeActive As Boolean

Public Sub SetMaintenanceMode(Optional ByVal showMessage As Boolean = True)
    Dim ws As Worksheet
    Dim index As Long
    Dim oldEnableEvents As Boolean

    On Error GoTo Fail

    If m_MaintenanceModeActive Then
        If showMessage Then
            MsgBox "Maintenance Mode is already active.", vbInformation
        End If
        Exit Sub
    End If

    oldEnableEvents = Application.EnableEvents
    Application.EnableEvents = False

    CaptureApplicationState

    On Error Resume Next
    ThisWorkbook.Unprotect
    On Error GoTo Fail

    m_SheetStateCount = ThisWorkbook.Worksheets.Count
    If m_SheetStateCount > 0 Then
        ReDim m_SheetStates(1 To m_SheetStateCount)
        index = 0
        For Each ws In ThisWorkbook.Worksheets
            index = index + 1
            StoreSheetState m_SheetStates(index), ws
            PrepareSheetForMaintenance ws
        Next ws
    Else
        Erase m_SheetStates
    End If

    ApplyMaintenanceAppSettings

    m_MaintenanceModeActive = True
    Application.EnableEvents = oldEnableEvents

    If showMessage Then
        MsgBox "AUTO_SPCL is now in Maintenance Mode.", vbInformation
    End If
    Exit Sub

Fail:
    Application.EnableEvents = oldEnableEvents
    If showMessage Then
        MsgBox "Unable to enable Maintenance Mode: " & Err.Description, vbCritical
    End If
End Sub

Public Sub SetNormalMode(Optional ByVal showMessage As Boolean = True)
    Dim ws As Worksheet
    Dim index As Long
    Dim oldEnableEvents As Boolean

    On Error GoTo Fail

    If Not m_MaintenanceModeActive Then
        If showMessage Then
            MsgBox "Maintenance Mode is not currently active.", vbInformation
        End If
        Exit Sub
    End If

    oldEnableEvents = Application.EnableEvents
    Application.EnableEvents = False

    RestoreSheetStates
    RestoreApplicationState

    m_MaintenanceModeActive = False
    Application.EnableEvents = oldEnableEvents

    If showMessage Then
        MsgBox "AUTO_SPCL has been returned to normal mode.", vbInformation
    End If
    Exit Sub

Fail:
    Application.EnableEvents = oldEnableEvents
    If showMessage Then
        MsgBox "Unable to restore normal mode: " & Err.Description, vbCritical
    End If
End Sub

Public Sub ToggleMaintenanceMode()
    If m_MaintenanceModeActive Then
        SetNormalMode True
    Else
        SetMaintenanceMode True
    End If
End Sub

Private Sub CaptureApplicationState()
    If Not m_AppStateSaved Then
        m_AppState.DisplayFormulaBar = Application.DisplayFormulaBar
        m_AppState.DisplayFullScreen = Application.DisplayFullScreen
        m_AppState.ScreenUpdating = Application.ScreenUpdating
        m_AppState.WorkbookStructureProtected = ThisWorkbook.ProtectStructure
        m_AppState.WorkbookWindowsProtected = ThisWorkbook.ProtectWindows
        On Error Resume Next
        m_AppState.DisplayWorkbookTabs = ActiveWindow.DisplayWorkbookTabs
        m_AppState.HasDisplayTabs = Err.Number = 0
        On Error GoTo 0
        m_AppStateSaved = True
    End If
End Sub

Private Sub ApplyMaintenanceAppSettings()
    Application.DisplayFormulaBar = True
    Application.DisplayFullScreen = False
    Application.ScreenUpdating = True
    On Error Resume Next
    ActiveWindow.DisplayWorkbookTabs = True
    On Error GoTo 0
End Sub

Private Sub RestoreApplicationState()
    If m_AppStateSaved Then
        Application.DisplayFormulaBar = m_AppState.DisplayFormulaBar
        Application.DisplayFullScreen = m_AppState.DisplayFullScreen
        Application.ScreenUpdating = m_AppState.ScreenUpdating
        If m_AppState.WorkbookStructureProtected Or m_AppState.WorkbookWindowsProtected Then
            On Error Resume Next
            ThisWorkbook.Protect Structure:=m_AppState.WorkbookStructureProtected, _
                Windows:=m_AppState.WorkbookWindowsProtected
            On Error GoTo 0
        End If
        If m_AppState.HasDisplayTabs Then
            On Error Resume Next
            ActiveWindow.DisplayWorkbookTabs = m_AppState.DisplayWorkbookTabs
            On Error GoTo 0
        End If
        m_AppStateSaved = False
    End If
End Sub

Private Sub StoreSheetState(ByRef state As SheetState, ByVal ws As Worksheet)
    With state
        .Name = ws.Name
        .WasVisible = ws.Visible
        .WasProtected = ws.ProtectContents
        On Error Resume Next
        .EnableSelection = ws.EnableSelection
        On Error GoTo 0
    End With
End Sub

Private Sub PrepareSheetForMaintenance(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Unprotect
    ws.Visible = xlSheetVisible
    ws.EnableSelection = xlNoRestrictions
    On Error GoTo 0
End Sub

Private Sub RestoreSheetStates()
    Dim index As Long
    Dim ws As Worksheet

    If m_SheetStateCount = 0 Then Exit Sub

    For index = 1 To m_SheetStateCount
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(m_SheetStates(index).Name)
        If Not ws Is Nothing Then
            ws.Visible = m_SheetStates(index).WasVisible
            If m_SheetStates(index).WasProtected Then
                ws.Protect
            End If
            On Error Resume Next
            ws.EnableSelection = m_SheetStates(index).EnableSelection
            On Error GoTo 0
        End If
        Set ws = Nothing
        On Error GoTo 0
    Next index

    Erase m_SheetStates
    m_SheetStateCount = 0
End Sub
