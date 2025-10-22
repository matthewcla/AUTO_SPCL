VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StartupForm 
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9255.001
   OleObjectBlob   =   "StartupForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "StartupForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==== UserForm: StartupForm ====
Option Explicit

' --- Constants for screen-text detection (Reflection / OAIS banners) ---
Private Const TXT_DISA As String = "Defense Information Systems Agency"
Private Const TXT_SESSION_MENU As String = "CL/SuperSession"
Private Const TXT_OAIS_BANNER As String = "Officer Assignment Information System"
Private Const OAIS_MENU_CMD As String = "Start OAIS2"

' --- Tuning knobs ---
Private Const SMALL_WAIT_SEC As Double = 0.75
Private Const RETRY_WAIT_SEC As Double = 1.2
Private Const retries As Long = 3

Private Sub UserForm_Initialize()
    On Error GoTo EH

    ' === 1. Connect OAIS and set indicators ===
    ConnectToRunningOAIS
    SetOAISStatus Not (iCS Is Nothing)

    ' === 2. Load board info safely ===
    lblBoardType.Caption = CStr(SafeCell("ID", "H4")) & " Board"
    lblBoardNum.Caption = "#  " & CStr(SafeCell("ID", "H2"))

    ' === 3. Center form on the active monitor ===
    On Error Resume Next
    CenterUserFormOnActiveMonitor Me
    On Error GoTo EH

    ' === 4. Run OAIS reflection logic (background setup) ===
    If Not iCS Is Nothing Then
        InitializeReflectionAndOAIS
    End If

    Exit Sub
EH:
    Debug.Print "StartupForm.Initialize error: "; Err.Number; Err.Description
End Sub


Private Sub UserForm_Activate()
    Static hasRun As Boolean
    If hasRun Then Exit Sub
    hasRun = True

    ' Optional: fade hint before labels appear
    SafePause 1

    RevealLabels Array(lblRadiate, lblNew, lblASTABone, lblASTABtwo), 0.75
End Sub

'--- Drive Reflection > Session menu > OAIS2 with light retries ---
Private Sub InitializeReflectionAndOAIS()
    ' (1) Reflection Workspace Intro Screen?
    If WaitForText(1, 1, 79, TXT_DISA, retries, RETRY_WAIT_SEC) Then
        HitEnter ' pass the splash / login handoff

        ' (2) Session selection menu?
        If WaitForText(3, 1, 79, TXT_SESSION_MENU, retries, RETRY_WAIT_SEC) Then
            entText 23, 15, OAIS_MENU_CMD

            ' (3) Wait for OAIS banner, allow one "enter" nudge if needed
            If Not WaitForText(2, 1, 79, TXT_OAIS_BANNER, 2, RETRY_WAIT_SEC) Then
                SafePause 0.6
                HitEnter
                ' final check
                Call WaitForText(2, 1, 79, TXT_OAIS_BANNER, 2, RETRY_WAIT_SEC)
            End If
        End If
    End If

    ' Refresh status light after attempts
    SetOAISStatus Not (iCS Is Nothing)
End Sub

'--- Label/status helpers -----------------------------------------------------

Private Sub SetOAISStatus(ByVal isConnected As Boolean)
    If isConnected Then
        bOAIS.BackColor = vbGreen
        bOAIS.Caption = "Connected to OAIS"
    Else
        bOAIS.BackColor = vbRed
        bOAIS.Caption = "OAIS Not Connected"
    End If
End Sub

Private Sub RevealLabels(ByVal labels As Variant, ByVal stepSeconds As Double)
    Dim i As Long
    For i = LBound(labels) To UBound(labels)
        On Error Resume Next
        labels(i).Visible = True
        On Error GoTo 0
        SafePause stepSeconds
    Next i
End Sub

'--- Mouse-over visuals (kept simple; fixed vbWhite typo) ---------------------

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblNew.ForeColor = vbWhite:        lblNewL.ForeColor = vbWhite
    lblSettings.ForeColor = vbBlack
    lblRadiate.ForeColor = vbWhite:    lblRadiateL.ForeColor = vbWhite
    bOAIS.ForeColor = vbBlack

    lblRadiateL.Visible = False
    lblNewL.Visible = False

    lblASTABone.ForeColor = vbWhite:   lblASTABoneL.ForeColor = vbWhite: lblASTABoneL.Visible = False
    lblASTABtwo.ForeColor = vbWhite:   lblASTABtwoL.ForeColor = vbWhite: lblASTABtwoL.Visible = False
End Sub

Private Sub bSettings_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblSettings.ForeColor = vbWhite
End Sub

Private Sub bRadiate_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblRadiate.ForeColor = vbRed: lblRadiateL.ForeColor = vbRed: lblRadiateL.Visible = True
End Sub

Private Sub bOAIS_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If bOAIS.BackColor <> vbGreen Then bOAIS.ForeColor = vbRed
End Sub

Private Sub bNew_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblNew.ForeColor = vbRed: lblNewL.ForeColor = vbRed: lblNewL.Visible = True
End Sub

Private Sub bastabone_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblASTABone.ForeColor = vbRed: lblASTABoneL.ForeColor = vbRed: lblASTABoneL.Visible = True
End Sub

Private Sub bastabtwo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblASTABtwo.ForeColor = vbRed: lblASTABtwoL.ForeColor = vbRed: lblASTABtwoL.Visible = True
End Sub

'--- Click handlers -----------------------------------------------------------

Private Sub bOAIS_Click()
    ' If not connected, try to connect; else toggle external frame if present.
    If bOAIS.BackColor = vbRed Then
        ConnectToRunningOAIS
        SetOAISStatus Not (iCS Is Nothing)
        Exit Sub
    End If

    ' Optional: toggle an external host frame if your environment exposes one.
    On Error Resume Next
    If Not (iFrame Is Nothing) Then
        ' Late-bound: property may not exist in all hosts; safe-guarded.
        If LCase$(CStr(CallByName(iFrame, "WindowState", VbGet))) = "0" Then
            CallByName iFrame, "WindowState", VbLet, 1   ' minimize
        Else
            CallByName iFrame, "WindowState", VbLet, 0   ' normal
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub bRadiate_Click()
    ConnectToRunningOAIS
    
    SetOAISStatus Not (iCS Is Nothing)
    
    ClearTableColumnsCD ("RED_Board")
    
    KeepAlive_Suspend
    
    HideAndReleaseStartupForm
    
    A_Record_Review
    
    ' Initialize and show progressform
    progressform.Show vbModeless ' Show modeless so code continues running
         
End Sub

Private Sub ClearTableColumnsCD(ByVal TableName As String)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lastRow As Long
    Dim firstRow As Long
    Dim targetRange As Range
    
    On Error GoTo ErrHandler
    
    ' Reference your sheet (adjust if needed)
    Set ws = ThisWorkbook.Worksheets("Eligibles RED Board")
    
    ' Get the table object
    Set lo = ws.ListObjects(TableName)
    
    ' Determine first and last data rows in the table
    firstRow = lo.DataBodyRange.row
    lastRow = firstRow + lo.ListRows.Count - 1
    
    If Not lastRow = 2 Then
        ' Build the range from C2 to D at last table row
        Set targetRange = ws.Range("C2:D" & lastRow)
        
        ' Clear contents only (keeps formatting and formulas)
        targetRange.ClearContents
    
        Exit Sub
    
    Else
    
        Exit Sub
        
    End If
    
ErrHandler:
    MsgBox "Error: " & Err.Description, vbExclamation
End Sub

Private Sub bASTABone_Click()
    Dim wsSource As Worksheet
    Dim wbNew As Workbook
    Dim strFileName As String
    Dim strPath As String

    On Error GoTo ErrHandler

        '--- Identify the source worksheet
        Set wsSource = ThisWorkbook.Worksheets("Eligibles Status Board")
        
        '--- Create a new workbook and copy the worksheet into it
        wsSource.Copy
        Set wbNew = ActiveWorkbook ' The workbook created by .Copy becomes ActiveWorkbook
        
        '--- Optional: set save path (same folder as this workbook)
        strPath = ThisWorkbook.Path
        If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
        
        '--- Optional: generate timestamped filename
        strFileName = Format(Now, "yyyy-mm-dd_hhmmss") & " " & BoardType & " - Eligibles Status Board Export" & ".xlsx"
        
        '--- Save the new workbook
        Application.DisplayAlerts = False
        wbNew.SaveAs Filename:=strPath & strFileName, FileFormat:=xlOpenXMLWorkbook ' .xlsx format
        Application.DisplayAlerts = True
        
        '--- Notify user
        MsgBox "Export complete!" & vbCrLf & _
               "File saved as:" & vbCrLf & strPath & strFileName, vbInformation, "Export Successful"
    
        Exit Sub

ErrHandler:
    Application.DisplayAlerts = True
    MsgBox "An error occurred during export: " & Err.Description, vbExclamation, "Export Failed"
End Sub

Private Sub bASTABtwo_Click()
    Dim wsSource As Worksheet
    Dim wbNew As Workbook
    Dim strFileName As String
    Dim strPath As String
    
    On Error GoTo ErrHandler
    
    '--- Identify the source worksheet
    Set wsSource = ThisWorkbook.Worksheets("Eligibles RED Board")
    
    '--- Create a new workbook and copy the worksheet into it
    wsSource.Copy
    Set wbNew = ActiveWorkbook ' The workbook created by .Copy becomes ActiveWorkbook
    
    '--- Optional: set save path (same folder as this workbook)
    strPath = ThisWorkbook.Path
    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
    
    '--- Optional: generate timestamped filename
    strFileName = Format(Now, "yyyy-mm-dd_hhmmss") & " " & BoardType & " - Eligibles RED Board Export" & ".xlsx"
    
    '--- Save the new workbook
    Application.DisplayAlerts = False
    wbNew.SaveAs Filename:=strPath & strFileName, FileFormat:=xlOpenXMLWorkbook ' .xlsx format
    Application.DisplayAlerts = True
    
    '--- Notify user
    MsgBox "Export complete!" & vbCrLf & _
           "File saved as:" & vbCrLf & strPath & strFileName, vbInformation, "Export Successful"

    Exit Sub

ErrHandler:
    Application.DisplayAlerts = True
    MsgBox "An error occurred during export: " & Err.Description, vbExclamation, "Export Failed"
End Sub

Private Sub bSettings_Click()
    ToggleThisWorkbookVisibility
End Sub

'--- Small utilities ----------------------------------------------------------

Private Function SafeCell(ByVal sheetName As String, ByVal addr As String) As Variant
    On Error GoTo EH
    SafeCell = ThisWorkbook.Worksheets(sheetName).Range(addr).Value2
    Exit Function
EH:
    SafeCell = vbNullString
End Function

' Non-blocking short wait that keeps UI responsive.
Private Sub SafePause(ByVal seconds As Double)
    Dim t As Single: t = Timer
    Do While Timer - t < seconds
        DoEvents
    Loop
End Sub


' Polls for substring on the Reflection screen text with retries.
Private Function WaitForText(ByVal row As Long, ByVal col As Long, ByVal nChars As Long, _
                             ByVal needle As String, ByVal retries As Long, ByVal waitSec As Double) As Boolean
    Dim i As Long, hay As String
    On Error Resume Next
    For i = 1 To retries
        hay = iCS.GetText(row, col, nChars)
        If InStr(1, hay, needle, vbTextCompare) > 0 Then
            WaitForText = True
            Exit Function
        End If
        SafePause waitSec
    Next i
    WaitForText = False
End Function

Private Sub UserForm_Terminate()
    On Error Resume Next
    Set startup = Nothing
End Sub
