'==== UserForm: StartupForm ====
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
    On Error GoTo CleanFail

    SetCursorWait

    titleBarHidden = False

    ' === 1. Connect OAIS and set indicators ===
    Dim isConnected As Boolean
    isConnected = EnsureReflectionsConnectionAlive(True)
    SetOAISStatus bOAIS, isConnected

    ' === 2. Load board info safely ===
    lblBoardType.Caption = CStr(SafeCell("ID", "H4")) & " Board"
    lblBoardNum.Caption = "#  " & CStr(SafeCell("ID", "H2"))

    ' === 3. Center form on the active monitor ===
    On Error Resume Next
    CenterUserFormOnActiveMonitor Me
    On Error GoTo CleanFail

    ' === 4. Run OAIS reflection logic (background setup) ===
    If Not iCS Is Nothing Then
        InitializeOAISSession bOAIS
    End If

    modReflectionsMonitor.RegisterReflectionsListener Me.Name

CleanExit:
    SetCursorDefault
    Exit Sub

CleanFail:
    Debug.Print "StartupForm.Initialize error: "; Err.Number; Err.Description
    Resume CleanExit
End Sub

Public Sub HandleReflectionsConnection(ByVal isConnected As Boolean)
    SetOAISStatus bOAIS, isConnected
End Sub

Private Sub UserForm_Terminate()
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    On Error GoTo CleanFail

    SetCursorWait

    On Error Resume Next
    modReflectionsMonitor.UnregisterReflectionsListener Me.Name
    On Error GoTo CleanFail

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

    Static hasRun As Boolean
    If hasRun Then Exit Sub
    hasRun = True

    ' Optional: fade hint before labels appear
    SafePause 1

    RevealLabels Array(lblRadiate, lblNew, lblASTABone, lblASTABtwo), 0.75
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
    tempCaption = "startup-" & Hex$(ObjPtr(Me))
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

'--- Mouse-over visuals (kept simple; fixed vbWhite typo) ---------------------

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    StartupForm.MousePointer = fmMousePointerDefault
    lblNew.ForeColor = vbWhite
    lblRadiate.ForeColor = vbWhite
    bOAIS.ForeColor = vbBlack

    lblASTABone.ForeColor = vbWhite
    lblASTABtwo.ForeColor = vbWhite
End Sub

Private Sub bRadiate_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblRadiate.ForeColor = vbRed
End Sub

Private Sub bOAIS_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If bOAIS.BackColor <> vbGreen Then bOAIS.ForeColor = vbRed
End Sub

Private Sub bNew_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblNew.ForeColor = vbRed
End Sub

Private Sub bastabone_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblASTABone.ForeColor = vbRed
End Sub

Private Sub bastabtwo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblASTABtwo.ForeColor = vbRed
End Sub

'--- Click handlers -----------------------------------------------------------

Private Sub bOAIS_Click()
    ' If not connected, try to connect; else toggle external frame if present.
    If bOAIS.BackColor = vbRed Then
        ConnectToRunningOAIS
        SetOAISStatus bOAIS, Not (iCS Is Nothing)
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
    Dim hasError As Boolean
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo RadiateError

    ConnectToRunningOAIS

    SetOAISStatus bOAIS, Not (iCS Is Nothing)

    ClearTableColumnsCD ("RED_Board")

    KeepAlive_Suspend

    HideAndReleaseStartupForm

    A_Record_Review

CleanExit:
    On Error Resume Next
    KeepAlive_Resume
    On Error GoTo 0

    If hasError Then
        MsgBox "Radiate encountered an error (" & errNumber & "): " & errDescription, vbExclamation, "Radiate"
    End If

    ' Progress UI is managed within A_Record_Review; no separate show call is needed here.
    Exit Sub

RadiateError:
    hasError = True
    errNumber = Err.Number
    errDescription = Err.Description
    Resume CleanExit

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
    
    ' Exit gracefully if the table has no data rows yet
    If lo.DataBodyRange Is Nothing Then
        Debug.Print "ClearTableColumnsCD: Table '" & TableName & "' has no data rows to clear."
        Exit Sub
    End If

    ' Determine first and last data rows in the table
    firstRow = lo.DataBodyRange.row
    lastRow = firstRow + lo.DataBodyRange.Rows.Count - 1

    If lastRow >= firstRow Then
        ' Build the range using the table's actual position
        Set targetRange = ws.Range(ws.Cells(firstRow, 3), ws.Cells(lastRow, 4))

        ' Clear contents only (keeps formatting and formulas)
        targetRange.ClearContents
    End If

    Exit Sub

ErrHandler:
    Debug.Print "ClearTableColumnsCD error (" & Err.Number & "): " & Err.Description
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

