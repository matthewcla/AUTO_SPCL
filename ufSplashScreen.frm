VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufSplashScreen 
   ClientHeight    =   4665
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9135.001
   OleObjectBlob   =   "ufSplashScreen.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufSplashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--- Global Variables ---
Dim progress As Integer
Dim progressText(5) As String
Dim totalProgressBarWidth As Long
Dim originalSubTitleForeColor As Long

Dim OAIS As Boolean

' Windows API declarations
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'--- Window Style Constants ---
Private Const GWL_STYLE As Long = -16
Private Const WS_CAPTION As Long = &HC00000

Private Sub UserForm_Initialize()
    
    CenterUserFormOnActiveMonitor Me
    
    ' Initialize the progress array (messages at different stages)
    progressText(0) = "Initializing CORE systems..."
    progressText(1) = "Loading AI Protocols..."
    progressText(2) = "Connecting to Reflections Session..."
    OAISConnect
    If OAIS = True Then
        progressText(3) = "Synchronizing Data Systems..."
        progressText(4) = "Finalizing Authorization Parameters..."
        progressText(5) = "AUTO_SPCL Ready for Action..."
    ElseIf OAIS = False Then
        progressText(3) = "Reflections Session not found...."
        progressText(4) = "Finalizing Authorization Parameters..."
        progressText(5) = "AUTO_SPCL Online but /// DEGRADED ///..."
    End If
    ' *** FIX 2: Store the final width of the progress bar (set in the designer) ***
    totalProgressBarWidth = lblProgressBar.Width
    originalSubTitleForeColor = lblSubTitle.ForeColor
    
    ' Set initial progress
    progress = 0
    lblProgressBar.Width = 0 ' Reset progress bar
End Sub

Private Sub OAISConnect()

    On Error GoTo EH

    Set iApp = GetObject(, "Attachmate_Reflection_Objects_Framework.ApplicationObject")
    Set iFrame = iApp.GetObject("Frame")
    Set iCT = iFrame.SelectedView.Control
    Set iCS = iCT.screen

    OAIS = True
    Exit Sub

EH:
    OAIS = False
End Sub
Private Sub UserForm_Activate()
    Dim hWnd As LongPtr
    Dim currentStyle As LongPtr
    Dim newStyle As LongPtr

    ' 1. Find the UserForm's window handle (hWnd)
    ' Note: This finds the form based on its *current* caption.
    hWnd = FindWindow("ThunderDFrame", Me.Caption)
    
    If hWnd = 0 Then Exit Sub ' Exit if window not found

    ' 2. Get the current style of the window
    ' *** CORRECTED CALL: Use GetWindowLong ***
    currentStyle = GetWindowLong(hWnd, GWL_STYLE)

    ' 3. Create the new style by removing the WS_CAPTION bit
    newStyle = currentStyle And (Not WS_CAPTION)

    ' 4. Apply the new, borderless style to the window
    ' *** CORRECTED CALL: Use SetWindowLong ***
    Call SetWindowLong(hWnd, GWL_STYLE, newStyle)
    
    ' 5. Force the window to redraw itself to show the change
    DrawMenuBar (hWnd)
    
    ' Start the progress animation AFTER the form is visible ***
    StartProgress
End Sub

Private Sub StartProgress()
    Dim i As Integer
    
    ' Loop to simulate progress bar increment
    For i = 1 To 100 Step 2 ' Increment by 2%
        progress = i
        
        ' *** FIX 2: Correct progress bar width calculation ***
        lblProgressBar.Width = (i / 100) * totalProgressBarWidth
        
        ' Update progress percentage label
        If i > 5 Then lblProgressBar.Caption = i & "%"
        
        ' Update the status text according to progress
        If i <= 20 Then
            lblSubTitle.Caption = progressText(0)
        ElseIf i <= 40 Then
            lblSubTitle.Caption = progressText(1)
        ElseIf i <= 60 Then
            lblSubTitle.Caption = progressText(2)
        ElseIf i <= 80 Then
            If OAIS = False Then lblSubTitle.ForeColor = vbRed
            lblSubTitle.Caption = progressText(3)
        ElseIf i < 100 Then
            lblSubTitle.Caption = progressText(4)
        Else
            ' This block is never reached because i never equals 100
            lblSubTitle.Caption = progressText(5)
        End If
        
        ' *** Use Sleep API for a 150ms non-blocking wait ***
        Sleep 150
        DoEvents ' Process events to keep the form responsive and update labels
        
    Next i
    
    ' --- THIS IS THE FIX ---
    ' Manually set to 100% since the loop stops at 99
    lblProgressBar.Width = totalProgressBarWidth
    lblProgressBar.Caption = "100%"
    lblSubTitle.Caption = progressText(5) ' Set the final message
    DoEvents ' Ensure the form redraws at 100%
    ' -----------------------
    
    ' Give a moment to see the "100%" before closing
    Sleep 1000

    modStartupForm.HandleSplashComplete
    ResetSplashUiState
    Unload Me
End Sub

Private Sub ResetSplashUiState()
    ' Reset any UI state that may have been changed during the splash sequence
    lblSubTitle.ForeColor = originalSubTitleForeColor
End Sub
