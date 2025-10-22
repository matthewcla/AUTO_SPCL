Attribute VB_Name = "modOAIS"
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'Written By: LCDR Zach Brown
'Update: 17 May 17 - Rewrote Change Screen To Accomodate starting on same screen, issues previously with a desync
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

'Lowest Level Interaction with Reflection - This Replaces the Session. code
Public iCS As Attachmate_Reflection_Objects_Emulation_IbmHosts.IbmScreen
Public iFrame As Attachmate_Reflection_Objects.frame
Const tOut As Integer = 15000

Sub ConnectToRunningOAIS()
'Written By: LCDR Zach Brown
'Date: 23 Apr 17
'Purpose: This subroutine establishes the connection between excel and the reflection session
'This has not been tested with multiple reflection screens open
Dim iApp As Attachmate_Reflection_Objects_Framework.ApplicationObject
Dim iCT As Attachmate_Reflection_Objects_Emulation_IbmHosts.IbmTerminal
    
    Set iApp = GetObject(, "Attachmate_Reflection_Objects_Framework.ApplicationObject")
    Set iFrame = iApp.GetObject("Frame")
    Set iCT = iFrame.SelectedView.Control
    Set iCS = iCT.screen

End Sub

Sub FlightStuds432N()
    DSEL "432N"
End Sub

Sub Placement433U()
    DSEL "433U"
End Sub

Sub DSEL(newDesk As String)
Dim i As Long
Dim deskMatchFlg As Boolean
    
    ChangeScreen "DSEL"
    
    For i = 11 To 18
        If iCS.GetText(i, 5, Len(newDesk)) = newDesk Then
            deskMatchFlg = True
            Exit For
        End If
    Next i
    
    If deskMatchFlg Then
        GoScreen i, 2
        iCS.SendKeys "s"
        HitF6
    Else
        MsgBox "Couldn't Find Desk match, stopping the program"
        End
    End If


End Sub

Sub ChangeScreen(newScrn As String)
Dim rVal As Long
Dim colSt As Long
Dim currScreen As String

    currScreen = Trim(iCS.GetText(1, 2, 5))
    Select Case newScrn
        Case "BLTR", "FTEX"
            colSt = 3
        Case "FORW"
            GoFORW
            Exit Sub
        Case Else
            colSt = 2
    End Select

    GoScreen 19, 11
    With iCS
        .SendKeys newScrn
        .SendControlKey (ControlKeyCode_Transmit)
    
        If newScrn = currScreen Then
            rVal = .WaitForKeyboardEnabled(tOut, 0)
        Else
            rVal = .WaitForText1(tOut, newScrn, 1, colSt, TextComparisonOption_IgnoreCase)
        End If
    End With
    If (rVal <> ReturnCode_Success) Then
            Err.Raise 5001, "WaitForCursor1", "Timeout waiting for cursor position.", "VBAHelp.chm", "5001"
    End If
End Sub

Sub GoFORW()
Dim rVal As Long
    GoScreen 19, 11
    With iCS
        .SendKeys "FORW"
        .SendControlKey (ControlKeyCode_Transmit)
        rVal = .WaitForHostSettle(tOut, 0)
    End With
    If (rVal <> ReturnCode_Success) Then
        Err.Raise 5001, "WaitForHostSettle", "Timeout waiting for host to settle.", "VBAHelp.chm", "5001"
    End If
End Sub

Sub GoScreen(sRow As Long, SCol As Long)
Dim rVal As Long
    With iCS
        .MoveCursorTo1 sRow, SCol
        rVal = .WaitForCursor1(tOut, sRow, SCol)
    End With
    If (rVal <> ReturnCode_Success) Then
        Err.Raise 5001, "WaitForCursor1", "Timeout waiting for cursor position.", "VBAHelp.chm", "5001"
    End If
End Sub

Sub EnterUic(AUIC As String)
    GoScreen 4, 8
    iCS.SendKeys AUIC
    HitEnter
End Sub

Sub HitEnter()
Dim rVal As Long
    With iCS
        .SendControlKey (ControlKeyCode_Transmit)
        rVal = .WaitForKeyboardEnabled(tOut, 0)
    End With
    If (rVal <> ReturnCode_Success) Then
        Err.Raise 5001, "WaitForKeyBoardEnabled", "Timeout waiting for Keyboard Enabled.", "VBAHelp.chm", "5001"
    End If
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'Not Tested Below Here
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX



Sub HitF2()
Dim rVal As Long
    With iCS
        .SendControlKey (ControlKeyCode_F2)
        rVal = .WaitForKeyboardEnabled(tOut, 0)
    End With
    If (rVal <> ReturnCode_Success) Then
        Err.Raise 5001, "WaitForKeyBoardEnabled", "Timeout waiting for Keyboard Enabled.", "VBAHelp.chm", "5001"
    End If
End Sub

Sub HitF3()
Dim rVal As Long
    With iCS
        .SendControlKey (ControlKeyCode_F3)
        rVal = .WaitForKeyboardEnabled(tOut, 0)
    End With
    If (rVal <> ReturnCode_Success) Then
        Err.Raise 5001, "WaitForKeyBoardEnabled", "Timeout waiting for Keyboard Enabled.", "VBAHelp.chm", "5001"
    End If
End Sub

Sub HitF4()
 Dim rVal As Long
    With iCS
        .SendControlKey (ControlKeyCode_F4)
        rVal = .WaitForKeyboardEnabled(tOut, 0)
    End With
    If (rVal <> ReturnCode_Success) Then
        Err.Raise 5001, "WaitForKeyBoardEnabled", "Timeout waiting for Keyboard Enabled.", "VBAHelp.chm", "5001"
    End If
End Sub

Sub HitF6()
Dim rVal As Long
    With iCS
        .SendControlKey (ControlKeyCode_F6)
        rVal = .WaitForKeyboardEnabled(tOut, 0)
    End With
    If (rVal <> ReturnCode_Success) Then
        Err.Raise 5001, "WaitForKeyBoardEnabled", "Timeout waiting for Keyboard Enabled.", "VBAHelp.chm", "5001"
    End If
End Sub

Sub HitF7()
Dim rVal As Long
    With iCS
        .SendControlKey (ControlKeyCode_F7)
        rVal = .WaitForKeyboardEnabled(tOut, 0)
    End With
    If (rVal <> ReturnCode_Success) Then
        Err.Raise 5001, "WaitForKeyBoardEnabled", "Timeout waiting for Keyboard Enabled.", "VBAHelp.chm", "5001"
    End If
End Sub

Sub HitF8()
Dim rVal As Long
    With iCS
        .SendControlKey (ControlKeyCode_F8)
        rVal = .WaitForKeyboardEnabled(tOut, 0)
    End With
    If (rVal <> ReturnCode_Success) Then
        Err.Raise 5001, "WaitForKeyBoardEnabled", "Timeout waiting for Keyboard Enabled.", "VBAHelp.chm", "5001"
    End If
End Sub

Sub HitF9()
Dim rVal As Long
    With iCS
        .SendControlKey (ControlKeyCode_F9)
        rVal = .WaitForKeyboardEnabled(tOut, 0)
    End With
    If (rVal <> ReturnCode_Success) Then
        Err.Raise 5001, "WaitForKeyBoardEnabled", "Timeout waiting for Keyboard Enabled.", "VBAHelp.chm", "5001"
    End If
End Sub

Sub HitF10()
Dim rVal As Long
    With iCS
        .SendControlKey (ControlKeyCode_F10)
        rVal = .WaitForKeyboardEnabled(tOut, 0)
    End With
    If (rVal <> ReturnCode_Success) Then
        Err.Raise 5001, "WaitForKeyBoardEnabled", "Timeout waiting for Keyboard Enabled.", "VBAHelp.chm", "5001"
    End If
End Sub

Sub HitEnd()
Dim rVal As Long
    With iCS
        .SendControlKey (ControlKeyCode_Erase_Eof)
        rVal = .WaitForKeyboardEnabled(tOut, 0)
    End With
    If (rVal <> ReturnCode_Success) Then
        Err.Raise 5001, "WaitForKeyBoardEnabled", "Timeout waiting for Keyboard Enabled.", "VBAHelp.chm", "5001"
    End If
End Sub

Function GetScreen() As String
    GetScreen = iCS.GetText(1, 2, 4)
End Function

Sub sWrite(sRow As Long, sColumn As Long, textToWrite As String)
    GoScreen sRow, sColumn
    HitEnd
    If textToWrite = "" Then
        iCS.SendKeys ""
    Else
        iCS.SendKeys textToWrite
    End If
End Sub

Sub entText(sRow As Long, sColumn As Long, textToWrite As String)
    GoScreen sRow, sColumn
    HitEnd
    If textToWrite = "" Then
        iCS.SendKeys ""
    Else
        iCS.SendKeys textToWrite
    End If
    HitEnter
End Sub


Function CheckSSN(sRow As Long, sColumn As Long, ssnToCheck As String) As Boolean
    Dim screenSSN As String

    screenSSN = Trim(iCS.GetText(sRow, sColumn, 9))
    If ssnToCheck = screenSSN Then
        CheckSSN = True
    Else
        CheckSSN = False
    End If
End Function





