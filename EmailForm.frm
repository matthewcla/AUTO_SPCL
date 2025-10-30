VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EmailForm 
   ClientHeight    =   11760
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14355
   OleObjectBlob   =   "EmailForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EmailForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Private mOriginalBodyTemplate As String
Private mSelectedMemberIndex As Long

Private Sub UserForm_Initialize()
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    On Error GoTo CleanFail

    SetCursorWait

    Dim templateKey As String

    titleBarHidden = False

    On Error Resume Next
    templateKey = Trim$(Me.cboTemplate.value)
    On Error GoTo CleanFail

    If LenB(templateKey) = 0 Then
        templateKey = Trim$(Me.txtTEMP.value)
    End If

    If LenB(templateKey) > 0 Then
        LoadEmailTemplateData templateKey, _
                              Me.txtTO, Me.txtCC, Me.txtAT, _
                              Me.txtSubj, Me.txtBody, Me.txtSignature
    End If

    mSelectedMemberIndex = 1
    mOriginalBodyTemplate = Me.txtBody.value
    If LenB(mOriginalBodyTemplate) > 0 Then
        ApplyBodyPlaceholders mSelectedMemberIndex
    End If

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

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    On Error GoTo CleanFail

    SetCursorWait

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
End Sub

Private Sub UserForm_Click()

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

    originalCaption = Me.caption
    tempCaption = "email-" & Hex$(ObjPtr(Me))
    Me.caption = tempCaption

    hWnd = FindWindow("ThunderDFrame", tempCaption)
    Me.caption = originalCaption

    If hWnd = 0 Then Exit Sub

    currentStyle = GetWindowLong(hWnd, GWL_STYLE)
    newStyle = currentStyle And (Not WS_CAPTION)
    SetWindowLong hWnd, GWL_STYLE, newStyle
    DrawMenuBar hWnd

    titleBarHidden = True
End Sub

Public Property Get SelectedMemberIndex() As Long
    SelectedMemberIndex = mSelectedMemberIndex
End Property

Public Property Let SelectedMemberIndex(ByVal value As Long)
    If value < 1 Then value = 1
    mSelectedMemberIndex = value
    ApplyBodyPlaceholders mSelectedMemberIndex
End Property

Public Sub RefreshBodyPlaceholders(Optional ByVal memberIndex As Long = -1, _
                                   Optional ByVal resetTemplate As Boolean = False)
    If resetTemplate Or LenB(mOriginalBodyTemplate) = 0 Then
        mOriginalBodyTemplate = Me.txtBody.value
    End If
    ApplyBodyPlaceholders memberIndex
End Sub

Public Sub LoadBodyTemplate(ByVal templateText As String, Optional ByVal memberIndex As Long = -1)
    mOriginalBodyTemplate = templateText
    Me.txtBody.value = templateText
    ApplyBodyPlaceholders memberIndex
End Sub

Private Sub ApplyBodyPlaceholders(Optional ByVal memberIndex As Long = -1)
    Dim baseText As String
    Dim targetIndex As Long
    Dim placeholderPairs As Variant

    baseText = mOriginalBodyTemplate
    If LenB(baseText) = 0 Then
        baseText = Me.txtBody.value
    End If

    If LenB(baseText) = 0 Then Exit Sub

    If memberIndex < 1 Then
        If mSelectedMemberIndex < 1 Then
            targetIndex = 1
        Else
            targetIndex = mSelectedMemberIndex
        End If
    Else
        targetIndex = memberIndex
        mSelectedMemberIndex = targetIndex
    End If

    placeholderPairs = BuildPlaceholderPairs(targetIndex)
    Me.txtBody.value = ReplacePlaceholdersArray(baseText, placeholderPairs)
End Sub

Private Function BuildPlaceholderPairs(ByVal memberIndex As Long) As Variant
    Dim placeholders As Object
    Dim maxIndex As Long
    Dim idx As Long
    Dim textValue As String
    Dim issues As Object
    Dim keys As Variant
    Dim key As Variant
    Dim arr() As Variant
    Dim nextSlot As Long

    Set placeholders = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    placeholders.CompareMode = vbTextCompare
    On Error GoTo 0

    maxIndex = DetermineMaxMemberIndex()
    If maxIndex < 1 Then maxIndex = 1

    For idx = 1 To maxIndex
        If ControlExists("lblNM" & CStr(idx)) Then
            textValue = SafeText(GetLabelCaptionByName("lblNM" & CStr(idx)))
            AddPlaceholderValue placeholders, "NAME" & CStr(idx), textValue
            If idx = memberIndex Then
                AddPlaceholderValue placeholders, "NAME", textValue
                AddPlaceholderValue placeholders, "MEMBERNAME", textValue
                AddPlaceholderValue placeholders, "SELECTEDNAME", textValue
                AddPlaceholderValue placeholders, "PRIMARYNAME", textValue
                AddPlaceholderValue placeholders, "CURRENTNAME", textValue
            End If
        End If

        If ControlExists("lblSSN" & CStr(idx)) Then
            textValue = SafeText(GetLabelCaptionByName("lblSSN" & CStr(idx)))
            AddPlaceholderValue placeholders, "SSN" & CStr(idx), textValue
            If idx = memberIndex Then
                AddPlaceholderValue placeholders, "SSN", textValue
                AddPlaceholderValue placeholders, "MEMBERSSN", textValue
                AddPlaceholderValue placeholders, "SELECTEDSSN", textValue
                AddPlaceholderValue placeholders, "CURRENTSSN", textValue
            End If
        End If

        If ControlExists("lblSTAT" & CStr(idx)) Then
            textValue = SafeText(GetLabelCaptionByName("lblSTAT" & CStr(idx)))
            AddPlaceholderValue placeholders, "STAT" & CStr(idx), textValue
            AddPlaceholderValue placeholders, "STATUS" & CStr(idx), textValue
            If idx = memberIndex Then
                AddPlaceholderValue placeholders, "STAT", textValue
                AddPlaceholderValue placeholders, "STATUS", textValue
                AddPlaceholderValue placeholders, "MEMBERSTATUS", textValue
                AddPlaceholderValue placeholders, "SELECTEDSTATUS", textValue
                AddPlaceholderValue placeholders, "CURRENTSTATUS", textValue
            End If
        End If
    Next idx

    AddPlaceholderValue placeholders, "MEMBERINDEX", CStr(memberIndex)
    AddPlaceholderValue placeholders, "CURRENTINDEX", CStr(memberIndex)
    AddPlaceholderValue placeholders, "SELECTEDINDEX", CStr(memberIndex)

    Set issues = CollectIssueMap()
    If Not issues Is Nothing Then
        AddPlaceholderValue placeholders, "ISSUECOUNT", CStr(issues.Count)
        AddPlaceholderValue placeholders, "ISSUESCOUNT", CStr(issues.Count)
        AddPlaceholderValue placeholders, "TOTALISSUES", CStr(issues.Count)
        AddPlaceholderValue placeholders, "ISSUES", BuildIssuesSummary(issues, False)
        AddPlaceholderValue placeholders, "ISSUES_LIST", BuildIssuesSummary(issues, False)
        AddPlaceholderValue placeholders, "ISSUES_BULLETED", BuildIssuesSummary(issues, True)

        keys = issues.keys
        SortNumericKeys keys
        If IsArray(keys) Then
            For Each key In keys
                AddPlaceholderValue placeholders, "ISSUE" & CStr(key), SafeText(issues(key))
            Next key
        End If
    End If

    AddPlaceholderValue placeholders, "NEWLINE", vbCrLf
    AddPlaceholderValue placeholders, "LINEBREAK", vbCrLf
    AddPlaceholderValue placeholders, "BR", vbCrLf
    AddPlaceholderValue placeholders, "TAB", vbTab

    Dim ctrl As MSForms.Control
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is MSForms.Label Then
            textValue = SafeText(ctrl.caption)
            AddPlaceholderValue placeholders, ctrl.Name, textValue, False
            If Len(ctrl.Name) > 3 Then
                AddPlaceholderValue placeholders, Mid$(ctrl.Name, 4), textValue, False
            End If
        End If
    Next ctrl

    If placeholders.Count = 0 Then
        BuildPlaceholderPairs = Array()
        Exit Function
    End If

    ReDim arr(0 To placeholders.Count * 2 - 1)
    nextSlot = 0
    For Each key In placeholders.keys
        arr(nextSlot) = key
        arr(nextSlot + 1) = placeholders(key)
        nextSlot = nextSlot + 2
    Next key

    BuildPlaceholderPairs = arr
End Function

Private Function CollectIssueMap() As Object
    Dim dict As Object
    Dim ctrl As MSForms.Control
    Dim idx As Long
    Dim caption As String

    Set dict = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    dict.CompareMode = vbTextCompare
    On Error GoTo 0

    For Each ctrl In Me.Controls
        If TypeOf ctrl Is MSForms.Label Then
            idx = ExtractIndex(ctrl.Name, "lblL")
            If idx > 0 Then
                caption = SafeText(ctrl.caption)
                If LenB(caption) > 0 Then
                    dict(CStr(idx)) = caption
                End If
            End If
        End If
    Next ctrl

    Set CollectIssueMap = dict
End Function

Private Function BuildIssuesSummary(ByVal issues As Object, Optional ByVal includeBullet As Boolean = False) As String
    Dim keys As Variant
    Dim lower As Long
    Dim upper As Long
    Dim parts() As String
    Dim i As Long
    Dim entry As String

    If issues Is Nothing Then Exit Function
    If issues.Count = 0 Then Exit Function

    keys = issues.keys
    SortNumericKeys keys

    If Not IsArray(keys) Then Exit Function

    On Error Resume Next
    lower = LBound(keys)
    upper = UBound(keys)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    If upper < lower Then Exit Function

    ReDim parts(0 To upper - lower)
    For i = lower To upper
        entry = SafeText(issues(keys(i)))
        If LenB(entry) = 0 Then entry = ""
        If includeBullet Then
            entry = "- " & entry
        End If
        parts(i - lower) = entry
    Next i

    BuildIssuesSummary = Join(parts, vbCrLf)
End Function

Private Sub SortNumericKeys(ByRef keys As Variant)
    Dim lower As Long
    Dim upper As Long
    Dim i As Long
    Dim j As Long
    Dim temp As Variant

    If Not IsArray(keys) Then Exit Sub

    On Error Resume Next
    lower = LBound(keys)
    upper = UBound(keys)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    If upper <= lower Then Exit Sub

    For i = lower To upper - 1
        For j = i + 1 To upper
            If Val(keys(j)) < Val(keys(i)) Then
                temp = keys(i)
                keys(i) = keys(j)
                keys(j) = temp
            End If
        Next j
    Next i
End Sub

Private Function DetermineMaxMemberIndex() As Long
    Dim ctrl As MSForms.Control
    Dim maxIndex As Long
    Dim idx As Long

    maxIndex = 0
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is MSForms.Label Then
            idx = ExtractIndex(ctrl.Name, "lblNM")
            If idx > maxIndex Then maxIndex = idx
            idx = ExtractIndex(ctrl.Name, "lblSSN")
            If idx > maxIndex Then maxIndex = idx
            idx = ExtractIndex(ctrl.Name, "lblSTAT")
            If idx > maxIndex Then maxIndex = idx
        End If
    Next ctrl

    If maxIndex = 0 Then maxIndex = 1
    DetermineMaxMemberIndex = maxIndex
End Function

Private Function ExtractIndex(ByVal controlName As String, ByVal prefix As String) As Long
    Dim upperName As String
    Dim upperPrefix As String
    Dim i As Long
    Dim digits As String
    Dim ch As String

    If Len(prefix) = 0 Then Exit Function

    upperName = UCase$(controlName)
    upperPrefix = UCase$(prefix)

    If Left$(upperName, Len(upperPrefix)) <> upperPrefix Then Exit Function

    For i = Len(prefix) + 1 To Len(controlName)
        ch = Mid$(controlName, i, 1)
        If ch >= "0" And ch <= "9" Then
            digits = digits & ch
        ElseIf Len(digits) > 0 Then
            Exit For
        End If
    Next i

    If Len(digits) > 0 Then
        ExtractIndex = Val(digits)
    End If
End Function

Private Function GetLabelCaptionByName(ByVal controlName As String) As String
    Dim ctrl As Object

    On Error Resume Next
    Set ctrl = Me.Controls(controlName)
    If Not ctrl Is Nothing Then
        GetLabelCaptionByName = SafeText(ctrl.caption)
    End If
    On Error GoTo 0
End Function

Private Function ControlExists(ByVal controlName As String) As Boolean
    Dim ctrl As Object

    On Error Resume Next
    Set ctrl = Me.Controls(controlName)
    ControlExists = Not ctrl Is Nothing
    On Error GoTo 0
End Function

Private Sub AddPlaceholderValue(ByVal dict As Object, ByVal key As String, ByVal value As String, _
                                Optional ByVal overwrite As Boolean = True)
    Dim normalizedKey As String

    If dict Is Nothing Then Exit Sub

    normalizedKey = UCase$(Trim$(key))
    If LenB(normalizedKey) = 0 Then Exit Sub

    If dict.Exists(normalizedKey) Then
        If overwrite Then
            dict(normalizedKey) = value
        End If
    Else
        dict.Add normalizedKey, value
    End If
End Sub

Private Function SafeText(ByVal value As Variant) As String
    If IsError(value) Then Exit Function
    If IsNull(value) Then Exit Function
    SafeText = Trim$(CStr(value))
End Function

Private Sub HandleLabelMouseMove(ByVal target As MSForms.Label)
    If target Is Nothing Then Exit Sub

    target.BorderStyle = fmBorderStyleSingle

    If target.BorderStyle = fmBorderStyleSingle Then
        target.BorderColor = vbWhite
    Else
        target.BorderColor = vbRed
    End If
End Sub

Private Sub ToggleEmailStatus(ByVal memberIndex As Long)
    Dim statusLabelName As String
    Dim statusControl As MSForms.Control
    Dim statusLabel As MSForms.Label
    Dim currentStatus As String

    statusLabelName = "lblSTAT" & CStr(memberIndex)

    On Error Resume Next
    Set statusControl = Me.Controls(statusLabelName)
    On Error GoTo 0

    If statusControl Is Nothing Then Exit Sub
    If Not TypeOf statusControl Is MSForms.Label Then Exit Sub

    Set statusLabel = statusControl

    currentStatus = Trim$(statusLabel.caption)

    If StrComp(currentStatus, "Draft", vbTextCompare) = 0 Then
        statusLabel.caption = "Cancel"
    Else
        statusLabel.caption = "Draft"
    End If

    ApplyStatusColor statusLabel
End Sub

Private Sub ApplyStatusColor(ByVal statusLabel As MSForms.Label)
    Dim statusText As String

    If statusLabel Is Nothing Then Exit Sub

    statusText = Trim$(statusLabel.caption)

    If StrComp(statusText, "Draft", vbTextCompare) = 0 Then
        statusLabel.ForeColor = vbGreen
    ElseIf StrComp(statusText, "Cancel", vbTextCompare) = 0 Then
        statusLabel.ForeColor = vbRed
    Else
        statusLabel.ForeColor = vbBlack
    End If
End Sub

Private Sub bCFC_Click()
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    Dim whitelist As Object

    On Error GoTo CleanFail

    SetCursorWait

    Set whitelist = BuildDraftWhitelist()

    If whitelist Is Nothing Then
        SetCursorDefault
        MsgBox "All members are currently marked as Cancel. There are no Draft emails to create.", _
               vbInformation
        GoTo CleanExit
    End If

    CreateDraftsFromID whitelist

CleanExit:
    SetCursorDefault
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
    Exit Sub

CleanFail:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    MsgBox "Unable to finalize drafts: " & errDescription, vbCritical
    errNumber = 0
    Resume CleanExit
End Sub

Private Function BuildDraftWhitelist() As Object
    Dim dict As Object
    Dim maxIndex As Long
    Dim idx As Long
    Dim statusCaption As String
    Dim nameCaption As String
    Dim key As String
    Dim draftCount As Long

    Set dict = CreateObject("Scripting.Dictionary")
    If dict Is Nothing Then Exit Function
    On Error Resume Next
    dict.CompareMode = vbTextCompare
    On Error GoTo 0

    maxIndex = DetermineMaxMemberIndex()
    If maxIndex < 1 Then maxIndex = 1

    For idx = 1 To maxIndex
        statusCaption = GetLabelCaptionByName("lblSTAT" & CStr(idx))
        If StrComp(statusCaption, "Draft", vbTextCompare) = 0 Then
            key = "IDX:" & CStr(idx)
            If Not dict.Exists(key) Then
                dict.Add key, True
                draftCount = draftCount + 1
            Else
                dict(key) = True
            End If

            nameCaption = GetLabelCaptionByName("lblNM" & CStr(idx))
            If LenB(nameCaption) > 0 Then
                key = "NAME:" & NormalizeDraftWhitelistValue(nameCaption)
                dict(key) = True
            End If
        End If
    Next idx

    If draftCount = 0 Then
        Set BuildDraftWhitelist = Nothing
    Else
        Set BuildDraftWhitelist = dict
    End If
End Function

Private Function NormalizeDraftWhitelistValue(ByVal value As String) As String
    NormalizeDraftWhitelistValue = UCase$(Trim$(value))
End Function

Private Sub lblL1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HandleLabelMouseMove Me.lblL1
End Sub

Private Sub lblL1_Click()
    ToggleEmailStatus 1
End Sub

Private Sub lblL2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HandleLabelMouseMove Me.lblL2
End Sub

Private Sub lblL2_Click()
    ToggleEmailStatus 2
End Sub

Private Sub lblL3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HandleLabelMouseMove Me.lblL3
End Sub

Private Sub lblL3_Click()
    ToggleEmailStatus 3
End Sub

Private Sub lblL4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HandleLabelMouseMove Me.lblL4
End Sub

Private Sub lblL4_Click()
    ToggleEmailStatus 4
End Sub

Private Sub lblL5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HandleLabelMouseMove Me.lblL5
End Sub

Private Sub lblL5_Click()
    ToggleEmailStatus 5
End Sub

Private Sub lblL6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HandleLabelMouseMove Me.lblL6
End Sub

Private Sub lblL6_Click()
    ToggleEmailStatus 6
End Sub

Private Sub lblL7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HandleLabelMouseMove Me.lblL7
End Sub

Private Sub lblL7_Click()
    ToggleEmailStatus 7
End Sub

Private Sub lblL8_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HandleLabelMouseMove Me.lblL8
End Sub

Private Sub lblL8_Click()
    ToggleEmailStatus 8
End Sub

