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
Private Const MSO_FILE_DIALOG_FILE_PICKER As Long = 3

Private titleBarHidden As Boolean
Private mOriginalBodyTemplate As String
Private mSelectedMemberIndex As Long
Private mTemplateAttachmentEntries As Collection
Private mTemplateAttachmentLookup As Object
Private mUserAttachmentEntries As Collection
Private mUserAttachmentLookup As Object
Private mCurrentTemplateKey As String

Private Type MemberRecord
    Name As String
    SSN As String
    Status As String
End Type

Private mMembers() As MemberRecord
Private mMemberCount As Long
Private mMembersLoaded As Boolean
Private mFirstVisibleMemberIndex As Long

Private Const MEMBERS_PER_PAGE As Long = 8
Private Const DEFAULT_EMAIL_STATUS As String = "Draft"

Private Sub UserForm_Initialize()
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    On Error GoTo CleanFail

    SetCursorWait

    Dim templateKey As String

    titleBarHidden = False

    Me.bSUB.Visible = False

    On Error Resume Next
    templateKey = Trim$(Me.cboTemplate.value)
    On Error GoTo CleanFail

    If LenB(templateKey) = 0 Then
        templateKey = Trim$(Me.txtTEMP.value)
    End If

    If LenB(templateKey) > 0 Then
        LoadEmailTemplateData templateKey, _
                              Me.txtTO, Me.txtCC, Me.lstAT, _
                              Me.txtSubj, Me.txtBody, Me.txtSignature
    End If

    InitializeAttachmentTracking templateKey

    LoadMemberRecords

    mOriginalBodyTemplate = Me.txtBody.value

    mFirstVisibleMemberIndex = 1

    If mMemberCount > 0 Then
        SelectedMemberIndex = 1
    Else
        mSelectedMemberIndex = 0
        RenderMemberPage
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
    EnsureMemberRecordsLoaded

    If mMemberCount = 0 Then
        mSelectedMemberIndex = 0
        Exit Property
    End If

    If value < 1 Then value = 1
    If value > mMemberCount Then value = mMemberCount

    mSelectedMemberIndex = value

    EnsureSelectedIndexVisible
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

Private Sub EnsureMemberRecordsLoaded()
    If mMembersLoaded Then Exit Sub
    LoadMemberRecords
End Sub

Private Sub LoadMemberRecords()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim data As Variant
    Dim rowCount As Long
    Dim idx As Long

    mMembersLoaded = True
    mMemberCount = 0
    Erase mMembers
    mFirstVisibleMemberIndex = 1

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ID")
    On Error GoTo 0

    If ws Is Nothing Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
    If lastRow < 2 Then Exit Sub

    data = ws.Range("A2:B" & lastRow).value
    If Not IsArray(data) Then Exit Sub

    rowCount = UBound(data, 1)
    If rowCount <= 0 Then Exit Sub

    ReDim mMembers(1 To rowCount)
    For idx = 1 To rowCount
        mMembers(idx).SSN = SafeText(data(idx, 1))
        mMembers(idx).Name = SafeText(data(idx, 2))
        mMembers(idx).Status = DEFAULT_EMAIL_STATUS
    Next idx

    mMemberCount = rowCount
End Sub

Private Sub RenderMemberPage()
    Dim slotIndex As Long
    Dim memberIndex As Long
    Dim nameLabel As MSForms.Label
    Dim ssnLabel As MSForms.Label
    Dim statusLabel As MSForms.Label

    EnsureMemberRecordsLoaded

    For slotIndex = 1 To MEMBERS_PER_PAGE
        memberIndex = mFirstVisibleMemberIndex + slotIndex - 1

        Set nameLabel = GetLabelControl("lblNM", slotIndex)
        Set ssnLabel = GetLabelControl("lblSSN", slotIndex)
        Set statusLabel = GetLabelControl("lblSTAT", slotIndex)

        If memberIndex >= 1 And memberIndex <= mMemberCount Then
            If Not nameLabel Is Nothing Then nameLabel.caption = GetMemberNameValue(memberIndex)
            If Not ssnLabel Is Nothing Then ssnLabel.caption = GetMemberSSNValue(memberIndex)
            If Not statusLabel Is Nothing Then
                statusLabel.caption = GetMemberStatusValue(memberIndex)
                ApplyStatusColor statusLabel
            End If
        Else
            If Not nameLabel Is Nothing Then nameLabel.caption = ""
            If Not ssnLabel Is Nothing Then ssnLabel.caption = ""
            If Not statusLabel Is Nothing Then
                statusLabel.caption = ""
                ApplyStatusColor statusLabel
            End If
        End If
    Next slotIndex
End Sub

Private Function GetLabelControl(ByVal prefix As String, ByVal index As Long) As MSForms.Label
    Dim ctrl As MSForms.Control
    Dim controlName As String

    controlName = prefix & CStr(index)

    On Error Resume Next
    Set ctrl = Me.Controls(controlName)
    On Error GoTo 0

    If ctrl Is Nothing Then Exit Function
    If Not TypeOf ctrl Is MSForms.Label Then Exit Function

    Set GetLabelControl = ctrl
End Function

Private Sub EnsureSelectedIndexVisible()
    Dim maxStart As Long

    EnsureMemberRecordsLoaded

    If mMemberCount = 0 Then
        mFirstVisibleMemberIndex = 1
        RenderMemberPage
        Exit Sub
    End If

    If mFirstVisibleMemberIndex < 1 Then mFirstVisibleMemberIndex = 1

    maxStart = mMemberCount - MEMBERS_PER_PAGE + 1
    If maxStart < 1 Then maxStart = 1

    If mSelectedMemberIndex < mFirstVisibleMemberIndex Then
        mFirstVisibleMemberIndex = mSelectedMemberIndex
    ElseIf mSelectedMemberIndex > mFirstVisibleMemberIndex + MEMBERS_PER_PAGE - 1 Then
        mFirstVisibleMemberIndex = mSelectedMemberIndex - MEMBERS_PER_PAGE + 1
    End If

    If mFirstVisibleMemberIndex > maxStart Then mFirstVisibleMemberIndex = maxStart
    If mFirstVisibleMemberIndex < 1 Then mFirstVisibleMemberIndex = 1

    RenderMemberPage
End Sub

Private Sub ScrollMembers(ByVal delta As Long)
    Dim maxStart As Long
    Dim newStart As Long

    EnsureMemberRecordsLoaded

    If mMemberCount = 0 Then Exit Sub
    If delta = 0 Then Exit Sub

    maxStart = mMemberCount - MEMBERS_PER_PAGE + 1
    If maxStart < 1 Then maxStart = 1

    newStart = mFirstVisibleMemberIndex + delta
    If newStart < 1 Then newStart = 1
    If newStart > maxStart Then newStart = maxStart

    If newStart = mFirstVisibleMemberIndex Then Exit Sub

    mFirstVisibleMemberIndex = newStart
    RenderMemberPage

    If mSelectedMemberIndex < mFirstVisibleMemberIndex Then
        SelectedMemberIndex = mFirstVisibleMemberIndex
    ElseIf mSelectedMemberIndex > mFirstVisibleMemberIndex + MEMBERS_PER_PAGE - 1 Then
        SelectedMemberIndex = mFirstVisibleMemberIndex + MEMBERS_PER_PAGE - 1
    Else
        ApplyBodyPlaceholders mSelectedMemberIndex
    End If
End Sub

Private Function DisplayIndexToMemberIndex(ByVal displayIndex As Long) As Long
    Dim memberIndex As Long

    EnsureMemberRecordsLoaded

    If displayIndex < 1 Then Exit Function

    memberIndex = mFirstVisibleMemberIndex + displayIndex - 1
    If memberIndex < 1 Or memberIndex > mMemberCount Then Exit Function

    DisplayIndexToMemberIndex = memberIndex
End Function

Private Function MemberIndexToDisplayIndex(ByVal memberIndex As Long) As Long
    Dim displayIndex As Long

    displayIndex = memberIndex - mFirstVisibleMemberIndex + 1
    If displayIndex < 1 Or displayIndex > MEMBERS_PER_PAGE Then Exit Function

    MemberIndexToDisplayIndex = displayIndex
End Function

Private Function GetMemberNameValue(ByVal memberIndex As Long) As String
    If memberIndex < 1 Or memberIndex > mMemberCount Then Exit Function
    GetMemberNameValue = mMembers(memberIndex).Name
End Function

Private Function GetMemberSSNValue(ByVal memberIndex As Long) As String
    If memberIndex < 1 Or memberIndex > mMemberCount Then Exit Function
    GetMemberSSNValue = mMembers(memberIndex).SSN
End Function

Private Function GetMemberStatusValue(ByVal memberIndex As Long) As String
    Dim statusText As String

    If memberIndex < 1 Or memberIndex > mMemberCount Then Exit Function

    statusText = mMembers(memberIndex).Status
    If LenB(statusText) = 0 Then
        statusText = DEFAULT_EMAIL_STATUS
        mMembers(memberIndex).Status = statusText
    End If

    GetMemberStatusValue = statusText
End Function

Private Sub SetMemberStatus(ByVal memberIndex As Long, ByVal statusText As String, _
                             Optional ByVal updateUI As Boolean = True)
    Dim normalized As String
    Dim displayIndex As Long
    Dim statusLabel As MSForms.Label

    If memberIndex < 1 Or memberIndex > mMemberCount Then Exit Sub

    normalized = Trim$(statusText)
    If LenB(normalized) = 0 Then
        normalized = DEFAULT_EMAIL_STATUS
    End If

    mMembers(memberIndex).Status = normalized

    If Not updateUI Then Exit Sub

    displayIndex = MemberIndexToDisplayIndex(memberIndex)
    If displayIndex < 1 Then Exit Sub

    Set statusLabel = GetLabelControl("lblSTAT", displayIndex)
    If statusLabel Is Nothing Then Exit Sub

    statusLabel.caption = normalized
    ApplyStatusColor statusLabel
End Sub

Private Sub ApplyBodyPlaceholders(Optional ByVal memberIndex As Long = -1)
    Dim baseText As String
    Dim targetIndex As Long
    Dim placeholderPairs As Variant

    EnsureMemberRecordsLoaded

    baseText = mOriginalBodyTemplate
    If LenB(baseText) = 0 Then
        baseText = Me.txtBody.value
    End If

    If LenB(baseText) = 0 Then Exit Sub

    If mMemberCount = 0 Then Exit Sub

    If memberIndex < 1 Then
        If mSelectedMemberIndex < 1 Then
            targetIndex = 1
        Else
            targetIndex = mSelectedMemberIndex
        End If
    Else
        targetIndex = memberIndex
    End If

    If targetIndex < 1 Then
        targetIndex = 1
    ElseIf targetIndex > mMemberCount Then
        targetIndex = mMemberCount
    End If

    mSelectedMemberIndex = targetIndex

    placeholderPairs = BuildPlaceholderPairs(targetIndex)
    Me.txtBody.value = ReplacePlaceholdersArray(baseText, placeholderPairs)
End Sub

Private Function BuildPlaceholderPairs(ByVal memberIndex As Long) As Variant
    Dim placeholders As Object
    Dim idx As Long
    Dim textValue As String
    Dim issues As Object
    Dim keys As Variant
    Dim key As Variant
    Dim arr() As Variant
    Dim nextSlot As Long

    EnsureMemberRecordsLoaded

    If mMemberCount = 0 Then
        BuildPlaceholderPairs = Array()
        Exit Function
    End If

    If memberIndex < 1 Or memberIndex > mMemberCount Then
        memberIndex = 1
    End If

    Set placeholders = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    placeholders.CompareMode = vbTextCompare
    On Error GoTo 0

    For idx = 1 To mMemberCount
        textValue = SafeText(GetMemberNameValue(idx))
        AddPlaceholderValue placeholders, "NAME" & CStr(idx), textValue
        If idx = memberIndex Then
            AddPlaceholderValue placeholders, "NAME", textValue
            AddPlaceholderValue placeholders, "MEMBERNAME", textValue
            AddPlaceholderValue placeholders, "SELECTEDNAME", textValue
            AddPlaceholderValue placeholders, "PRIMARYNAME", textValue
            AddPlaceholderValue placeholders, "CURRENTNAME", textValue
        End If

        textValue = SafeText(GetMemberSSNValue(idx))
        AddPlaceholderValue placeholders, "SSN" & CStr(idx), textValue
        If idx = memberIndex Then
            AddPlaceholderValue placeholders, "SSN", textValue
            AddPlaceholderValue placeholders, "MEMBERSSN", textValue
            AddPlaceholderValue placeholders, "SELECTEDSSN", textValue
            AddPlaceholderValue placeholders, "CURRENTSSN", textValue
        End If

        textValue = SafeText(GetMemberStatusValue(idx))
        AddPlaceholderValue placeholders, "STAT" & CStr(idx), textValue
        AddPlaceholderValue placeholders, "STATUS" & CStr(idx), textValue
        If idx = memberIndex Then
            AddPlaceholderValue placeholders, "STAT", textValue
            AddPlaceholderValue placeholders, "STATUS", textValue
            AddPlaceholderValue placeholders, "MEMBERSTATUS", textValue
            AddPlaceholderValue placeholders, "SELECTEDSTATUS", textValue
            AddPlaceholderValue placeholders, "CURRENTSTATUS", textValue
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
    EnsureMemberRecordsLoaded

    DetermineMaxMemberIndex = mMemberCount
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

Private Sub lblUP_Click()
    ScrollMembers -1
End Sub

Private Sub lblDOWN_Click()
    ScrollMembers 1
End Sub

Private Sub ToggleEmailStatus(ByVal memberIndex As Long)
    Dim currentStatus As String
    Dim newStatus As String

    EnsureMemberRecordsLoaded

    If memberIndex < 1 Or memberIndex > mMemberCount Then Exit Sub

    currentStatus = GetMemberStatusValue(memberIndex)

    If StrComp(currentStatus, DEFAULT_EMAIL_STATUS, vbTextCompare) = 0 Then
        newStatus = "Cancel"
    Else
        newStatus = DEFAULT_EMAIL_STATUS
    End If

    SetMemberStatus memberIndex, newStatus, True
End Sub

Private Sub ApplyStatusColor(ByVal statusLabel As MSForms.Label)
    Dim statusText As String

    If statusLabel Is Nothing Then Exit Sub

    statusText = Trim$(statusLabel.caption)

    If StrComp(statusText, DEFAULT_EMAIL_STATUS, vbTextCompare) = 0 Then
        statusLabel.ForeColor = vbGreen
    ElseIf StrComp(statusText, "Cancel", vbTextCompare) = 0 Then
        statusLabel.ForeColor = vbRed
    Else
        statusLabel.ForeColor = vbBlack
    End If
End Sub

Private Function GetListBoxEntries(ByVal listControl As MSForms.ListBox) As Collection
    Dim entries As Collection
    Dim idx As Long

    Set entries = New Collection
    If listControl Is Nothing Then
        Set GetListBoxEntries = entries
        Exit Function
    End If

    For idx = 0 To listControl.ListCount - 1
        entries.Add CStr(listControl.List(idx))
    Next idx

    Set GetListBoxEntries = entries
End Function

Private Sub LoadListBoxFromCollection(ByVal listControl As MSForms.ListBox, _
                                      ByVal entries As Collection)
    Dim entry As Variant

    If listControl Is Nothing Then Exit Sub

    listControl.Clear

    If entries Is Nothing Then Exit Sub

    For Each entry In entries
        listControl.AddItem CStr(entry)
    Next entry
End Sub

Private Sub InitializeAttachmentTracking(ByVal templateKey As String)
    Dim entry As Variant
    Dim normalizedKey As String

    mCurrentTemplateKey = templateKey

    Set mTemplateAttachmentEntries = GetTemplateAttachmentEntriesForKey(templateKey)
    If mTemplateAttachmentEntries Is Nothing Then
        Set mTemplateAttachmentEntries = New Collection
    End If

    Set mUserAttachmentEntries = GetUserAttachmentEntries(templateKey)
    If mUserAttachmentEntries Is Nothing Then
        Set mUserAttachmentEntries = New Collection
    End If

    Set mTemplateAttachmentLookup = CreateCaseInsensitiveDictionary()
    Set mUserAttachmentLookup = CreateCaseInsensitiveDictionary()

    If Not mTemplateAttachmentLookup Is Nothing Then
        For Each entry In mTemplateAttachmentEntries
            normalizedKey = NormalizeTemplateAttachmentEntry(CStr(entry))
            If LenB(normalizedKey) = 0 Then GoTo NextTemplateEntry
            If Not mTemplateAttachmentLookup.Exists(normalizedKey) Then
                mTemplateAttachmentLookup(normalizedKey) = True
            End If
NextTemplateEntry:
        Next entry
    End If

    If Not mUserAttachmentLookup Is Nothing Then
        For Each entry In mUserAttachmentEntries
            normalizedKey = NormalizeTemplateAttachmentEntry(CStr(entry))
            If LenB(normalizedKey) = 0 Then GoTo NextUserEntry
            If Not mUserAttachmentLookup.Exists(normalizedKey) Then
                mUserAttachmentLookup(normalizedKey) = True
            End If
NextUserEntry:
        Next entry
    End If

    RefreshAttachmentListDisplay
End Sub

Private Function GetAttachmentEntriesFromListBox(ByRef listBox As MSForms.ListBox) As Collection
    Dim entries As Collection
    Dim idx As Long
    Dim entryText As String

    Set entries = New Collection

    If listBox Is Nothing Then
        Set GetAttachmentEntriesFromListBox = entries
        Exit Function
    End If

    If listBox.ListCount = 0 Then
        Set GetAttachmentEntriesFromListBox = entries
        Exit Function
    End If

    For idx = 0 To listBox.ListCount - 1
        On Error Resume Next
        entryText = Trim$(CStr(listBox.List(idx)))
        If Err.Number <> 0 Then
            entryText = vbNullString
            Err.Clear
        End If
        On Error GoTo 0
        If LenB(entryText) > 0 Then
            entries.Add entryText
        End If
    Next idx

    Set GetAttachmentEntriesFromListBox = entries
End Function

Private Sub EnsureAttachmentTracking(ByVal templateKey As String)
    If StrComp(Trim$(mCurrentTemplateKey), Trim$(templateKey), vbTextCompare) <> 0 Then
        InitializeAttachmentTracking templateKey
        Exit Sub
    End If

    If mTemplateAttachmentEntries Is Nothing Then
        InitializeAttachmentTracking templateKey
        Exit Sub
    End If

    If mUserAttachmentEntries Is Nothing Then
        Set mUserAttachmentEntries = New Collection
    End If

    If mTemplateAttachmentLookup Is Nothing Then
        Set mTemplateAttachmentLookup = CreateCaseInsensitiveDictionary()
        RebuildLookupFromCollection mTemplateAttachmentLookup, mTemplateAttachmentEntries
    End If

    If mUserAttachmentLookup Is Nothing Then
        Set mUserAttachmentLookup = CreateCaseInsensitiveDictionary()
        RebuildLookupFromCollection mUserAttachmentLookup, mUserAttachmentEntries
    End If
End Sub

Private Sub RebuildLookupFromCollection(ByRef dict As Object, ByVal entries As Collection)
    Dim entry As Variant
    Dim normalizedKey As String

    If dict Is Nothing Then Exit Sub
    If entries Is Nothing Then Exit Sub

    For Each entry In entries
        normalizedKey = NormalizeTemplateAttachmentEntry(CStr(entry))
        If LenB(normalizedKey) > 0 Then
            If Not dict.Exists(normalizedKey) Then
                dict(normalizedKey) = True
            End If
        End If
    Next entry
End Sub

Private Function CreateCaseInsensitiveDictionary() As Object
    Dim dict As Object

    On Error Resume Next
    Set dict = CreateObject("Scripting.Dictionary")
    If Not dict Is Nothing Then
        dict.CompareMode = vbTextCompare
    End If
    On Error GoTo 0

    Set CreateCaseInsensitiveDictionary = dict
End Function

Private Function AddUserAttachmentFromPath(ByVal filePath As String) As Boolean
    Dim normalizedKey As String
    Dim entry As String
    Dim resolvedPath As String
    Dim displayName As String

    resolvedPath = Trim$(filePath)
    displayName = vbNullString

    If Not CheckIfAttachmentExists(displayName, resolvedPath) Then Exit Function

    normalizedKey = NormalizeTemplateAttachmentPath(resolvedPath)
    If LenB(normalizedKey) = 0 Then Exit Function

    If AttachmentExistsInTemplate(normalizedKey) Then Exit Function
    If AttachmentExistsInUser(normalizedKey) Then Exit Function

    entry = BuildAttachmentEntryFromComponents(displayName, resolvedPath)
    If LenB(entry) = 0 Then Exit Function

    If mUserAttachmentEntries Is Nothing Then
        Set mUserAttachmentEntries = New Collection
    End If
    mUserAttachmentEntries.Add entry

    If Not mUserAttachmentLookup Is Nothing Then
        mUserAttachmentLookup(normalizedKey) = True
    End If

    AddUserAttachmentFromPath = True
End Function

Private Function BuildUserAttachmentPaths() As Collection
    Dim paths As Collection
    Dim entry As Variant
    Dim trimmedEntry As String
    Dim separatorPos As Long
    Dim filePath As String

    If mUserAttachmentEntries Is Nothing Then Exit Function
    If mUserAttachmentEntries.Count = 0 Then Exit Function

    For Each entry In mUserAttachmentEntries
        trimmedEntry = Trim$(CStr(entry))
        If LenB(trimmedEntry) = 0 Then GoTo NextEntry

        separatorPos = InStr(trimmedEntry, "|")
        If separatorPos > 0 Then
            filePath = Trim$(Mid$(trimmedEntry, separatorPos + 1))
        Else
            filePath = trimmedEntry
        End If

        If LenB(filePath) = 0 Then GoTo NextEntry

        If paths Is Nothing Then
            Set paths = New Collection
        End If
        paths.Add filePath

NextEntry:
    Next entry

    Set BuildUserAttachmentPaths = paths
End Function

Private Function ResolveInitialAttachmentDialogPath(ByVal userPaths As Collection) As String
    Dim entry As Variant
    Dim candidate As String
    Dim resolved As String
    Dim candidateResult As String

    If userPaths Is Nothing Then Exit Function

    For Each entry In userPaths
        candidate = Trim$(CStr(entry))
        If LenB(candidate) = 0 Then GoTo NextEntry

        If LenB(resolved) = 0 Then
            resolved = candidate
        End If

        On Error Resume Next
        Err.Clear
        candidateResult = Dir$(candidate, vbNormal)
        If Err.Number = 0 Then
            If LenB(candidateResult) > 0 Then
                ResolveInitialAttachmentDialogPath = candidate
                On Error GoTo 0
                Exit Function
            End If
        Else
            Err.Clear
        End If
        On Error GoTo 0
NextEntry:
    Next entry

    ResolveInitialAttachmentDialogPath = resolved
End Function

Private Function RemoveUserAttachmentFromPath(ByVal filePath As String) As Boolean
    Dim normalizedKey As String

    normalizedKey = NormalizeTemplateAttachmentPath(filePath)
    If LenB(normalizedKey) = 0 Then Exit Function

    If AttachmentExistsInTemplate(normalizedKey) Then Exit Function
    If Not AttachmentExistsInUser(normalizedKey) Then Exit Function

    RemoveUserAttachmentByNormalizedKey normalizedKey
    RemoveUserAttachmentFromPath = True
End Function

Private Sub RemoveUserAttachmentByNormalizedKey(ByVal normalizedKey As String)
    Dim filtered As Collection
    Dim entry As Variant
    Dim entryKey As String
    Dim removed As Boolean

    If mUserAttachmentEntries Is Nothing Then Exit Sub

    Set filtered = New Collection

    For Each entry In mUserAttachmentEntries
        entryKey = NormalizeTemplateAttachmentEntry(CStr(entry))
        If LenB(entryKey) = 0 Then
            filtered.Add CStr(entry)
        ElseIf Not removed And StrComp(entryKey, normalizedKey, vbTextCompare) = 0 Then
            removed = True
        Else
            filtered.Add CStr(entry)
        End If
    Next entry

    Set mUserAttachmentEntries = filtered

    If Not mUserAttachmentLookup Is Nothing Then
        If mUserAttachmentLookup.Exists(normalizedKey) Then
            mUserAttachmentLookup.Remove normalizedKey
        End If
    End If
End Sub

Private Function AttachmentExistsInTemplate(ByVal normalizedKey As String) As Boolean
    If LenB(normalizedKey) = 0 Then Exit Function

    If Not mTemplateAttachmentLookup Is Nothing Then
        AttachmentExistsInTemplate = mTemplateAttachmentLookup.Exists(normalizedKey)
        Exit Function
    End If

    AttachmentExistsInTemplate = CollectionContainsNormalized(mTemplateAttachmentEntries, normalizedKey)
End Function

Private Function AttachmentExistsInUser(ByVal normalizedKey As String) As Boolean
    If LenB(normalizedKey) = 0 Then Exit Function

    If Not mUserAttachmentLookup Is Nothing Then
        AttachmentExistsInUser = mUserAttachmentLookup.Exists(normalizedKey)
        Exit Function
    End If

    AttachmentExistsInUser = CollectionContainsNormalized(mUserAttachmentEntries, normalizedKey)
End Function

Private Function CollectionContainsNormalized(ByVal entries As Collection, _
                                              ByVal normalizedKey As String) As Boolean
    Dim entry As Variant
    Dim entryKey As String

    If entries Is Nothing Then Exit Function

    For Each entry In entries
        entryKey = NormalizeTemplateAttachmentEntry(CStr(entry))
        If LenB(entryKey) = 0 Then GoTo NextEntry
        If StrComp(entryKey, normalizedKey, vbTextCompare) = 0 Then
            CollectionContainsNormalized = True
            Exit Function
        End If
NextEntry:
    Next entry
End Function

Private Function BuildCombinedAttachmentEntries() As Collection
    Dim combined As Collection
    Dim entry As Variant

    Set combined = New Collection

    If Not mTemplateAttachmentEntries Is Nothing Then
        For Each entry In mTemplateAttachmentEntries
            combined.Add CStr(entry)
        Next entry
    End If

    If Not mUserAttachmentEntries Is Nothing Then
        For Each entry In mUserAttachmentEntries
            combined.Add CStr(entry)
        Next entry
    End If

    Set BuildCombinedAttachmentEntries = combined
End Function

Private Sub ApplyAttachmentUpdates(ByVal templateKey As String)
    Dim combined As Collection

    Set combined = BuildCombinedAttachmentEntries()
    PopulateAttachmentListControl Me.lstAT, combined
    UpdateAttachmentActionsVisibility combined
    PersistUserAttachmentsToWorksheet templateKey
End Sub

Private Sub PersistUserAttachmentsToWorksheet(ByVal templateKey As String)
    Dim activeTemplateKey As String

    activeTemplateKey = Trim$(templateKey)
    If LenB(activeTemplateKey) = 0 Then
        activeTemplateKey = Trim$(mCurrentTemplateKey)
    End If

    If LenB(activeTemplateKey) = 0 Then Exit Sub

    On Error Resume Next
    WriteUserAttachmentEntries activeTemplateKey, mUserAttachmentEntries
    On Error GoTo 0
End Sub

Private Sub PopulateAttachmentListControl(ByRef listBox As MSForms.ListBox, ByVal entries As Collection)
    Dim entry As Variant

    If listBox Is Nothing Then Exit Sub

    listBox.Clear

    If entries Is Nothing Then Exit Sub

    For Each entry In entries
        On Error Resume Next
        listBox.AddItem CStr(entry)
        On Error GoTo 0
    Next entry
End Sub

Private Sub RefreshAttachmentListDisplay()
    Dim combined As Collection

    Set combined = BuildCombinedAttachmentEntries()
    PopulateAttachmentListControl Me.lstAT, combined
    UpdateAttachmentActionsVisibility combined
End Sub

Private Sub UpdateAttachmentActionsVisibility(ByVal combined As Collection)
    Dim hasAttachments As Boolean

    hasAttachments = Not combined Is Nothing And combined.Count > 0

    On Error Resume Next
    Me.bSUB.Visible = hasAttachments
    On Error GoTo 0
End Sub

Private Sub bADD_Click()
    Dim fd As FileDialog
    Dim selectedPaths As Collection
    Dim selectedItem As Variant
    Dim templateKey As String
    Dim waitApplied As Boolean
    Dim addedCount As Long

    On Error GoTo CleanFail

    templateKey = Trim$(Me.cboTemplate.value)
    If LenB(templateKey) = 0 Then
        templateKey = Trim$(Me.txtTEMP.value)
    End If

    If LenB(templateKey) = 0 Then
        MsgBox "Please select a template before adding attachments.", vbExclamation
        GoTo CleanExit
    End If

    EnsureAttachmentTracking templateKey

    Set fd = Application.FileDialog(MSO_FILE_DIALOG_FILE_PICKER)
    If fd Is Nothing Then GoTo CleanExit

    With fd
        .title = "Select attachments to include"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "All Files", "*.*"
        If .Show <> -1 Then GoTo CleanExit
        If .SelectedItems.Count = 0 Then GoTo CleanExit
        Set selectedPaths = New Collection
        For Each selectedItem In .SelectedItems
            If LenB(CStr(selectedItem)) > 0 Then
                selectedPaths.Add CStr(selectedItem)
            End If
        Next selectedItem
    End With

    If selectedPaths Is Nothing Then GoTo CleanExit
    If selectedPaths.Count = 0 Then GoTo CleanExit

    For Each selectedItem In selectedPaths
        If AddUserAttachmentFromPath(CStr(selectedItem)) Then
            addedCount = addedCount + 1
        End If
    Next selectedItem

    If addedCount = 0 Then GoTo CleanExit

    SetCursorWait
    waitApplied = True

    ApplyAttachmentUpdates templateKey

CleanExit:
    If waitApplied Then SetCursorDefault
    Set fd = Nothing
    Set selectedPaths = Nothing
    Exit Sub

CleanFail:
    If waitApplied Then SetCursorDefault
    MsgBox "Unable to add attachments: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

Private Sub bSUB_Click()
    Dim fd As FileDialog
    Dim selectedPaths As Collection
    Dim selectedItem As Variant
    Dim templateKey As String
    Dim waitApplied As Boolean
    Dim removedCount As Long
    Dim userPaths As Collection
    Dim initialFileName As String
    Dim normalizedKey As String
    Dim ignoredCount As Long

    On Error GoTo CleanFail

    templateKey = Trim$(Me.cboTemplate.value)
    If LenB(templateKey) = 0 Then
        templateKey = Trim$(Me.txtTEMP.value)
    End If

    If LenB(templateKey) = 0 Then
        MsgBox "Please select a template before removing attachments.", vbExclamation
        GoTo CleanExit
    End If

    EnsureAttachmentTracking templateKey

    Set userPaths = BuildUserAttachmentPaths()
    If userPaths Is Nothing Or userPaths.Count = 0 Then
        MsgBox "There are no user-added attachments to remove.", vbInformation
        GoTo CleanExit
    End If

    initialFileName = ResolveInitialAttachmentDialogPath(userPaths)

    Set fd = Application.FileDialog(MSO_FILE_DIALOG_FILE_PICKER)
    If fd Is Nothing Then GoTo CleanExit

    With fd
        .title = "Select attachments to remove"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "All Files", "*.*"
        If LenB(initialFileName) > 0 Then
            .initialFileName = initialFileName
        End If
        If .Show <> -1 Then GoTo CleanExit
        If .SelectedItems.Count = 0 Then GoTo CleanExit
        Set selectedPaths = New Collection
        For Each selectedItem In .SelectedItems
            normalizedKey = NormalizeTemplateAttachmentPath(CStr(selectedItem))
            If LenB(normalizedKey) = 0 Then GoTo NextSelection
            If AttachmentExistsInUser(normalizedKey) Then
                selectedPaths.Add CStr(selectedItem)
            Else
                ignoredCount = ignoredCount + 1
            End If
NextSelection:
        Next selectedItem
    End With

    If selectedPaths Is Nothing Then GoTo CleanExit
    If selectedPaths.Count = 0 Then
        MsgBox "None of the selected files were added by this draft. Template attachments cannot be removed.", _
               vbInformation
        GoTo CleanExit
    End If

    For Each selectedItem In selectedPaths
        If RemoveUserAttachmentFromPath(CStr(selectedItem)) Then
            removedCount = removedCount + 1
        End If
    Next selectedItem

    If removedCount = 0 Then GoTo CleanExit

    SetCursorWait
    waitApplied = True

    ApplyAttachmentUpdates templateKey

    If ignoredCount > 0 Then
        MsgBox "Some selected files were ignored because they belong to the template and cannot be removed.", _
               vbInformation
    End If

CleanExit:
    If waitApplied Then SetCursorDefault
    Set fd = Nothing
    Set selectedPaths = Nothing
    Exit Sub

CleanFail:
    If waitApplied Then SetCursorDefault
    MsgBox "Unable to remove attachments: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

Private Sub bCF_Click()
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    On Error GoTo CleanFail

    SetCursorWait

    Unload Me
    ShowStartupFormOnce True

CleanExit:
    SetCursorDefault
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
    Exit Sub

CleanFail:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    MsgBox "Unable to cancel email creation: " & errDescription, vbCritical
    Resume CleanExit
End Sub

Private Sub bCFC_Click()
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    Dim whitelist As Object
    Dim templateKey As String
    Dim templateEntries As Collection
    Dim userEntries As Collection

    On Error GoTo CleanFail

    SetCursorWait

    Set whitelist = BuildDraftWhitelist()

    If whitelist Is Nothing Then
        SetCursorDefault
        MsgBox "All members are currently marked as Cancel. There are no Draft emails to create.", _
               vbInformation
        GoTo CleanExit
    End If

    templateKey = Trim$(Me.cboTemplate.value)
    If LenB(templateKey) = 0 Then
        templateKey = Trim$(Me.txtTEMP.value)
    End If

    EnsureAttachmentTracking templateKey

    Set templateEntries = mTemplateAttachmentEntries
    Set userEntries = mUserAttachmentEntries

    CreateDraftsFromID whitelist, templateKey, templateEntries, userEntries

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

    EnsureMemberRecordsLoaded

    maxIndex = DetermineMaxMemberIndex()
    If maxIndex < 1 Then
        Set BuildDraftWhitelist = Nothing
        Exit Function
    End If

    For idx = 1 To maxIndex
        statusCaption = GetMemberStatusValue(idx)
        If StrComp(statusCaption, DEFAULT_EMAIL_STATUS, vbTextCompare) = 0 Then
            key = "IDX:" & CStr(idx)
            If Not dict.Exists(key) Then
                dict.Add key, True
                draftCount = draftCount + 1
            Else
                dict(key) = True
            End If

            nameCaption = GetMemberNameValue(idx)
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
Private Sub HandleEmailToggleClick(ByVal memberIndex As Long)
    EnsureMemberRecordsLoaded

    If mMemberCount = 0 Then Exit Sub

    If memberIndex < 1 Then
        memberIndex = mFirstVisibleMemberIndex
    End If

    If memberIndex < 1 Then
        memberIndex = 1
    ElseIf memberIndex > mMemberCount Then
        memberIndex = mMemberCount
    End If

    SelectedMemberIndex = memberIndex
    ToggleEmailStatus memberIndex
    ApplyBodyPlaceholders mSelectedMemberIndex
End Sub

Private Sub bBE_Click()
    HandleEmailToggleClick SelectedMemberIndex
End Sub

Private Sub lblL1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HandleLabelMouseMove Me.lblL1
End Sub

Private Sub lblL1_Click()
    Dim memberIndex As Long

    memberIndex = DisplayIndexToMemberIndex(1)
    If memberIndex = 0 Then Exit Sub

    HandleEmailToggleClick memberIndex
End Sub

Private Sub lblL2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HandleLabelMouseMove Me.lblL2
End Sub

Private Sub lblL2_Click()
    Dim memberIndex As Long

    memberIndex = DisplayIndexToMemberIndex(2)
    If memberIndex = 0 Then Exit Sub

    HandleEmailToggleClick memberIndex
End Sub

Private Sub lblL3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HandleLabelMouseMove Me.lblL3
End Sub

Private Sub lblL3_Click()
    Dim memberIndex As Long

    memberIndex = DisplayIndexToMemberIndex(3)
    If memberIndex = 0 Then Exit Sub

    HandleEmailToggleClick memberIndex
End Sub

Private Sub lblL4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HandleLabelMouseMove Me.lblL4
End Sub

Private Sub lblL4_Click()
    Dim memberIndex As Long

    memberIndex = DisplayIndexToMemberIndex(4)
    If memberIndex = 0 Then Exit Sub

    HandleEmailToggleClick memberIndex
End Sub

Private Sub lblL5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HandleLabelMouseMove Me.lblL5
End Sub

Private Sub lblL5_Click()
    Dim memberIndex As Long

    memberIndex = DisplayIndexToMemberIndex(5)
    If memberIndex = 0 Then Exit Sub

    HandleEmailToggleClick memberIndex
End Sub

Private Sub lblL6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HandleLabelMouseMove Me.lblL6
End Sub

Private Sub lblL6_Click()
    Dim memberIndex As Long

    memberIndex = DisplayIndexToMemberIndex(6)
    If memberIndex = 0 Then Exit Sub

    HandleEmailToggleClick memberIndex
End Sub

Private Sub lblL7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HandleLabelMouseMove Me.lblL7
End Sub

Private Sub lblL7_Click()
    Dim memberIndex As Long

    memberIndex = DisplayIndexToMemberIndex(7)
    If memberIndex = 0 Then Exit Sub

    HandleEmailToggleClick memberIndex
End Sub

Private Sub lblL8_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HandleLabelMouseMove Me.lblL8
End Sub

Private Sub lblL8_Click()
    Dim memberIndex As Long

    memberIndex = DisplayIndexToMemberIndex(8)
    If memberIndex = 0 Then Exit Sub

    HandleEmailToggleClick memberIndex
End Sub


