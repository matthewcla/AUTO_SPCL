Option Explicit
'------------------------------------------
' AUTO_SPCL Refactor Update:
' EmailForm control map:
'   txtsubj - Subject input
'   lstAT - Attachment list
'------------------------------------------
'-------------------------------------------------------------------------------
' Form: EmailForm
' Role   : Primary workspace for reviewing member records and composing outbound
'          messages. Provides template selection, body preview, and attachment
'          management before drafts are created.
' Coordinates:
'   - Uses modEmailTemplates to load template content and persisted attachments.
'   - Calls modEmail helpers to clear fields, sync attachments, and create Outlook drafts.
'   - Receives hand-off from ProgressForm once the record review pipeline completes.
'-------------------------------------------------------------------------------

Private Const MSO_FILE_DIALOG_FILE_PICKER As Long = 3

Private mTitleBarHidden As Boolean
Private mOriginalBodyTemplate As String
Private mOriginalSubjectTemplate As String
Private mSelectedMemberIndex As Long
Private mTemplateAttachmentEntries As Collection
Private mTemplateAttachmentLookup As Object
Private mUserAttachmentEntries As Collection
Private mUserAttachmentLookup As Object
Private mCurrentTemplateKey As String
Private mActiveHoverLabel As MSForms.label

Private mtxtTEMP As MSForms.TextBox
Private mtxtTO As MSForms.TextBox
Private mTxtcc As MSForms.TextBox
Private mTxtsubj As MSForms.TextBox
Private mTxtbody As MSForms.TextBox
Private mLstAT As MSForms.listBox
Private mbADD As MSForms.CommandButton
Private mbSUB As MSForms.CommandButton

Private mTemplateFieldWarningsShown As Object
Private mTemplateAvailabilityWarningShown As Boolean
Private mIsLoading As Boolean
Private mInitialized As Boolean

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
Private Const ENABLE_TEMPLATE_TRACE As Boolean = False

Private Sub InitializeControlReferences()
    Set mtxtTEMP = TryGetTextBox("txtTEMP")
    Set mtxtTO = TryGetTextBox("txtTO")
    Set mTxtcc = TryGetTextBox("txtcc")
    Set mTxtsubj = TryGetTextBox("txtsubj")
    Set mTxtbody = TryGetTextBox("txtbody")
    Set mLstAT = TryGetListBox("lstAT")
    Set mbADD = TryGetButton("bADD")
    Set mbSUB = TryGetButton("bSUB")
End Sub

Private Function TryGetControl(ByVal controlName As String) As MSForms.control
    Dim ctrl As MSForms.control

    On Error Resume Next
    Set ctrl = Me.controls(controlName)
    On Error GoTo 0

    Set TryGetControl = ctrl
End Function

Private Function TryGetTextBox(ByVal controlName As String) As MSForms.TextBox
    Dim ctrl As MSForms.control

    Set ctrl = TryGetControl(controlName)
    If ctrl Is Nothing Then Exit Function
    If TypeOf ctrl Is MSForms.TextBox Then Set TryGetTextBox = ctrl
End Function

Private Function TryGetListBox(ByVal controlName As String) As MSForms.listBox
    Dim ctrl As MSForms.control

    Set ctrl = TryGetControl(controlName)
    If ctrl Is Nothing Then Exit Function
    If TypeOf ctrl Is MSForms.listBox Then Set TryGetListBox = ctrl
End Function

Private Function TryGetButton(ByVal controlName As String) As MSForms.CommandButton
    Dim ctrl As MSForms.control

    Set ctrl = TryGetControl(controlName)
    If ctrl Is Nothing Then Exit Function
    If TypeOf ctrl Is MSForms.CommandButton Then Set TryGetButton = ctrl
End Function

Private Function TryGetLabel(ByVal controlName As String) As MSForms.label
    Dim ctrl As MSForms.control

    Set ctrl = TryGetControl(controlName)
    If ctrl Is Nothing Then Exit Function
    If TypeOf ctrl Is MSForms.label Then Set TryGetLabel = ctrl
End Function

Private Sub FocusTemplateSelector()
    If Not mtxtTEMP Is Nothing Then
        modUIHelpers.FocusControl mtxtTEMP
    Else
        modUIHelpers.EnsureFormFocus Me
    End If
End Sub

Private Sub FocusAttachmentList()
    If Not mLstAT Is Nothing Then
        modUIHelpers.FocusControl mLstAT
    Else
        modUIHelpers.EnsureFormFocus Me
    End If
End Sub

Private Sub FocusComposerField()
    If Not mtxtTO Is Nothing Then
        modUIHelpers.FocusControl mtxtTO
    Else
        modUIHelpers.EnsureFormFocus Me
    End If
End Sub

Private Function GetLabelByDisplayIndex(ByVal displayIndex As Long) As MSForms.label
    Dim labelName As String

    labelName = "lblL" & CStr(displayIndex)
    Set GetLabelByDisplayIndex = TryGetLabel(labelName)
End Function

Private Sub UpdateToFieldFromHighlightedRecord()
    Dim ssnValue As String
    Dim selectedIndex As Long
    Dim displayIndex As Long
    Dim selectionLabel As MSForms.label
    Dim ssnLabel As MSForms.label

    EnsureMemberRecordsLoaded

    selectedIndex = mSelectedMemberIndex
    If selectedIndex >= 1 And selectedIndex <= mMemberCount Then
        ssnValue = SafeText(GetMemberSSNValue(selectedIndex))
        PopulateToFieldFromSSN ssnValue
        Exit Sub
    End If

    For displayIndex = 1 To MEMBERS_PER_PAGE
        Set selectionLabel = GetLabelByDisplayIndex(displayIndex)
        If selectionLabel Is Nothing Then GoTo nextSlot

        If selectionLabel.BorderColor = vbRed Then
            Set ssnLabel = GetLabelControl("lblSSN", displayIndex)
            If Not ssnLabel Is Nothing Then
                ssnValue = SafeText(ssnLabel.caption)
            Else
                ssnValue = vbNullString
            End If

            PopulateToFieldFromSSN ssnValue
            Exit Sub
        End If

nextSlot:
    Next displayIndex

    PopulateToFieldFromSSN vbNullString
End Sub

Private Sub PopulateToFieldFromSSN(ByVal ssnValue As String)
    Dim ws As Worksheet
    Dim searchRange As Range
    Dim foundCell As Range
    Dim recipients As String
    Dim columnIndex As Long
    Dim part As String

    If mtxtTO Is Nothing Then Exit Sub

    recipients = vbNullString
    ssnValue = Trim$(ssnValue)

    If LenB(ssnValue) > 0 Then
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets("ID")
        On Error GoTo 0

        If Not ws Is Nothing Then
            On Error Resume Next
            Set searchRange = ws.Columns(1)
            On Error GoTo 0

            If Not searchRange Is Nothing Then
                On Error Resume Next
                Set foundCell = searchRange.Find(What:=ssnValue, LookIn:=xlValues, _
                                                 LookAt:=xlWhole, SearchOrder:=xlByRows, _
                                                 SearchDirection:=xlNext, MatchCase:=False)
                On Error GoTo 0

                If Not foundCell Is Nothing Then
                    For columnIndex = 3 To 6
                        part = SafeText(ws.Cells(foundCell.row, columnIndex).value)
                        If LenB(part) > 0 Then
                            If LenB(recipients) > 0 Then recipients = recipients & ";"
                            recipients = recipients & part
                        End If
                    Next columnIndex
                End If
            End If
        End If
    End If

    SetTextBoxText mtxtTO, recipients
End Sub

Private Sub HandleLabelMouseMoveByIndex(ByVal displayIndex As Long)
    Dim target As MSForms.label

    Set target = GetLabelByDisplayIndex(displayIndex)
    If target Is Nothing Then Exit Sub

    HandleLabelMouseMove target
End Sub

Private Sub HandleLabelClickByIndex(ByVal displayIndex As Long)
    Dim memberIndex As Long
    Dim target As MSForms.label
    Dim resolvedDisplayIndex As Long

    memberIndex = DisplayIndexToMemberIndex(displayIndex)

    If memberIndex = 0 Then
        DeselectMemberSelectionLabels 0
        UpdateToFieldFromHighlightedRecord
        Exit Sub
    End If

    SelectedMemberIndex = memberIndex

    resolvedDisplayIndex = MemberIndexToDisplayIndex(memberIndex)

    If resolvedDisplayIndex < 1 Then
        DeselectMemberSelectionLabels 0
        resolvedDisplayIndex = displayIndex
    Else
        DeselectMemberSelectionLabels resolvedDisplayIndex
    End If

    Set target = GetLabelByDisplayIndex(resolvedDisplayIndex)
    If Not target Is Nothing Then
        If target.BorderStyle <> fmBorderStyleSingle Then
            target.BorderStyle = fmBorderStyleSingle
        End If

        If target.BorderColor <> vbRed Then
            target.BorderColor = vbRed
        End If
    End If

    RefreshSelectedMemberDetails memberIndex, resolvedDisplayIndex
End Sub

Private Sub UpdateIssuePlaceholderForDisplayIndex(ByVal displayIndex As Long)
    Dim nameLabel As MSForms.label
    Dim ssnLabel As MSForms.label
    Dim memberName As String
    Dim memberSSN As String
    Dim issueDescription As String

    Debug.Print "[EmailForm] UpdateIssuePlaceholder: displayIndex=" & displayIndex

    Set nameLabel = GetLabelControl("lblNM", displayIndex)
    If nameLabel Is Nothing Then
        Debug.Print "[EmailForm] UpdateIssuePlaceholder: lblNM" & displayIndex & " not found"
        Exit Sub
    End If

    memberName = Trim$(SafeText(nameLabel.caption))

    Set ssnLabel = GetLabelControl("lblSSN", displayIndex)
    If ssnLabel Is Nothing Then
        memberSSN = vbNullString
    Else
        memberSSN = Trim$(SafeText(ssnLabel.caption))
    End If

    Debug.Print "[EmailForm] UpdateIssuePlaceholder: memberName='" & memberName & "'"
    Debug.Print "[EmailForm] UpdateIssuePlaceholder: memberSSN='" & memberSSN & "'"

    If LenB(memberName) = 0 And LenB(memberSSN) = 0 Then
        Debug.Print "[EmailForm] UpdateIssuePlaceholder: member identifiers unavailable"
        Exit Sub
    End If

    issueDescription = GetIssuesFromRedBoard(memberName, memberSSN)
    If LenB(issueDescription) = 0 Then
        Debug.Print "[EmailForm] UpdateIssuePlaceholder: issue lookup returned empty result"
        Exit Sub
    End If

    Debug.Print "[EmailForm] UpdateIssuePlaceholder: issueText='" & issueDescription & "'"

    UpdateIssuePlaceholder issueDescription
End Sub

Private Function GetIssuesFromRedBoard(ByVal memberName As String, _
                                       ByVal memberSSN As String) As String
    Const DEFAULT_ISSUE_DESCRIPTION As String = "Issue details currently unavailable."

    Dim lo As ListObject
    Dim nameColumn As Range
    Dim ssnColumn As Range
    Dim issueColumn As Range
    Dim matchIndex As Variant
    Dim matchedRow As Long
    Dim resolvedName As String
    Dim resolvedSSN As String
    Dim issueDescription As String

    resolvedName = Trim$(SafeText(memberName))
    resolvedSSN = Trim$(SafeText(memberSSN))

    Set lo = TryGetListObject("RED_Board")
    If lo Is Nothing Then
        Debug.Print "[EmailForm] UpdateIssuePlaceholder: table 'RED_Board' not found"
        Exit Function
    End If

    On Error Resume Next
    Set nameColumn = lo.ListColumns(1).DataBodyRange
    Set ssnColumn = lo.ListColumns(2).DataBodyRange
    Set issueColumn = lo.ListColumns(3).DataBodyRange
    On Error GoTo 0

    If issueColumn Is Nothing Then
        Debug.Print "[EmailForm] UpdateIssuePlaceholder: issue column unavailable"
        Exit Function
    End If

    matchedRow = 0

    If LenB(resolvedSSN) > 0 Then
        If Not ssnColumn Is Nothing Then
            matchIndex = Application.Match(resolvedSSN, ssnColumn, 0)
            If IsError(matchIndex) Then
                Debug.Print "[EmailForm] UpdateIssuePlaceholder: Application.Match by SSN failed; attempting manual search"
                matchedRow = FindMemberRowIndexBySSN(resolvedSSN, ssnColumn)
            Else
                matchedRow = CLng(matchIndex)
            End If
        Else
            Debug.Print "[EmailForm] UpdateIssuePlaceholder: SSN column unavailable"
        End If
    End If

    If matchedRow = 0 Then
        If nameColumn Is Nothing Then
            Debug.Print "[EmailForm] UpdateIssuePlaceholder: name column unavailable"
        ElseIf LenB(resolvedName) > 0 Then
            matchIndex = Application.Match(resolvedName, nameColumn, 0)
            If IsError(matchIndex) Then
                Debug.Print "[EmailForm] UpdateIssuePlaceholder: Application.Match failed; attempting manual search"
                matchedRow = FindMemberRowIndex(resolvedName, nameColumn)
            Else
                matchedRow = CLng(matchIndex)
            End If
        Else
            Debug.Print "[EmailForm] UpdateIssuePlaceholder: member name empty; skipping name lookup"
        End If
    End If

    If matchedRow <= 0 Or matchedRow > issueColumn.Rows.Count Then
        Debug.Print "[EmailForm] UpdateIssuePlaceholder: member not found in table; using default description"
        issueDescription = DEFAULT_ISSUE_DESCRIPTION
    Else
        issueDescription = SafeText(issueColumn.Cells(matchedRow, 1).Value)
        If LenB(issueDescription) = 0 Then
            Debug.Print "[EmailForm] UpdateIssuePlaceholder: issue description empty; using default description"
            issueDescription = DEFAULT_ISSUE_DESCRIPTION
        End If
    End If

    GetIssuesFromRedBoard = issueDescription
End Function

Private Sub UpdateIssuePlaceholder(ByVal issueDescription As String)
    Const ISSUE_PLACEHOLDER As String = "{Issues}"

    Dim bodyControl As MSForms.TextBox
    Dim emailBody As String
    Dim normalizedTarget As String
    Dim position As Long
    Dim placeholderStart As Long
    Dim placeholderEnd As Long
    Dim candidate As String
    Dim normalizedCandidate As String
    Dim replacementOutcome As String
    Dim replacementsApplied As Long

    If mTxtbody Is Nothing Then
        Set bodyControl = TryGetTextBox("txtbody")
        If bodyControl Is Nothing Then
            Debug.Print "[EmailForm] UpdateIssuePlaceholder: txtbody control missing"
            Exit Sub
        End If
        Set mTxtbody = bodyControl
    End If

    emailBody = GetBodyText()
    Debug.Print "[EmailForm] UpdateIssuePlaceholder: bodyLength=" & Len(emailBody)
    Debug.Print "[EmailForm] UpdateIssuePlaceholder: bodyPreview='" & Left$(emailBody, 200) & "'"

    normalizedTarget = NormalizeBraceToken(ISSUE_PLACEHOLDER)
    replacementOutcome = "No changes applied"
    replacementsApplied = 0

    position = 1
    Do While position > 0
        placeholderStart = InStr(position, emailBody, "{", vbTextCompare)
        If placeholderStart = 0 Then Exit Do

        placeholderEnd = InStr(placeholderStart + 1, emailBody, "}", vbTextCompare)
        If placeholderEnd = 0 Then Exit Do

        candidate = Mid$(emailBody, placeholderStart, placeholderEnd - placeholderStart + 1)
        normalizedCandidate = NormalizeBraceToken(candidate)

        position = placeholderEnd + 1

        If LenB(normalizedCandidate) > 0 Then
            If StrComp(normalizedCandidate, normalizedTarget, vbBinaryCompare) = 0 Then
                Debug.Print "[EmailForm] UpdateIssuePlaceholder: placeholder match='" & candidate & "' at position=" & placeholderStart
                emailBody = Left$(emailBody, placeholderStart - 1) & issueDescription & Mid$(emailBody, placeholderEnd + 1)
                replacementsApplied = replacementsApplied + 1
                position = placeholderStart + Len(issueDescription)
            End If
        End If
    Loop

    If replacementsApplied > 0 Then
        If replacementsApplied = 1 Then
            replacementOutcome = "Placeholder replaced"
        Else
            replacementOutcome = "Placeholder replaced (" & replacementsApplied & " matches)"
        End If
    End If

    If replacementsApplied = 0 Then
        If LenB(issueDescription) > 0 Then
            If LenB(emailBody) > 0 Then
                emailBody = emailBody & vbNewLine & vbNewLine & "Issues:" & vbNewLine & issueDescription
            Else
                emailBody = "Issues:" & vbNewLine & issueDescription
            End If
            replacementOutcome = "Issues section appended"
            Debug.Print "[EmailForm] UpdateIssuePlaceholder: placeholder not found; appended Issues section"
        Else
            Debug.Print "[EmailForm] UpdateIssuePlaceholder: no issue text available; body left unchanged"
        End If
    End If

    Debug.Print "[EmailForm] UpdateIssuePlaceholder: replacementOutcome=" & replacementOutcome
    SetBodyText emailBody
End Sub

Private Function NormalizeBraceToken(ByVal token As String) As String
    Dim cleaned As String
    Dim innerValue As String

    cleaned = SafeText(token)
    cleaned = Replace(cleaned, vbCr, " ")
    cleaned = Replace(cleaned, vbLf, " ")
    cleaned = Replace(cleaned, vbTab, " ")
    cleaned = Trim$(cleaned)

    If LenB(cleaned) = 0 Then
        NormalizeBraceToken = vbNullString
        Exit Function
    End If

    If Left$(cleaned, 1) = "{" And Right$(cleaned, 1) = "}" Then
        innerValue = Mid$(cleaned, 2, Len(cleaned) - 2)
    Else
        innerValue = cleaned
    End If

    innerValue = Replace(innerValue, vbCr, " ")
    innerValue = Replace(innerValue, vbLf, " ")
    innerValue = Replace(innerValue, vbTab, " ")
    innerValue = Trim$(innerValue)

    If LenB(innerValue) = 0 Then
        NormalizeBraceToken = vbNullString
    Else
        NormalizeBraceToken = "{" & UCase$(innerValue) & "}"
    End If
End Function

Private Function FindMemberRowIndex(ByVal memberName As String, ByVal nameColumn As Range) As Long
    Dim cell As Range
    Dim normalizedTarget As String
    Dim normalizedCandidate As String
    Dim indexCounter As Long

    normalizedTarget = NormalizeDraftWhitelistValue(memberName)

    For Each cell In nameColumn.Cells
        indexCounter = indexCounter + 1
        normalizedCandidate = NormalizeDraftWhitelistValue(SafeText(cell.value))

        If StrComp(normalizedCandidate, normalizedTarget, vbBinaryCompare) = 0 Then
            FindMemberRowIndex = indexCounter
            Exit Function
        End If
    Next cell

    Debug.Print "[EmailForm] UpdateIssuePlaceholder: manual search did not find member '" & memberName & "'"
End Function

Private Function FindMemberRowIndexBySSN(ByVal memberSSN As String, ByVal ssnColumn As Range) As Long
    Dim cell As Range
    Dim normalizedTarget As String
    Dim normalizedCandidate As String
    Dim indexCounter As Long

    normalizedTarget = NormalizeSSNValue(memberSSN)
    If LenB(normalizedTarget) = 0 Then Exit Function

    For Each cell In ssnColumn.Cells
        indexCounter = indexCounter + 1
        normalizedCandidate = NormalizeSSNValue(SafeText(cell.Value))

        If LenB(normalizedCandidate) = 0 Then GoTo NextCell

        If StrComp(normalizedCandidate, normalizedTarget, vbBinaryCompare) = 0 Then
            FindMemberRowIndexBySSN = indexCounter
            Exit Function
        End If

NextCell:
    Next cell

    Debug.Print "[EmailForm] UpdateIssuePlaceholder: manual SSN search did not find member with SSN '" & memberSSN & "'"
End Function

Private Function NormalizeSSNValue(ByVal value As String) As String
    Dim cleaned As String

    cleaned = SafeText(value)
    cleaned = Replace$(cleaned, "-", vbNullString)
    cleaned = Replace$(cleaned, " ", vbNullString)

    NormalizeSSNValue = cleaned
End Function

Private Function TryGetListObject(ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set lo = ws.ListObjects(tableName)
        On Error GoTo 0

        If Not lo Is Nothing Then
            Set TryGetListObject = lo
            Exit Function
        End If
    Next ws
End Function

Private Sub DeselectMemberSelectionLabels(Optional ByVal exceptDisplayIndex As Long = 0)
    Dim slotIndex As Long
    Dim candidate As MSForms.label

    For slotIndex = 1 To MEMBERS_PER_PAGE
        Set candidate = GetLabelByDisplayIndex(slotIndex)
        If candidate Is Nothing Then GoTo NextLabel

        If exceptDisplayIndex >= 1 Then
            If slotIndex = exceptDisplayIndex Then GoTo NextLabel
        End If

        candidate.BorderColor = vbWhite
        ResetHoverLabel candidate

        If Not mActiveHoverLabel Is Nothing Then
            If candidate Is mActiveHoverLabel Then
                Set mActiveHoverLabel = Nothing
            End If
        End If

NextLabel:
    Next slotIndex
End Sub

Private Sub RefreshSelectedMemberDetails(ByVal memberIndex As Long, ByVal displayIndex As Long)
    Dim nameLabel As MSForms.label
    Dim ssnLabel As MSForms.label
    Dim statusLabel As MSForms.label

    EnsureMemberRecordsLoaded

    If memberIndex >= 1 And memberIndex <= mMemberCount Then
        If displayIndex >= 1 And displayIndex <= MEMBERS_PER_PAGE Then
            Set nameLabel = GetLabelControl("lblNM", displayIndex)
            If Not nameLabel Is Nothing Then
                nameLabel.caption = SafeText(GetMemberNameValue(memberIndex))
            End If

            Set ssnLabel = GetLabelControl("lblSSN", displayIndex)
            If Not ssnLabel Is Nothing Then
                ssnLabel.caption = SafeText(GetMemberSSNValue(memberIndex))
            End If

            Set statusLabel = GetLabelControl("lblSTAT", displayIndex)
            If Not statusLabel Is Nothing Then
                statusLabel.caption = SafeText(GetMemberStatusValue(memberIndex))
                ApplyStatusColor statusLabel
            End If
        End If
    End If

    UpdateToFieldFromHighlightedRecord
End Sub

Private Function EnsureRequiredControls() As Boolean
    Dim missing As Collection

    Set missing = New Collection

    If mtxtTEMP Is Nothing Then missing.Add "txtTEMP (TextBox)"
    If mtxtTO Is Nothing Then missing.Add "txtTO (TextBox)"
    If mTxtcc Is Nothing Then missing.Add "txtcc (TextBox)"
    If mTxtsubj Is Nothing Then missing.Add "txtsubj (TextBox)"
    If mTxtbody Is Nothing Then missing.Add "txtbody (TextBox)"
    If mLstAT Is Nothing Then missing.Add "lstAT (ListBox)"
    If mbADD Is Nothing Then missing.Add "bADD (CommandButton)"
    If mbSUB Is Nothing Then missing.Add "bSUB (CommandButton)"

    EnsureRequiredControls = missing.Count = 0

    If Not EnsureRequiredControls Then
        modEmailFormDiagnostics.ReportMissingControls missing
    End If
End Function

Private Function JoinCollectionString(ByVal items As Collection, Optional ByVal delimiter As String = ", ") As String
    Dim entry As Variant
    Dim buffer As String

    If items Is Nothing Then Exit Function
    For Each entry In items
        If LenB(buffer) > 0 Then buffer = buffer & delimiter
        buffer = buffer & CStr(entry)
    Next entry

    JoinCollectionString = buffer
End Function

Private Sub AddFailureReason(ByRef reasons As Collection, ByVal message As String)
    Dim existing As Variant

    message = Trim$(message)
    If LenB(message) = 0 Then Exit Sub

    If reasons Is Nothing Then Set reasons = New Collection

    For Each existing In reasons
        If StrComp(CStr(existing), message, vbTextCompare) = 0 Then Exit Sub
    Next existing

    reasons.Add message
End Sub

Private Function GetTextBoxText(ByVal target As MSForms.TextBox, Optional ByVal trimResult As Boolean = True) As String
    If target Is Nothing Then Exit Function

    If trimResult Then
        GetTextBoxText = Trim$(CStr(target.value))
    Else
        GetTextBoxText = CStr(target.value)
    End If
End Function

Private Sub SetTextBoxText(ByVal target As MSForms.TextBox, ByVal value As String)
    If target Is Nothing Then Exit Sub
    target.value = value
End Sub

Private Function GetBodyText() As String
    Dim normalized As String

    If Not mTxtbody Is Nothing Then
        normalized = CStr(mTxtbody.Value)
    Else
        On Error Resume Next
        normalized = CStr(Me.txtbody.Value)
        On Error GoTo 0
    End If

    normalized = Replace$(normalized, vbCrLf, vbLf)
    normalized = Replace$(normalized, vbCr, vbLf)
    GetBodyText = Replace$(normalized, vbLf, vbCrLf)
End Function

Private Sub SetBodyText(ByVal value As String)
    Dim normalized As String

    normalized = Replace$(value, vbCrLf, vbLf)
    normalized = Replace$(normalized, vbCr, vbLf)
    normalized = Replace$(normalized, vbLf, vbCrLf)

    If Not mTxtbody Is Nothing Then
        mTxtbody.Value = normalized
    Else
        On Error Resume Next
        Me.txtbody.Value = normalized
        On Error GoTo 0
    End If
End Sub

Private Sub EnsureTemplateWarningCache()
    If Not mTemplateFieldWarningsShown Is Nothing Then Exit Sub

    On Error Resume Next
    Set mTemplateFieldWarningsShown = CreateObject("Scripting.Dictionary")
    If Not mTemplateFieldWarningsShown Is Nothing Then
        mTemplateFieldWarningsShown.CompareMode = vbTextCompare
    End If
    On Error GoTo 0
End Sub

Private Function PopulateTemplateDropdown() As Collection
    Dim keys As Collection
    On Error Resume Next
    Set keys = modEmailTemplates.GetAvailableTemplateKeys()
    On Error GoTo 0

    UpdateTemplateAvailabilityState keys

    Set PopulateTemplateDropdown = keys
End Function

Private Function ResolveInitialTemplateKey(Optional ByVal templateKeys As Collection) As String
    Dim candidate As String

    candidate = GetTextBoxText(mtxtTEMP)

    If LenB(candidate) = 0 Then
        candidate = GetFirstTemplateKey(templateKeys)
    End If

    ResolveInitialTemplateKey = candidate
End Function

Private Function GetFirstTemplateKey(ByVal templateKeys As Collection) As String
    Dim entry As Variant

    If templateKeys Is Nothing Then Exit Function

    For Each entry In templateKeys
        If LenB(Trim$(CStr(entry))) > 0 Then
            GetFirstTemplateKey = Trim$(CStr(entry))
            Exit Function
        End If
    Next entry
End Function

Private Function TemplateKeyExists(ByVal templateKey As String, _
                                   ByVal templateKeys As Collection) As Boolean
    Dim entry As Variant

    templateKey = Trim$(templateKey)
    If LenB(templateKey) = 0 Then Exit Function
    If templateKeys Is Nothing Then Exit Function

    For Each entry In templateKeys
        If StrComp(Trim$(CStr(entry)), templateKey, vbTextCompare) = 0 Then
            TemplateKeyExists = True
            Exit Function
        End If
    Next entry
End Function

Private Sub LoadTemplate(ByVal templateKey As String)
    Dim normalizedKey As String
    Dim loadSucceeded As Boolean
    Dim toValue As String
    Dim ccValue As String
    Dim subjectValue As String
    Dim bodyValue As String
    Dim attachmentCount As Long
    Dim previousStatus As Variant
    Dim statusActive As Boolean
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    Dim requestedTemplateKey As String
    Dim resolvedTemplateKey As Variant
    Dim fallbackUsed As Boolean

    On Error GoTo CleanFail

    normalizedKey = Trim$(templateKey)

    If LenB(normalizedKey) = 0 Then
        modEmail.ClearEmailFields mtxtTO, mTxtcc, mTxtsubj, mTxtbody, _
                                  mLstAT, mbSUB
        mOriginalBodyTemplate = vbNullString
        mOriginalSubjectTemplate = vbNullString
        mCurrentTemplateKey = vbNullString
        SetTextBoxText mtxtTEMP, vbNullString
        TraceTemplateSelection normalizedKey, False, vbNullString, vbNullString, vbNullString, vbNullString, 0
        GoTo CleanExit
    End If

    previousStatus = Application.StatusBar
    Application.StatusBar = "Loading template '" & normalizedKey & "'..."
    statusActive = True
    modUIHelpers.SetCursorWait
    If modProgressUI.IsFormLoaded("ProgressForm") Then
        modProgressUI.Progress_Log "Loading template '" & normalizedKey & "'..."
    End If

    requestedTemplateKey = normalizedKey
    resolvedTemplateKey = vbNullString
    loadSucceeded = LoadEmailTemplateIntoControls(normalizedKey, _
                                                  mtxtTO, mTxtcc, mLstAT, _
                                                  mTxtsubj, mTxtbody, _
                                                  resolvedTemplateKey)

    If loadSucceeded Then
        If VarType(resolvedTemplateKey) = vbString Then
            If LenB(CStr(resolvedTemplateKey)) > 0 Then
                If StrComp(requestedTemplateKey, CStr(resolvedTemplateKey), vbTextCompare) <> 0 Then
                    fallbackUsed = True
                End If
                normalizedKey = CStr(resolvedTemplateKey)
            End If
        End If
    End If

    toValue = GetTextBoxText(mtxtTO, False)
    ccValue = GetTextBoxText(mTxtcc, False)
    subjectValue = GetTextBoxText(mTxtsubj, False)
    bodyValue = GetBodyText()
    attachmentCount = 0
    If loadSucceeded Then attachmentCount = GetAttachmentListCount()

    TraceTemplateSelection normalizedKey, loadSucceeded, toValue, ccValue, subjectValue, bodyValue, attachmentCount

    If Not loadSucceeded Then
        ShowTemplateLoadFailure requestedTemplateKey
        modEmail.ClearEmailFields mtxtTO, mTxtcc, mTxtsubj, mTxtbody, _
                                  mLstAT, mbSUB
        mOriginalBodyTemplate = vbNullString
        mOriginalSubjectTemplate = vbNullString
        mCurrentTemplateKey = vbNullString
        SetTextBoxText mtxtTEMP, vbNullString
        GoTo CleanExit
    End If

    InitializeAttachmentTracking normalizedKey
    mOriginalSubjectTemplate = GetTextBoxText(mTxtsubj, False)
    mOriginalBodyTemplate = GetBodyText()
    mCurrentTemplateKey = normalizedKey
    SetTextBoxText mtxtTEMP, normalizedKey

    If fallbackUsed Then
        Debug.Print "[EmailForm] Template '" & requestedTemplateKey & "' not found. Loaded '" & normalizedKey & "' instead."
        Application.StatusBar = "Template '" & requestedTemplateKey & "' not found. Default template loaded."
        statusActive = False
    End If

    If modProgressUI.IsFormLoaded("ProgressForm") Then
        modProgressUI.Progress_Log "Template '" & normalizedKey & "' loaded."
    End If

    ValidateLoadedTemplateFields normalizedKey
    TraceEmailFieldState "LoadTemplate", normalizedKey

CleanExit:
    If statusActive Then
        Application.StatusBar = previousStatus
    End If
    modUIHelpers.SetCursorDefault
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
    Exit Sub

CleanFail:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Resume CleanExit
End Sub

Private Sub UpdateTemplateAvailabilityState(Optional ByVal templateKeys As Collection)
    Dim hasTemplates As Boolean
    Dim combinedAttachments As Collection

    hasTemplates = CollectionHasItems(templateKeys)

    modUIHelpers.SetControlsEnabled Array(mtxtTO, mTxtcc, mTxtsubj, _
                                          mTxtbody, mLstAT, _
                                          mbADD), hasTemplates

    If hasTemplates Then
        mTemplateAvailabilityWarningShown = False
        Application.StatusBar = False
        Set combinedAttachments = modEmail.BuildAttachmentDisplayList(mTemplateAttachmentEntries, _
                                                                      mUserAttachmentEntries)
        modEmail.UpdateAttachmentRemoveButton mbSUB, combinedAttachments
    Else
        modEmail.ClearEmailFields mtxtTO, mTxtcc, mTxtsubj, mTxtbody, _
                                  mLstAT, mbSUB
        mOriginalBodyTemplate = vbNullString
        mOriginalSubjectTemplate = vbNullString
        SetTextBoxText mtxtTEMP, vbNullString
        mCurrentTemplateKey = vbNullString
        If Not mTemplateAvailabilityWarningShown Then
            Application.StatusBar = "No email templates available. Add template columns on the 'Email Templates' sheet to enable this form."
            Debug.Print "[EmailForm] No email templates available; user informed via status bar."
            FocusTemplateSelector
            mTemplateAvailabilityWarningShown = True
        End If
    End If
End Sub

Private Function CollectionHasItems(ByVal values As Collection) As Boolean
    If values Is Nothing Then Exit Function

    On Error Resume Next
    CollectionHasItems = (values.Count > 0)
    If Err.Number <> 0 Then
        Err.Clear
        CollectionHasItems = False
    End If
    On Error GoTo 0
End Function

Private Function ResolveActiveTemplateKey(Optional ByVal includeCurrent As Boolean = True) As String
    Dim templateKey As String

    templateKey = GetTextBoxText(mtxtTEMP)

    If LenB(templateKey) = 0 And includeCurrent Then
        templateKey = Trim$(mCurrentTemplateKey)
    End If

    ResolveActiveTemplateKey = templateKey
End Function

Public Function activeTemplateKey(Optional ByVal includeCurrent As Boolean = True) As String
    activeTemplateKey = ResolveActiveTemplateKey(includeCurrent)
End Function

Private Sub TraceTemplateSelection(ByVal templateKey As String, _
                                   ByVal loadSucceeded As Boolean, _
                                   ByVal toValue As String, _
                                   ByVal ccValue As String, _
                                   ByVal subjectValue As String, _
                                   ByVal bodyValue As String, _
                                   ByVal attachmentCount As Long)
    If Not ENABLE_TEMPLATE_TRACE Then Exit Sub

    Debug.Print "[EmailForm] Template '" & templateKey & "' load=" & loadSucceeded & _
                " TO='" & toValue & "' CC='" & ccValue & _
                "' Subject='" & subjectValue & "' BodyLen=" & Len(bodyValue) & _
                " Attachments=" & attachmentCount
End Sub

Private Sub TraceEmailFieldState(ByVal stage As String, ByVal templateKey As String)
    Dim toValue As String
    Dim ccValue As String
    Dim subjectValue As String
    Dim bodyValue As String
    Dim attachmentCount As Long

    If Not ENABLE_TEMPLATE_TRACE Then Exit Sub

    toValue = GetTextBoxText(mtxtTO, False)
    ccValue = GetTextBoxText(mTxtcc, False)
    subjectValue = GetTextBoxText(mTxtsubj, False)
    bodyValue = GetBodyText()
    attachmentCount = GetAttachmentListCount()

    Debug.Print "[EmailForm] State '" & stage & "' template='" & templateKey & _
                "' TO='" & toValue & "' CC='" & ccValue & "' Subject='" & subjectValue & _
                "' BodyLen=" & Len(bodyValue) & " Attachments=" & attachmentCount
End Sub

Private Sub ValidateLoadedTemplateFields(ByVal templateKey As String)
    Dim warnings As Collection
    Dim normalizedKey As String
    Dim warningKey As String
    Dim warningText As String

    Set warnings = New Collection

    Dim toValue As String
    Dim subjectValue As String
    Dim bodyValue As String

    toValue = GetTextBoxText(mtxtTO)
    subjectValue = GetTextBoxText(mTxtsubj)
    bodyValue = GetBodyText()

    If LenB(toValue) = 0 Then
        warnings.Add "To"
        SetTextBoxText mtxtTO, "<enter recipients>"
    End If

    If LenB(subjectValue) = 0 Then
        warnings.Add "Subject"
        SetTextBoxText mTxtsubj, "<enter subject>"
    End If

    If LenB(bodyValue) = 0 Then
        warnings.Add "Body"
        SetBodyText "(No body content provided)"
    End If

    If warnings.Count = 0 Then Exit Sub

    normalizedKey = UCase$(Trim$(templateKey))
    warningKey = normalizedKey & "|" & JoinCollectionString(warnings, ";")

    If Not mTemplateFieldWarningsShown Is Nothing Then
        If mTemplateFieldWarningsShown.Exists(warningKey) Then Exit Sub
        mTemplateFieldWarningsShown.Add warningKey, True
    End If

    warningText = "Template '" & templateKey & "' is missing: " & _
                  JoinCollectionString(warnings, ", ") & ". Update the Email Templates worksheet or complete the highlighted fields before sending."

    'Ref: Template prompt cleanup - emit debug details instead of modal prompts.
    Debug.Print "[EmailForm] " & warningText
    FocusComposerField
End Sub

Private Sub ShowTemplateLoadFailure(ByVal templateKey As String)
    If LenB(templateKey) = 0 Then Exit Sub

    'Ref: Template load prompts removed per template cleanup requirements.
    Debug.Print "EmailForm.ShowTemplateLoadFailure: Template column '" & templateKey & "' not found."
    FocusTemplateSelector
End Sub

Private Sub UserForm_Initialize()
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    On Error GoTo CleanFail

    mIsLoading = True
    mInitialized = False
    Debug.Print "[EmailForm] Initialize start"

    SetCursorWait

    Dim templateKey As String
    Dim templateKeys As Collection
    Dim tpl As EmailTemplate
    Dim templateEntries As Collection
    Dim entryVariant As Variant
    Dim entryValue As String
    Dim resolvedTemplateKey As String
    Dim requestedTemplateKey As String
    Dim displayName As String
    Dim fullPath As String
    Dim listIndex As Long
    Dim attachmentCount As Long
    Dim combinedAttachments As Collection
    Dim highlightLabel As MSForms.label
    Dim labelIndex As Long

    Debug.Print "EmailTemplate structure successfully recognized."

    InitializeControlReferences
    EnsureTemplateWarningCache
    mTemplateAvailabilityWarningShown = False

    mTitleBarHidden = False

    If Not mbSUB Is Nothing Then
        mbSUB.Visible = False
    End If

    If Not EnsureRequiredControls() Then
        errNumber = vbObjectError + 701
        errSource = "EmailForm.Initialize"
        errDescription = "One or more required controls are missing from the Email form."
        GoTo CleanExit
    End If

    Set templateKeys = PopulateTemplateDropdown()

    templateKey = ResolveInitialTemplateKey(templateKeys)
    requestedTemplateKey = templateKey

    If TemplateKeyExists(templateKey, templateKeys) Then
        tpl = modEmailTemplates.ReadTemplateByName(templateKey)
        resolvedTemplateKey = Trim$(templateKey)
    Else
        tpl = modEmailTemplates.ReadDefaultEmailTemplate()
        resolvedTemplateKey = Trim$(tpl.templateName)
        If LenB(requestedTemplateKey) > 0 Then
            Debug.Print "UserForm_Initialize: Template '" & requestedTemplateKey & _
                        "' not found. Using default template '" & resolvedTemplateKey & "'."
        End If
    End If

    If LenB(resolvedTemplateKey) = 0 Then
        resolvedTemplateKey = Trim$(tpl.templateName)
    End If

    modEmailTemplates.DebugPrintTemplate "EmailForm.Initialize [" & resolvedTemplateKey & "]", tpl

    SetTextBoxText mTxtcc, tpl.Cc
    SetTextBoxText mTxtsubj, tpl.Subject
    SetBodyText tpl.Body

    mOriginalSubjectTemplate = tpl.Subject
    mOriginalBodyTemplate = tpl.Body
    mCurrentTemplateKey = resolvedTemplateKey
    SetTextBoxText mtxtTEMP, resolvedTemplateKey

    'Ref: Template field cleanup - load serialized attachments directly from template entries.
    Set templateEntries = modEmailTemplates.GetTemplateAttachmentEntriesForKey(resolvedTemplateKey)

    If Not mLstAT Is Nothing Then
        On Error Resume Next
        mLstAT.Clear
        mLstAT.ColumnCount = 2
        mLstAT.ColumnWidths = CStr(mLstAT.Width) & " pt;0 pt"
        On Error GoTo 0
    End If

    Set mTemplateAttachmentEntries = New Collection

    If Not templateEntries Is Nothing Then
        For Each entryVariant In templateEntries
            entryValue = CStr(entryVariant)
            displayName = Trim$(modEmailTemplates.GetAttachmentEntryName(entryValue))
            fullPath = Trim$(modEmailTemplates.GetAttachmentEntryPath(entryValue))

            If LenB(displayName) = 0 Then displayName = fullPath

            If LenB(fullPath) = 0 Then
                Debug.Print "UserForm_Initialize: Attachment path missing for '" & displayName & "'."
            Else
                If LenB(entryValue) > 0 Then
                    mTemplateAttachmentEntries.Add entryValue
                End If
            End If

            If Not mLstAT Is Nothing Then
                On Error Resume Next
                mLstAT.AddItem displayName
                If Err.Number = 0 Then
                    listIndex = mLstAT.ListCount - 1
                    If listIndex >= 0 Then
                        mLstAT.List(listIndex, 1) = fullPath
                    End If
                Else
                    Debug.Print "UserForm_Initialize: Failed to add attachment '" & displayName & "' to lstAT."
                    Err.Clear
                End If
                On Error GoTo 0
            End If
        Next entryVariant
    End If

    If mTemplateAttachmentEntries Is Nothing Then
        Set mTemplateAttachmentEntries = New Collection
    End If

    Set mUserAttachmentEntries = GetUserAttachmentEntries(resolvedTemplateKey)
    If mUserAttachmentEntries Is Nothing Then
        Set mUserAttachmentEntries = New Collection
    End If

    Set mTemplateAttachmentLookup = CreateCaseInsensitiveDictionary()
    PopulateLookupFromEntries mTemplateAttachmentEntries, mTemplateAttachmentLookup

    If mUserAttachmentLookup Is Nothing Then
        Set mUserAttachmentLookup = CreateCaseInsensitiveDictionary()
    Else
        On Error Resume Next
        mUserAttachmentLookup.RemoveAll
        On Error GoTo 0
    End If
    RebuildLookupFromCollection mUserAttachmentLookup, mUserAttachmentEntries

    Set combinedAttachments = modEmail.BuildAttachmentDisplayList(mTemplateAttachmentEntries, _
                                                                 mUserAttachmentEntries)
    modEmail.UpdateAttachmentRemoveButton mbSUB, combinedAttachments

    attachmentCount = 0
    If Not combinedAttachments Is Nothing Then
        On Error Resume Next
        attachmentCount = combinedAttachments.Count
        On Error GoTo 0
    End If

    TraceTemplateSelection resolvedTemplateKey, True, vbNullString, _
                           GetTextBoxText(mTxtcc, False), _
                           GetTextBoxText(mTxtsubj, False), _
                           GetBodyText(), attachmentCount
    ValidateLoadedTemplateFields resolvedTemplateKey
    TraceEmailFieldState "InitializeTemplate", resolvedTemplateKey

    LoadMemberRecords

    mFirstVisibleMemberIndex = 1

    If mMemberCount > 0 Then
        SelectedMemberIndex = 1
    Else
        mSelectedMemberIndex = 0
        RenderMemberPage
    End If

    Debug.Print "[EmailForm] Initialize: set highlightLabel border"

    Set highlightLabel = GetLabelByDisplayIndex(1)
    If Not highlightLabel Is Nothing Then
        highlightLabel.BorderStyle = fmBorderStyleSingle
        highlightLabel.BorderColor = vbRed
    End If

    For labelIndex = 2 To 8
        Set highlightLabel = GetLabelByDisplayIndex(labelIndex)
        If highlightLabel Is Nothing Then GoTo NextLabel
        highlightLabel.BorderStyle = fmBorderStyleNone
NextLabel:
    Next labelIndex

    Debug.Print "[EmailForm] Initialize: highlightLabel border set"

    UpdateToFieldFromHighlightedRecord

    PopulateFromIndex 1

    ' The first member index represents worksheet row 2 because row 1 stores headers.
    ' We highlight that row silently so reviewers land on the initial record without prompts.

    CenterUserFormOnActiveMonitor Me

    mInitialized = True
    Debug.Print "[EmailForm] Initialize done"

CleanExit:
    mIsLoading = False
    SetCursorDefault
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
    Exit Sub

CleanFail:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    mInitialized = False
    Debug.Print "[EmailForm] Initialize error: " & errDescription
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
    modUIHelpers.HideUserFormTitleBar Me, mTitleBarHidden, "email"
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ClearActiveHoverLabel
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
        mOriginalBodyTemplate = GetBodyText()
    End If
    If resetTemplate Or LenB(mOriginalSubjectTemplate) = 0 Then
        mOriginalSubjectTemplate = GetTextBoxText(mTxtsubj, False)
    End If
    ApplyBodyPlaceholders memberIndex
End Sub

Public Sub LoadBodyTemplate(ByVal templateText As String, Optional ByVal memberIndex As Long = -1)
    mOriginalBodyTemplate = templateText
    SetBodyText templateText
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
    Dim nameLabel As MSForms.label
    Dim ssnLabel As MSForms.label
    Dim statusLabel As MSForms.label

    ClearActiveHoverLabel

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

Private Function GetLabelControl(ByVal prefix As String, ByVal index As Long) As MSForms.label
    Dim ctrl As MSForms.control
    Dim controlName As String

    controlName = prefix & CStr(index)

    On Error Resume Next
    Set ctrl = Me.controls(controlName)
    On Error GoTo 0

    If ctrl Is Nothing Then Exit Function
    If Not TypeOf ctrl Is MSForms.label Then Exit Function

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
    Dim statusLabel As MSForms.label

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
        baseText = GetBodyText()
    End If

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

    If LenB(baseText) > 0 Then
        SetBodyText modEmailPlaceholders.ReplacePlaceholdersArray(baseText, placeholderPairs)
    End If

    ApplySubjectPlaceholders placeholderPairs

    TraceEmailFieldState "ApplyBodyPlaceholders", ResolveActiveTemplateKey(False)
End Sub

Private Sub ApplySubjectPlaceholders(ByRef placeholderPairs As Variant)
    Dim subjectTemplate As String

    If mTxtsubj Is Nothing Then Exit Sub

    subjectTemplate = mOriginalSubjectTemplate
    If LenB(subjectTemplate) = 0 Then
        subjectTemplate = GetTextBoxText(mTxtsubj, False)
    End If

    If LenB(subjectTemplate) = 0 Then Exit Sub

    SetTextBoxText mTxtsubj, modEmailPlaceholders.ReplacePlaceholdersArray(subjectTemplate, placeholderPairs)
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
        AddPlaceholderValue placeholders, "ISSUES_SUMMARY", BuildIssuesSummary(issues, False)
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

    Dim ctrl As MSForms.control
    For Each ctrl In Me.controls
        If TypeOf ctrl Is MSForms.label Then
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
    Dim ctrl As MSForms.control
    Dim idx As Long
    Dim caption As String

    Set dict = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    dict.CompareMode = vbTextCompare
    On Error GoTo 0

    For Each ctrl In Me.controls
        If TypeOf ctrl Is MSForms.label Then
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

Private Sub HandleLabelMouseMove(ByVal target As MSForms.label)
    If target Is Nothing Then Exit Sub

    UpdateHoverLabel target
End Sub

Private Sub UpdateHoverLabel(ByVal target As MSForms.label)
    If target Is Nothing Then Exit Sub

    If Not mActiveHoverLabel Is Nothing Then
        If Not target Is mActiveHoverLabel Then
            ResetHoverLabel mActiveHoverLabel
        End If
    End If

    Set mActiveHoverLabel = target

    If target.BorderStyle <> fmBorderStyleSingle Then
        target.BorderStyle = fmBorderStyleSingle
    End If

    If target.BorderColor <> vbRed Then
        target.BorderColor = vbWhite
    End If
End Sub

Private Sub ResetHoverLabel(ByVal target As MSForms.label)
    If target Is Nothing Then Exit Sub

    If target.BorderColor <> vbRed Then
        target.BorderStyle = fmBorderStyleNone
    End If
End Sub

Private Sub ClearActiveHoverLabel()
    If mActiveHoverLabel Is Nothing Then Exit Sub

    ResetHoverLabel mActiveHoverLabel
    Set mActiveHoverLabel = Nothing
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

Private Sub ApplyStatusColor(ByVal statusLabel As MSForms.label)
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

Private Function GetListBoxEntries(ByVal listControl As MSForms.listBox) As Collection
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

Private Function GetAttachmentListCount() As Long
    If mLstAT Is Nothing Then Exit Function

    On Error GoTo CleanFail
    GetAttachmentListCount = mLstAT.ListCount
    On Error GoTo 0
    Exit Function

CleanFail:
    Err.Clear
    On Error GoTo 0
End Function

Private Sub PopulateLookupFromEntries(ByVal entries As Collection, ByVal lookup As Object)
    Dim entry As Variant
    Dim normalizedKey As String

    If lookup Is Nothing Then Exit Sub
    If entries Is Nothing Then Exit Sub

    For Each entry In entries
        normalizedKey = NormalizeTemplateAttachmentEntry(CStr(entry))
        If LenB(normalizedKey) > 0 Then
            If Not lookup.Exists(normalizedKey) Then
                lookup(normalizedKey) = True
            End If
        End If
    Next entry
End Sub

Private Sub LoadListBoxFromCollection(ByVal listControl As MSForms.listBox, _
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

    PopulateLookupFromEntries mTemplateAttachmentEntries, mTemplateAttachmentLookup
    PopulateLookupFromEntries mUserAttachmentEntries, mUserAttachmentLookup

    RefreshAttachmentListDisplay
End Sub

Private Function GetAttachmentEntriesFromListBox(ByRef listBox As MSForms.listBox) As Collection
    Dim entries As Collection
    Dim idx As Long
    Dim entryText As String

    Set entries = New Collection
    If listBox Is Nothing Then GoTo CleanExit
    If listBox.ListCount = 0 Then GoTo CleanExit

    On Error GoTo EntryError

    For idx = 0 To listBox.ListCount - 1
        entryText = Trim$(CStr(listBox.List(idx)))
        If LenB(entryText) > 0 Then
            entries.Add entryText
        End If
    Next idx

    On Error GoTo 0

CleanExit:
    Set GetAttachmentEntriesFromListBox = entries
    Exit Function

EntryError:
    entryText = vbNullString
    Err.Clear
    Resume Next
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

Private Sub RefreshTemplateAttachmentTrackingFromWorksheet(ByVal templateKey As String)
    Dim entries As Collection
    Dim lookup As Object

    templateKey = Trim$(templateKey)
    If LenB(templateKey) = 0 Then Exit Sub

    Set entries = GetTemplateAttachmentEntriesForKey(templateKey)
    If entries Is Nothing Then
        Set entries = New Collection
    End If

    Set mTemplateAttachmentEntries = entries

    Set lookup = CreateCaseInsensitiveDictionary()
    If lookup Is Nothing Then
        Set lookup = mTemplateAttachmentLookup
        If Not lookup Is Nothing Then
            On Error Resume Next
            lookup.RemoveAll
            On Error GoTo 0
        End If
    End If

    Set mTemplateAttachmentLookup = lookup
    RebuildLookupFromCollection mTemplateAttachmentLookup, mTemplateAttachmentEntries

    RefreshAttachmentListDisplay
End Sub

Private Sub RebuildLookupFromCollection(ByRef dict As Object, ByVal entries As Collection)
    PopulateLookupFromEntries entries, dict
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

Private Function AddUserAttachmentFromPath(ByVal filePath As String, _
                                           Optional ByRef failureReason As String = vbNullString) As Boolean
    Dim normalizedKey As String
    Dim entry As String
    Dim resolvedPath As String
    Dim displayName As String
    Dim originalNormalized As String
    Dim originalSelection As String
    Dim missingLabel As String

    resolvedPath = Trim$(filePath)
    displayName = vbNullString
    originalSelection = resolvedPath

    originalNormalized = NormalizeTemplateAttachmentPath(resolvedPath)

    If LenB(resolvedPath) = 0 Then
        failureReason = "No file was selected."
        Exit Function
    End If

    If Not CheckIfAttachmentExists(displayName, resolvedPath) Then
        If LenB(failureReason) = 0 Then
            missingLabel = Trim$(originalSelection)
            If LenB(missingLabel) = 0 Then missingLabel = Trim$(filePath)
            If LenB(missingLabel) = 0 Then missingLabel = "(unknown file)"
            failureReason = "The file '" & missingLabel & "' could not be found."
        End If
        Exit Function
    End If

    If LenB(displayName) = 0 Then
        displayName = ResolveDisplayNameFromPath(resolvedPath)
    End If
    If LenB(displayName) = 0 Then displayName = resolvedPath

    normalizedKey = NormalizeTemplateAttachmentPath(resolvedPath)
    If LenB(normalizedKey) = 0 Then Exit Function

    If LenB(mCurrentTemplateKey) > 0 Then
        If StrComp(originalNormalized, normalizedKey, vbTextCompare) <> 0 Then
            RefreshTemplateAttachmentTrackingFromWorksheet mCurrentTemplateKey
        End If
    End If

    If AttachmentExistsInTemplate(normalizedKey) Then
        failureReason = "The file '" & displayName & "' is already included with the template."
        Exit Function
    End If
    If AttachmentExistsInUser(normalizedKey) Then
        failureReason = "The file '" & displayName & "' has already been added."
        Exit Function
    End If

    entry = BuildAttachmentEntryFromComponents(displayName, resolvedPath)
    If LenB(entry) = 0 Then
        failureReason = "Unable to add '" & displayName & "' because the entry could not be created."
        Exit Function
    End If

    If mUserAttachmentEntries Is Nothing Then
        Set mUserAttachmentEntries = New Collection
    End If
    mUserAttachmentEntries.Add entry

    If Not mUserAttachmentLookup Is Nothing Then
        mUserAttachmentLookup(normalizedKey) = True
    End If

    AddUserAttachmentFromPath = True
    failureReason = vbNullString
End Function

Private Function ResolveDisplayNameFromPath(ByVal filePath As String) As String
    Dim separatorPos As Long

    filePath = Trim$(filePath)
    If LenB(filePath) = 0 Then Exit Function

    separatorPos = InStrRev(filePath, Application.PathSeparator)
    If separatorPos > 0 Then
        ResolveDisplayNameFromPath = Mid$(filePath, separatorPos + 1)
        If LenB(ResolveDisplayNameFromPath) = 0 Then
            ResolveDisplayNameFromPath = filePath
        End If
    Else
        ResolveDisplayNameFromPath = filePath
    End If
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

Private Sub SyncTemplateAttachmentList(ByVal templateKey As String)
    modEmail.SyncAttachmentList mLstAT, mbSUB, _
                                mTemplateAttachmentEntries, mUserAttachmentEntries
    PersistUserAttachmentsToWorksheet templateKey
    TraceEmailFieldState "SyncTemplateAttachmentList", ResolveActiveTemplateKey(False)
End Sub

Private Sub PersistUserAttachmentsToWorksheet(ByVal templateKey As String)
    Dim activeTemplateKey As String

    activeTemplateKey = Trim$(templateKey)
    If LenB(activeTemplateKey) = 0 Then
        activeTemplateKey = Trim$(mCurrentTemplateKey)
    End If

    If LenB(activeTemplateKey) = 0 Then Exit Sub

    On Error GoTo WriteFail
    WriteUserAttachmentEntries activeTemplateKey, mUserAttachmentEntries
    On Error GoTo 0
    Exit Sub

WriteFail:
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub RefreshAttachmentListDisplay()
    modEmail.SyncAttachmentList mLstAT, mbSUB, _
                                mTemplateAttachmentEntries, mUserAttachmentEntries
End Sub

Private Sub bADD_Click()
    Dim fd As FileDialog
    Dim selectedPaths As Collection
    Dim selectedItem As Variant
    Dim templateKey As String
    Dim waitApplied As Boolean
    Dim addedCount As Long
    Dim failureReasons As Collection
    Dim failureReason As String

    On Error GoTo CleanFail

    templateKey = ResolveActiveTemplateKey()

    If LenB(templateKey) = 0 Then
        modUIHelpers.ShowWarningMessage "Select a template before adding attachments."
        FocusTemplateSelector
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
        failureReason = vbNullString
        If AddUserAttachmentFromPath(CStr(selectedItem), failureReason) Then
            addedCount = addedCount + 1
        Else
            AddFailureReason failureReasons, failureReason
        End If
    Next selectedItem

    If addedCount = 0 Then
        If Not failureReasons Is Nothing Then
            modUIHelpers.ShowWarningMessage "AUTO_SPCL couldn't add any attachments:" & vbCrLf & " - " & _
                                           JoinCollectionString(failureReasons, vbCrLf & " - ") & vbCrLf & _
                                           "Review the issues and try again."
        Else
            modUIHelpers.ShowInfoMessage "No attachments were added. The selected files may already be listed in this draft or were unavailable."
        End If
        FocusAttachmentList
        GoTo CleanExit
    End If

    SetCursorWait
    waitApplied = True

    SyncTemplateAttachmentList templateKey

    If Not failureReasons Is Nothing Then
        modUIHelpers.ShowWarningMessage "Some files were skipped:" & vbCrLf & " - " & _
                                        JoinCollectionString(failureReasons, vbCrLf & " - ") & vbCrLf & _
                                        "Those files remain unchanged."
        FocusAttachmentList
    End If

CleanExit:
    If waitApplied Then SetCursorDefault
    Set fd = Nothing
    Set selectedPaths = Nothing
    Exit Sub

CleanFail:
    If waitApplied Then SetCursorDefault
    modUIHelpers.ShowErrorMessage "AUTO_SPCL couldn't add attachments: " & Err.Description
    FocusAttachmentList
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

    templateKey = ResolveActiveTemplateKey()

    If LenB(templateKey) = 0 Then
        modUIHelpers.ShowWarningMessage "Select a template before removing attachments."
        FocusTemplateSelector
        GoTo CleanExit
    End If

    EnsureAttachmentTracking templateKey

    Set userPaths = BuildUserAttachmentPaths()
    If userPaths Is Nothing Or userPaths.Count = 0 Then
        modUIHelpers.ShowInfoMessage "There are no user-added attachments to remove."
        FocusAttachmentList
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
        modUIHelpers.ShowInfoMessage "None of the selected files were added by this draft. Template attachments cannot be removed."
        FocusAttachmentList
        GoTo CleanExit
    End If

    For Each selectedItem In selectedPaths
        If RemoveUserAttachmentFromPath(CStr(selectedItem)) Then
            removedCount = removedCount + 1
        End If
    Next selectedItem

    If removedCount = 0 Then
        modUIHelpers.ShowInfoMessage "No attachments were removed. They may have already been cleared."
        FocusAttachmentList
        GoTo CleanExit
    End If

    SetCursorWait
    waitApplied = True

    SyncTemplateAttachmentList templateKey

    If ignoredCount > 0 Then
        modUIHelpers.ShowInfoMessage "Some selected files were ignored because they belong to the template and cannot be removed."
        FocusAttachmentList
    End If

CleanExit:
    If waitApplied Then SetCursorDefault
    Set fd = Nothing
    Set selectedPaths = Nothing
    Exit Sub

CleanFail:
    If waitApplied Then SetCursorDefault
    modUIHelpers.ShowErrorMessage "AUTO_SPCL couldn't remove attachments: " & Err.Description
    FocusAttachmentList
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
    modUIHelpers.ShowErrorMessage "AUTO_SPCL couldn't close the email workspace: " & errDescription
    modUIHelpers.EnsureFormFocus Me
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
        modUIHelpers.ShowInfoMessage "All members are marked as Cancel. Mark at least one member as Draft, then try again."
        modUIHelpers.EnsureFormFocus Me
        GoTo CleanExit
    End If

    templateKey = ResolveActiveTemplateKey()
    If LenB(templateKey) = 0 Then
        SetCursorDefault
        modUIHelpers.ShowWarningMessage "Select a template before creating drafts."
        FocusTemplateSelector
        GoTo CleanExit
    End If

    EnsureAttachmentTracking templateKey

    Set templateEntries = mTemplateAttachmentEntries
    Set userEntries = mUserAttachmentEntries

    CreateDraftsFromID whitelist, templateKey, templateEntries, userEntries
    modUIHelpers.EnsureFormFocus Me

CleanExit:
    SetCursorDefault
    modUIHelpers.EnsureFormFocus Me
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
    Exit Sub

CleanFail:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    modUIHelpers.ShowErrorMessage "AUTO_SPCL couldn't create Outlook drafts: " & errDescription
    modUIHelpers.EnsureFormFocus Me
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

Private Sub PopulateFromIndex(ByVal idx As Long)
    Debug.Print "[EmailForm] PopulateFromIndex idx=" & idx
    UpdateIssuePlaceholderForDisplayIndex idx
End Sub

Private Sub bBE_Click()
    HandleEmailToggleClick SelectedMemberIndex
End Sub

Private Sub lblL1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HandleLabelMouseMoveByIndex 1
End Sub

Private Sub lblL1_Click()
    If mIsLoading Then Exit Sub
    HandleLabelClickByIndex 1
    PopulateFromIndex 1
End Sub

Private Sub lblL2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HandleLabelMouseMoveByIndex 2
End Sub

Private Sub lblL2_Click()
    If mIsLoading Then Exit Sub
    HandleLabelClickByIndex 2
    PopulateFromIndex 2
End Sub

Private Sub lblL3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HandleLabelMouseMoveByIndex 3
End Sub

Private Sub lblL3_Click()
    If mIsLoading Then Exit Sub
    HandleLabelClickByIndex 3
    PopulateFromIndex 3
End Sub

Private Sub lblL4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HandleLabelMouseMoveByIndex 4
End Sub

Private Sub lblL4_Click()
    If mIsLoading Then Exit Sub
    HandleLabelClickByIndex 4
    PopulateFromIndex 4
End Sub

Private Sub lblL5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HandleLabelMouseMoveByIndex 5
End Sub

Private Sub lblL5_Click()
    If mIsLoading Then Exit Sub
    HandleLabelClickByIndex 5
    PopulateFromIndex 5
End Sub

Private Sub lblL6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HandleLabelMouseMoveByIndex 6
End Sub

Private Sub lblL6_Click()
    If mIsLoading Then Exit Sub
    HandleLabelClickByIndex 6
    PopulateFromIndex 6
End Sub

Private Sub lblL7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HandleLabelMouseMoveByIndex 7
End Sub

Private Sub lblL7_Click()
    If mIsLoading Then Exit Sub
    HandleLabelClickByIndex 7
    PopulateFromIndex 7
End Sub

Private Sub lblL8_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HandleLabelMouseMoveByIndex 8
End Sub

Private Sub lblL8_Click()
    If mIsLoading Then Exit Sub
    HandleLabelClickByIndex 8
    PopulateFromIndex 8
End Sub





