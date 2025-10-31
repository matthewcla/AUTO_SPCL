Attribute VB_Name = "modEmail"
Option Explicit

Private Const DEFAULT_CC_LIST As String = vbNullString
Private Const DEFAULT_SUBJECT_TEMPLATE As String = "{Name} - AUTO_SPCL Draft"
Private Const DEFAULT_BODY_TEMPLATE As String = _
        "Dear {Name}," & vbCrLf & vbCrLf & _
        "Eligibility note: {EligiblesNote}" & vbCrLf & vbCrLf & _
        "Respectfully," & vbCrLf & _
        "AUTO_SPCL Team"
Private Const DEFAULT_ELIG_NOTE_TEXT As String = "(no note found)"

'-------------------------------------------------------------------------------
' Procedure: ClearEmailFields
' Purpose  : Reset the outbound email form so the next draft starts from a clean state.
' Parameters:
'   txtTo - Text box that collects the primary recipients.
'   txtCc - Text box that collects the carbon copy recipients.
'   txtSubject - Text box containing the email subject.
'   txtBody - Text box containing the email body template.
'   txtSignature - Text box that holds the saved signature block.
'   lstAttachments - Optional list box showing the current attachments.
'   btnRemoveAttachment - Optional button used to remove selected attachments.
' Returns  : None.
' Side Effects:
'   Clears UI controls and hides/disables the remove-attachment button when no files remain.
'-------------------------------------------------------------------------------
Public Sub ClearEmailFields(ByRef txtTo As MSForms.TextBox, _
                            ByRef txtCc As MSForms.TextBox, _
                            ByRef txtSubject As MSForms.TextBox, _
                            ByRef txtBody As MSForms.TextBox, _
                            ByRef txtSignature As MSForms.TextBox, _
                            Optional ByRef lstAttachments As MSForms.ListBox, _
                            Optional ByRef btnRemoveAttachment As MSForms.CommandButton)

    AssignTextBoxValue txtTo, vbNullString
    AssignTextBoxValue txtCc, vbNullString
    AssignTextBoxValue txtSubject, vbNullString
    AssignTextBoxValue txtBody, vbNullString
    AssignTextBoxValue txtSignature, vbNullString

    If Not lstAttachments Is Nothing Then
        On Error Resume Next
        lstAttachments.Clear
        On Error GoTo 0
    End If

    UpdateAttachmentRemoveButton btnRemoveAttachment, Nothing
End Sub

'-------------------------------------------------------------------------------
' Procedure: BuildAttachmentDisplayList
' Purpose  : Combine template-specified and user-specified attachment entries for display.
' Parameters:
'   templateEntries - Collection of attachment descriptors defined by the template.
'   userEntries - Collection of attachment descriptors added by the user.
' Returns  : Collection containing the concatenated attachment descriptors as strings.
' Side Effects:
'   None.
'-------------------------------------------------------------------------------
Public Function BuildAttachmentDisplayList(ByVal templateEntries As Collection, _
                                           ByVal userEntries As Collection) As Collection
    Dim combined As Collection
    Dim entry As Variant

    Set combined = New Collection

    If Not templateEntries Is Nothing Then
        For Each entry In templateEntries
            combined.Add CStr(entry)
        Next entry
    End If

    If Not userEntries Is Nothing Then
        For Each entry In userEntries
            combined.Add CStr(entry)
        Next entry
    End If

    Set BuildAttachmentDisplayList = combined
End Function

'-------------------------------------------------------------------------------
' Procedure: SyncAttachmentList
' Purpose  : Refresh the attachment list UI and button state based on current sources.
' Parameters:
'   lstAttachments - List box displaying attachment summaries.
'   btnRemoveAttachment - Button that removes the selected attachment.
'   templateEntries - Template-level attachment entries to include.
'   userEntries - User-level attachment entries to include.
' Returns  : Collection representing the combined attachments currently shown.
' Side Effects:
'   Updates list box items and remove button enabled/visible flags.
'-------------------------------------------------------------------------------
Public Function SyncAttachmentList(ByRef lstAttachments As MSForms.ListBox, _
                                   ByRef btnRemoveAttachment As MSForms.CommandButton, _
                                   ByVal templateEntries As Collection, _
                                   ByVal userEntries As Collection) As Collection
    Dim combined As Collection

    Set combined = BuildAttachmentDisplayList(templateEntries, userEntries)
    PopulateAttachmentList lstAttachments, combined
    UpdateAttachmentRemoveButton btnRemoveAttachment, combined

    Set SyncAttachmentList = combined
End Function

'-------------------------------------------------------------------------------
' Procedure: UpdateAttachmentRemoveButton
' Purpose  : Toggle the remove-attachment button so it mirrors whether files are present.
' Parameters:
'   btnRemoveAttachment - Button control used to remove attachments.
'   attachments - Collection of attachment entries currently available.
' Returns  : None.
' Side Effects:
'   Shows/hides and enables/disables the button through the MSForms API.
'-------------------------------------------------------------------------------
Public Sub UpdateAttachmentRemoveButton(ByRef btnRemoveAttachment As MSForms.CommandButton, _
                                        ByVal attachments As Collection)
    Dim hasAttachments As Boolean

    hasAttachments = HasAttachmentEntries(attachments)

    If btnRemoveAttachment Is Nothing Then Exit Sub

    ' Only expose the remove button when at least one attachment is available to act on.
    On Error Resume Next
    btnRemoveAttachment.Visible = hasAttachments
    btnRemoveAttachment.Enabled = hasAttachments
    On Error GoTo 0
End Sub

Private Sub PopulateAttachmentList(ByRef lstAttachments As MSForms.ListBox, _
                                   ByVal entries As Collection)
    Dim entry As Variant

    If lstAttachments Is Nothing Then Exit Sub

    lstAttachments.Clear

    If entries Is Nothing Then Exit Sub

    For Each entry In entries
        On Error Resume Next
        lstAttachments.AddItem CStr(entry)
        On Error GoTo 0
    Next entry
End Sub

Private Function HasAttachmentEntries(ByVal attachments As Collection) As Boolean
    If attachments Is Nothing Then Exit Function

    On Error Resume Next
    HasAttachmentEntries = (attachments.Count > 0)
    If Err.Number <> 0 Then
        Err.Clear
        HasAttachmentEntries = False
    End If
    On Error GoTo 0
End Function

Private Sub AssignTextBoxValue(ByRef target As MSForms.TextBox, ByVal value As String)
    If target Is Nothing Then Exit Sub
    target.Value = value
End Sub

'-------------------------------------------------------------------------------
' Procedure: CreateDraftsFromID
' Purpose  : Generate hidden Outlook draft messages for members listed on the ID sheet.
' Parameters:
'   allowedMembers - Optional whitelist of member indexes or names to restrict processing.
'   templateKey - Optional template identifier to pre-populate content and attachments.
'   templateAttachmentEntries - Optional pre-resolved template attachment entries.
'   userAttachmentEntries - Optional pre-resolved user attachment entries.
' Returns  : None.
' Side Effects:
'   Starts Outlook if necessary, reads workbook data, creates Outlook draft items, and
'   shows modal messages summarizing success/failure counts.
'-------------------------------------------------------------------------------
Public Sub CreateDraftsFromID(Optional ByVal allowedMembers As Variant, _
                              Optional ByVal templateKey As String = vbNullString, _
                              Optional ByVal templateAttachmentEntries As Variant, _
                              Optional ByVal userAttachmentEntries As Variant)
    Dim wsID As Worksheet, wsElig As Worksheet
    Dim lastRow As Long, r As Long
    Dim personName As String, toList As String, eligNote As String
    Dim olApp As Object, olMail As Object  ' Outlook.Application / MailItem (late bound)
    Dim createdCount As Long, skippedCount As Long
    Dim whitelist As Object
    Dim hasWhitelist As Boolean
    Dim memberIndex As Long
    Dim skipNote As String
    Dim templateAttachmentPaths As Collection
    Dim userAttachmentPaths As Collection
    Dim providedTemplateEntries As Collection
    Dim providedUserEntries As Collection
    Dim storedUserEntries As Collection
    Dim attachmentPath As Variant
    Dim progressTotal As Long
    Dim progressProcessed As Long
    Dim progressActive As Boolean
    Dim progressClosed As Boolean
    Dim cancelledByUser As Boolean
    Dim progressLabel As String
    Dim progressOutcome As String
    Dim summary As String
    Dim finalNote As String
    Dim screenUpdatingSuspended As Boolean
    Dim ccList As String
    Dim subjectTemplate As String
    Dim bodyTemplate As String
    Dim mailSubject As String
    Dim mailBody As String

    On Error GoTo CleanFail

    Set wsID = ThisWorkbook.Worksheets("ID")
    Set wsElig = ThisWorkbook.Worksheets("Eligibles RED Board")

    If Not IsMissing(allowedMembers) Then
        Set whitelist = NormalizeDraftWhitelist(allowedMembers)
        hasWhitelist = Not whitelist Is Nothing
    End If

    lastRow = wsID.Cells(wsID.Rows.Count, "B").End(xlUp).row
    If lastRow < 2 Then
        modUIHelpers.ShowWarningMessage "AUTO_SPCL couldn't find any member rows on 'ID'. Add member names to column B, then try again."
        GoTo CleanExit
    End If

    progressTotal = lastRow - 1
    If progressTotal > 0 Then
        modProgressUI.Progress_Show progressTotal, "Creating Outlook Drafts"
        modProgressUI.Progress_Log "Preparing Outlook session..."
        progressActive = True
    End If

    ' Get or start Outlook
    On Error Resume Next
    Set olApp = GetObject(Class:="Outlook.Application")
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    On Error GoTo CleanFail
    If olApp Is Nothing Then
        If progressActive Then
            modProgressUI.Progress_Close "Unable to connect to Outlook."
            progressClosed = True
            progressActive = False
        End If
        modUIHelpers.ShowErrorMessage "AUTO_SPCL couldn't connect to Outlook. Start Outlook and run draft creation again."
        GoTo CleanExit
    End If

    If Not IsMissing(templateAttachmentEntries) Then
        If IsObject(templateAttachmentEntries) Then
            On Error Resume Next
            Set providedTemplateEntries = templateAttachmentEntries
            On Error GoTo CleanFail
        End If
    End If

    If Not providedTemplateEntries Is Nothing Then
        Set templateAttachmentPaths = ResolveAttachmentPathsFromEntries(providedTemplateEntries)
    End If

    If templateAttachmentPaths Is Nothing Then
        If LenB(templateKey) > 0 Then
            Set templateAttachmentPaths = GetValidatedTemplateAttachmentPaths(templateKey)
        End If
    End If

    If Not IsMissing(userAttachmentEntries) Then
        If IsObject(userAttachmentEntries) Then
            On Error Resume Next
            Set providedUserEntries = userAttachmentEntries
            On Error GoTo CleanFail
        End If
    End If

    If Not providedUserEntries Is Nothing Then
        Set userAttachmentPaths = ResolveAttachmentPathsFromEntries(providedUserEntries)
    End If

    If userAttachmentPaths Is Nothing Then
        If LenB(templateKey) > 0 Then
            Set storedUserEntries = GetUserAttachmentEntries(templateKey)
            If Not storedUserEntries Is Nothing Then
                Set userAttachmentPaths = ResolveAttachmentPathsFromEntries(storedUserEntries)
            End If
        End If
    End If

    If progressActive Then
        modProgressUI.Progress_Log "Attachment lists prepared."
    End If

    ResolveDraftTemplateContent templateKey, ccList, subjectTemplate, bodyTemplate

    Application.ScreenUpdating = False
    screenUpdatingSuspended = True

    For r = 2 To lastRow
        memberIndex = r - 1
        personName = Trim$(wsID.Cells(r, "B").Value)
        progressLabel = ResolveDraftProgressLabel(personName, memberIndex)
        progressOutcome = "Skipped"

        If progressActive Then
            If Not modProgressUI.Progress_WaitIfPaused() Then
                cancelledByUser = True
                Exit For
            End If
            If modProgressUI.Progress_Cancelled() Then
                cancelledByUser = True
                Exit For
            End If
        End If

        If hasWhitelist Then
            If Not DraftWhitelistAllowsMember(memberIndex, personName, whitelist) Then
                skippedCount = skippedCount + 1
                progressOutcome = "Skipped (not marked Draft)"
                GoTo nextRow
            End If
        End If

        If Len(personName) = 0 Then
            skippedCount = skippedCount + 1
            progressOutcome = "Skipped (missing name)"
            GoTo nextRow
        End If

        ' Build To: from columns C:F (semicolon-separated)
        toList = BuildEmailList(wsID, r, "C", "F")
        If Len(toList) = 0 Then
            ' No valid email addresses found for this row
            skippedCount = skippedCount + 1
            progressOutcome = "Skipped (no email addresses)"
            GoTo nextRow
        End If

        ' Lookup note from Eligibles col A -> take col C
        eligNote = GetEligiblesNote(wsElig, personName)

        mailSubject = BuildSubject(personName, eligNote, subjectTemplate)
        mailBody = BuildBody(personName, eligNote, bodyTemplate)

        ' Create the draft (hidden; saved to Drafts)
        Set olMail = olApp.CreateItem(0) ' olMailItem = 0
        With olMail
            .To = toList
            .CC = ccList
            .Subject = mailSubject
            .Body = mailBody
            If Not templateAttachmentPaths Is Nothing Then
                For Each attachmentPath In templateAttachmentPaths
                    If LenB(Trim$(CStr(attachmentPath))) > 0 Then
                        On Error Resume Next
                        .Attachments.Add CStr(attachmentPath)
                        If Err.Number <> 0 Then Err.Clear
                        On Error GoTo CleanFail
                    End If
                Next attachmentPath
            End If
            If Not userAttachmentPaths Is Nothing Then
                For Each attachmentPath In userAttachmentPaths
                    If LenB(Trim$(CStr(attachmentPath))) > 0 Then
                        On Error Resume Next
                        .Attachments.Add CStr(attachmentPath)
                        If Err.Number <> 0 Then Err.Clear
                        On Error GoTo CleanFail
                    End If
                Next attachmentPath
            End If
            .Save            ' <-- creates draft in Outlook Drafts
            ' .Display       ' (intentionally NOT displayed to keep it hidden)
        End With
        createdCount = createdCount + 1
        progressOutcome = "Draft created"
nextRow:
        If progressActive Then
            progressProcessed = progressProcessed + 1
            modProgressUI.Progress_Update progressProcessed, progressTotal, _
                                          FormatDraftProgressStatus(progressLabel, progressOutcome)
            modProgressUI.Progress_Log FormatDraftProgressStatus(progressLabel, progressOutcome)
        End If
    Next r

    Application.ScreenUpdating = True
    screenUpdatingSuspended = False

    If hasWhitelist Then skipNote = " (includes members not marked as Draft)"
    summary = BuildDraftSummary(createdCount, skippedCount, skipNote)

    If progressActive And Not progressClosed Then
        If cancelledByUser Then
            finalNote = "Draft creation cancelled." & vbCrLf & summary
        Else
            finalNote = "Draft creation complete." & vbCrLf & summary
        End If
        modProgressUI.Progress_Close finalNote
        progressClosed = True
    End If

    If cancelledByUser Then
        modUIHelpers.ShowWarningMessage "Draft creation was cancelled." & vbCrLf & summary
    Else
        modUIHelpers.ShowInfoMessage "Draft creation complete." & vbCrLf & summary
    End If
    GoTo CleanExit

CleanFail:
    If screenUpdatingSuspended Then
        Application.ScreenUpdating = True
        screenUpdatingSuspended = False
    End If
    If progressActive And Not progressClosed Then
        modProgressUI.Progress_Close "Draft creation failed."
        progressClosed = True
    End If
    modUIHelpers.ShowErrorMessage "Draft creation failed (" & Err.Number & "): " & Err.Description

CleanExit:
    If screenUpdatingSuspended Then
        Application.ScreenUpdating = True
        screenUpdatingSuspended = False
    End If
    If progressActive And Not progressClosed Then
        modProgressUI.Progress_Close
        progressClosed = True
    End If
End Sub

Private Function ResolveDraftProgressLabel(ByVal personName As String, _
                                           ByVal memberIndex As Long) As String
    Dim trimmedName As String

    trimmedName = Trim$(personName)
    If LenB(trimmedName) > 0 Then
        ResolveDraftProgressLabel = trimmedName
    ElseIf memberIndex > 0 Then
        ResolveDraftProgressLabel = "Row " & CStr(memberIndex)
    Else
        ResolveDraftProgressLabel = "Row ?"
    End If
End Function

Private Function FormatDraftProgressStatus(ByVal label As String, _
                                           ByVal outcome As String) As String
    label = Trim$(label)
    If LenB(label) = 0 Then label = "Current member"

    outcome = Trim$(outcome)
    If LenB(outcome) > 0 Then
        FormatDraftProgressStatus = label & " - " & outcome
    Else
        FormatDraftProgressStatus = label
    End If
End Function

Private Function BuildDraftSummary(ByVal createdCount As Long, _
                                   ByVal skippedCount As Long, _
                                   ByVal skipNote As String) As String
    BuildDraftSummary = "Created: " & createdCount & vbCrLf & _
                        "Skipped (no name, no emails, or filtered out): " & skippedCount & skipNote
End Function

' Build a semicolon-separated list of valid emails from columns startCol to endCol on a given row.
Private Function BuildEmailList(ws As Worksheet, ByVal rowNum As Long, ByVal startCol As String, ByVal endCol As String) As String
    Dim c As Long, firstCol As Long, lastCol As Long
    Dim v As String, outList As String
    
    firstCol = ws.Columns(startCol).Column
    lastCol = ws.Columns(endCol).Column
    
    For c = firstCol To lastCol
        v = Trim$(ws.Cells(rowNum, c).Value)
        If IsLikelyEmail(v) Then
            If Len(outList) > 0 Then outList = outList & "; "
            outList = outList & v
        End If
    Next c
    
    BuildEmailList = outList
End Function

' Very light email heuristic (avoid hard failures on odd data).
Private Function IsLikelyEmail(ByVal s As String) As Boolean
    If Len(s) = 0 Then Exit Function
    ' Basic checks; tweak if you need stricter validation
    IsLikelyEmail = (InStr(1, s, "@") > 0 And InStrRev(s, ".") > InStr(1, s, "@"))
End Function

' Find personName in column A of Eligibles and return the value from column C (same row),
' or an empty string if not found.
Private Function GetEligiblesNote(wsElig As Worksheet, ByVal personName As String) As String
    Dim lastRow As Long, r As Long
    lastRow = wsElig.Cells(wsElig.Rows.Count, "A").End(xlUp).row
    For r = 2 To lastRow
        If StrComp(Trim$(wsElig.Cells(r, "A").Value), personName, vbTextCompare) = 0 Then
            GetEligiblesNote = Trim$(wsElig.Cells(r, "C").Value)
            Exit Function
        End If
    Next r
    ' Not found; return empty (or a default note if you prefer)
    GetEligiblesNote = ""
End Function

' Build the email body by replacing placeholders in the supplied template text.
Private Function BuildBody(ByVal personName As String, _
                           ByVal eligNote As String, _
                           ByVal bodyTemplate As String) As String
    Dim replacements As Variant

    If LenB(bodyTemplate) = 0 Then bodyTemplate = DEFAULT_BODY_TEMPLATE

    replacements = BuildDraftPlaceholderReplacements(personName, eligNote)
    BuildBody = modEmailPlaceholders.ReplacePlaceholdersArray(bodyTemplate, replacements)
End Function

Private Function BuildSubject(ByVal personName As String, _
                              ByVal eligNote As String, _
                              ByVal subjectTemplate As String) As String
    Dim replacements As Variant

    If LenB(subjectTemplate) = 0 Then subjectTemplate = DEFAULT_SUBJECT_TEMPLATE

    replacements = BuildDraftPlaceholderReplacements(personName, eligNote)
    BuildSubject = modEmailPlaceholders.ReplacePlaceholdersArray(subjectTemplate, replacements)
End Function

Private Function BuildDraftPlaceholderReplacements(ByVal personName As String, _
                                                   ByVal eligNote As String) As Variant
    Dim noteText As String

    noteText = ResolveEligibilityNote(eligNote)

    BuildDraftPlaceholderReplacements = Array( _
        "Name", personName, _
        "EligiblesNote", noteText, _
        "ISSUES", noteText, _
        "ISSUE", noteText _
    )
End Function

Private Function ResolveEligibilityNote(ByVal eligNote As String) As String
    Dim trimmedNote As String

    trimmedNote = Trim$(eligNote)
    If LenB(trimmedNote) > 0 Then
        ResolveEligibilityNote = trimmedNote
    Else
        ResolveEligibilityNote = DEFAULT_ELIG_NOTE_TEXT
    End If
End Function

Private Sub ResolveDraftTemplateContent(ByVal templateKey As String, _
                                        ByRef ccList As String, _
                                        ByRef subjectTemplate As String, _
                                        ByRef bodyTemplate As String)
    Dim templateCc As String
    Dim templateSubject As String
    Dim templateGreeting As String
    Dim templateBody As String
    Dim templateSignature As String
    Dim combinedBody As String

    ccList = DEFAULT_CC_LIST
    subjectTemplate = DEFAULT_SUBJECT_TEMPLATE
    bodyTemplate = DEFAULT_BODY_TEMPLATE

    If LenB(templateKey) = 0 Then Exit Sub

    If modEmailTemplates.TryGetTemplateDraftContent(templateKey, templateCc, templateSubject, _
                                                   templateGreeting, templateBody, templateSignature) Then
        combinedBody = CombineDraftBodyTemplate(templateGreeting, templateBody, templateSignature)

        If LenB(templateCc) > 0 Then ccList = templateCc
        If LenB(templateSubject) > 0 Then subjectTemplate = templateSubject
        If LenB(combinedBody) > 0 Then
            bodyTemplate = combinedBody
        ElseIf LenB(templateBody) > 0 Then
            bodyTemplate = templateBody
        End If
    End If
End Sub

Private Function CombineDraftBodyTemplate(ByVal greetingValue As String, _
                                          ByVal bodyValue As String, _
                                          ByVal signatureValue As String) As String
    Dim builder As String

    greetingValue = Trim$(greetingValue)
    bodyValue = Trim$(bodyValue)
    signatureValue = Trim$(signatureValue)

    If LenB(greetingValue) > 0 Then
        builder = greetingValue
    End If

    If LenB(bodyValue) > 0 Then
        If LenB(builder) > 0 Then builder = builder & vbCrLf & vbCrLf
        builder = builder & bodyValue
    End If

    If LenB(signatureValue) > 0 Then
        If LenB(builder) > 0 Then builder = builder & vbCrLf & vbCrLf
        builder = builder & signatureValue
    End If

    CombineDraftBodyTemplate = builder
End Function

Private Function NormalizeDraftWhitelist(ByVal allowedMembers As Variant) As Object
    Dim dict As Object
    Dim key As Variant
    Dim normalizedKey As String

    If IsObject(allowedMembers) Then
        If allowedMembers Is Nothing Then Exit Function
    ElseIf IsArray(allowedMembers) Then
        ' continue
    Else
        If VarType(allowedMembers) = vbEmpty Then Exit Function
    End If

    Set dict = CreateObject("Scripting.Dictionary")
    If dict Is Nothing Then Exit Function
    On Error Resume Next
    dict.CompareMode = vbTextCompare
    On Error GoTo 0

    If IsObject(allowedMembers) Then
        For Each key In allowedMembers
            normalizedKey = NormalizeDraftWhitelistKey(CStr(key))
            If LenB(normalizedKey) > 0 Then dict(normalizedKey) = True
        Next key
    ElseIf IsArray(allowedMembers) Then
        For Each key In allowedMembers
            normalizedKey = DraftWhitelistIndexKey(key)
            If LenB(normalizedKey) > 0 Then dict(normalizedKey) = True
        Next key
    Else
        normalizedKey = DraftWhitelistIndexKey(allowedMembers)
        If LenB(normalizedKey) > 0 Then dict(normalizedKey) = True
    End If

    If dict.Count = 0 Then Exit Function
    Set NormalizeDraftWhitelist = dict
End Function

Private Function NormalizeDraftWhitelistKey(ByVal rawKey As String) As String
    Dim trimmedKey As String

    trimmedKey = UCase$(Trim$(rawKey))
    If LenB(trimmedKey) = 0 Then Exit Function

    If Left$(trimmedKey, 4) = "IDX:" Or Left$(trimmedKey, 5) = "NAME:" Then
        NormalizeDraftWhitelistKey = trimmedKey
    ElseIf IsNumeric(trimmedKey) Then
        NormalizeDraftWhitelistKey = DraftWhitelistIndexKey(trimmedKey)
    End If
End Function

Private Function DraftWhitelistAllowsMember(ByVal memberIndex As Long, _
                                            ByVal personName As String, _
                                            ByVal whitelist As Object) As Boolean
    Dim indexKey As String
    Dim nameKey As String

    If whitelist Is Nothing Then
        DraftWhitelistAllowsMember = True
        Exit Function
    End If

    indexKey = DraftWhitelistIndexKey(memberIndex)
    If LenB(indexKey) > 0 Then
        If whitelist.Exists(indexKey) Then
            DraftWhitelistAllowsMember = True
            Exit Function
        End If
    End If

    nameKey = DraftWhitelistNameKey(personName)
    If LenB(nameKey) > 0 Then
        DraftWhitelistAllowsMember = whitelist.Exists(nameKey)
    End If
End Function

Private Function DraftWhitelistIndexKey(ByVal candidate As Variant) As String
    Dim idx As Long

    If IsNumeric(candidate) Then
        idx = CLng(candidate)
        If idx > 0 Then DraftWhitelistIndexKey = "IDX:" & CStr(idx)
    End If
End Function

Private Function DraftWhitelistNameKey(ByVal candidate As String) As String
    Dim normalized As String

    normalized = DraftWhitelistNormalizeName(candidate)
    If LenB(normalized) > 0 Then DraftWhitelistNameKey = "NAME:" & normalized
End Function

Private Function DraftWhitelistNormalizeName(ByVal value As String) As String
    DraftWhitelistNormalizeName = UCase$(Trim$(value))
End Function




