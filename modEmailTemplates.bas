Attribute VB_Name = "modEmailTemplates"
Option Explicit

Private Const TEMPLATE_SHEET_NAME_PRIMARY As String = "EmailTemplate"
Private Const TEMPLATE_SHEET_NAME_ALT As String = "EmailTemplates"
Private Const TEMPLATE_SHEET_NAME_ALT2 As String = "Email Templates"
Private Const TEMPLATE_COLUMN_INDEX As Long = 2
Private Const DEFAULT_TEMPLATE_KEY As String = "Default"

Private Const EMAIL_ROW_TO As Long = 2
Private Const EMAIL_ROW_CC As Long = 3
Private Const EMAIL_ROW_SUBJECT As Long = 4
Private Const EMAIL_ROW_BODY As Long = 5
Private Const EMAIL_ROW_USER_ATTACHMENT_NAMES As Long = 11
Private Const EMAIL_ROW_USER_ATTACHMENT_PATHS As Long = 12
Private Const EMAIL_ROW_GREETING As Long = 6
Private Const EMAIL_ROW_SIGNATURE As Long = 7
Private Const EMAIL_ROW_ATTACHMENTS As Long = 9
Private Const ENABLE_TEMPLATE_TRACE As Boolean = False

Private mTemplateWorksheet As Worksheet
Private mAttachmentExistsCache As Object

'-------------------------------------------------------------------------------
' Procedure: LoadEmailTemplateIntoControls
' Purpose  : Populate the email composition controls with content pulled from the
'            template worksheet column matching the provided key.
' Parameters:
'   templateKey - Column header identifying which template to load.
'   txtTO - Text box receiving the To recipients.
'   txtCC - Text box receiving the CC recipients.
'   lstAT - List box that surfaces template and user attachment summaries.
'   txtSubj - Text box that receives the subject line.
'   txtBody - Text box that receives the greeting/body combination.
'   txtSignature - (Optional) Text box that receives the template signature block.
' Returns  : True when the template column is found and controls are populated; False otherwise.
' Side Effects:
'   Clears and updates supplied controls; combines template and stored user attachments.
'-------------------------------------------------------------------------------
Public Function LoadEmailTemplateIntoControls(ByVal templateKey As String, _
                                      ByRef txtTO As MSForms.TextBox, _
                                      ByRef txtCC As MSForms.TextBox, _
                                      ByRef lstAT As MSForms.ListBox, _
                                      ByRef txtSubj As MSForms.TextBox, _
                                      ByRef txtBody As MSForms.TextBox, _
                                      ByRef txtSignature As MSForms.TextBox) As Boolean
    Dim ws As Worksheet
    Dim templateColumn As Long
    Dim toValue As String
    Dim ccValue As String
    Dim subjValue As String
    Dim bodyValue As String
    Dim greetingValue As String
    Dim signatureValue As String
    Dim attachmentValue As String
    Dim attachmentEntries As Collection
    Dim userAttachmentEntries As Collection
    Dim combinedAttachments As Collection

    Set ws = ResolveTemplateWorksheet()
    If ws Is Nothing Then Exit Function

    templateKey = NormaliseTemplateKey(templateKey, ws)

    ClearTemplateControls txtTO, txtCC, lstAT, txtSubj, txtBody, txtSignature

    templateColumn = ResolveTemplateColumnIndex(ws, templateKey)
    If templateColumn = 0 Then Exit Function

    toValue = Trim$(CStrSafe(ws.Cells(EMAIL_ROW_TO, templateColumn).Value))
    ccValue = Trim$(CStrSafe(ws.Cells(EMAIL_ROW_CC, templateColumn).Value))
    subjValue = Trim$(CStrSafe(ws.Cells(EMAIL_ROW_SUBJECT, templateColumn).Value))
    bodyValue = Trim$(CStrSafe(ws.Cells(EMAIL_ROW_BODY, templateColumn).Value))
    greetingValue = Trim$(CStrSafe(ws.Cells(EMAIL_ROW_GREETING, templateColumn).Value))
    signatureValue = Trim$(CStrSafe(ws.Cells(EMAIL_ROW_SIGNATURE, templateColumn).Value))
    attachmentValue = Trim$(CStrSafe(ws.Cells(EMAIL_ROW_ATTACHMENTS, templateColumn).Value))
    Set attachmentEntries = ValidateTemplateAttachmentPaths(ws, templateColumn, attachmentValue)
    Set userAttachmentEntries = ReadUserAttachmentEntriesFromWorksheet(ws, templateColumn)
    Set combinedAttachments = CombineAttachmentCollections(attachmentEntries, userAttachmentEntries)

    AssignTemplateTextBoxValue txtTO, toValue
    AssignTemplateTextBoxValue txtCC, ccValue
    AssignTemplateTextBoxValue txtSubj, subjValue
    AssignAttachmentList lstAT, combinedAttachments
    AssignTemplateTextBoxValue txtBody, BuildBodyValue(greetingValue, bodyValue, signatureValue)

    If Not txtSignature Is Nothing Then
        AssignTemplateTextBoxValue txtSignature, vbNullString
    End If

    TraceTemplateLoad templateKey, toValue, ccValue, subjValue, bodyValue, signatureValue, combinedAttachments

    LoadEmailTemplateIntoControls = True
End Function

Private Function ResolveTemplateWorksheet() As Worksheet
    Dim candidateNames As Variant
    Dim candidate As Variant
    Dim ws As Worksheet

    On Error Resume Next
    If Not mTemplateWorksheet Is Nothing Then
        Dim tmpName As String
        tmpName = mTemplateWorksheet.Name
        If Err.Number = 0 Then
            Set ResolveTemplateWorksheet = mTemplateWorksheet
            On Error GoTo 0
            Exit Function
        End If
        Err.Clear
        Set mTemplateWorksheet = Nothing
    End If
    On Error GoTo 0

    candidateNames = Array(TEMPLATE_SHEET_NAME_PRIMARY, TEMPLATE_SHEET_NAME_ALT, TEMPLATE_SHEET_NAME_ALT2)

    For Each candidate In candidateNames
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(candidate))
        On Error GoTo 0
        If Not ws Is Nothing Then
            Set mTemplateWorksheet = ws
            Exit For
        End If
    Next candidate

    Set ResolveTemplateWorksheet = mTemplateWorksheet
End Function

Private Function NormaliseTemplateKey(ByVal templateKey As String, ByRef ws As Worksheet) As String
    Dim headerValue As String
    Dim candidate As String

    candidate = Trim$(templateKey)
    headerValue = ResolveTemplateColumnHeader(ws)

    If LenB(candidate) > 0 Then
        NormaliseTemplateKey = candidate
        Exit Function
    End If

    If LenB(headerValue) > 0 Then
        NormaliseTemplateKey = headerValue
    Else
        NormaliseTemplateKey = DEFAULT_TEMPLATE_KEY
    End If
End Function

Private Function ResolveTemplateColumnHeader(ByRef ws As Worksheet) As String
    If ws Is Nothing Then Exit Function

    ResolveTemplateColumnHeader = Trim$(CStrSafe(ws.Cells(1, TEMPLATE_COLUMN_INDEX).Value))
End Function

Private Function ResolveTemplateColumnIndex(ByRef ws As Worksheet, _
                                            ByVal templateKey As String) As Long
    If ws Is Nothing Then Exit Function

    If TEMPLATE_COLUMN_INDEX < 1 Then Exit Function
    If TEMPLATE_COLUMN_INDEX > ws.Columns.Count Then Exit Function

    Dim headerValue As String
    Dim searchRange As Range
    Dim candidate As Range

    headerValue = ResolveTemplateColumnHeader(ws)

    If StrComp(headerValue, templateKey, vbTextCompare) = 0 Then
        ResolveTemplateColumnIndex = TEMPLATE_COLUMN_INDEX
        Exit Function
    End If

    If LenB(headerValue) = 0 Then
        If StrComp(templateKey, DEFAULT_TEMPLATE_KEY, vbTextCompare) = 0 Then
            ResolveTemplateColumnIndex = TEMPLATE_COLUMN_INDEX
            Exit Function
        End If
    End If

    On Error Resume Next
    Set searchRange = Intersect(ws.Rows(1), ws.UsedRange)
    On Error GoTo 0

    If searchRange Is Nothing Then
        TraceTemplateColumnMismatch templateKey
        Exit Function
    End If

    For Each candidate In searchRange.Cells
        If candidate.Column <> TEMPLATE_COLUMN_INDEX Then
            headerValue = Trim$(CStrSafe(candidate.Value))
            If LenB(headerValue) > 0 Then
                If StrComp(headerValue, templateKey, vbTextCompare) = 0 Then
                    ResolveTemplateColumnIndex = candidate.Column
                    Exit Function
                End If
            End If
        End If
    Next candidate

    TraceTemplateColumnMismatch templateKey
End Function

Public Function TryGetTemplateDraftContent(ByVal templateKey As String, _
                                           ByRef ccValue As String, _
                                           ByRef subjectValue As String, _
                                           ByRef bodyValue As String) As Boolean
    Dim ws As Worksheet
    Dim templateColumn As Long
    Dim greetingValue As String
    Dim coreBodyValue As String
    Dim signatureValue As String

    Set ws = ResolveTemplateWorksheet()
    If ws Is Nothing Then Exit Function

    templateKey = NormaliseTemplateKey(templateKey, ws)

    templateColumn = ResolveTemplateColumnIndex(ws, templateKey)
    If templateColumn = 0 Then Exit Function

    ccValue = Trim$(CStrSafe(ws.Cells(EMAIL_ROW_CC, templateColumn).Value))
    subjectValue = Trim$(CStrSafe(ws.Cells(EMAIL_ROW_SUBJECT, templateColumn).Value))
    greetingValue = Trim$(CStrSafe(ws.Cells(EMAIL_ROW_GREETING, templateColumn).Value))
    coreBodyValue = Trim$(CStrSafe(ws.Cells(EMAIL_ROW_BODY, templateColumn).Value))
    signatureValue = Trim$(CStrSafe(ws.Cells(EMAIL_ROW_SIGNATURE, templateColumn).Value))

    bodyValue = BuildBodyValue(greetingValue, coreBodyValue, signatureValue)

    TryGetTemplateDraftContent = True
End Function

Private Function GetAttachmentExistsCache() As Object
    If mAttachmentExistsCache Is Nothing Then
        On Error Resume Next
        Set mAttachmentExistsCache = CreateObject("Scripting.Dictionary")
        If Err.Number = 0 Then
            mAttachmentExistsCache.CompareMode = vbTextCompare
        Else
            Set mAttachmentExistsCache = Nothing
            Err.Clear
        End If
        On Error GoTo 0
    End If

    Set GetAttachmentExistsCache = mAttachmentExistsCache
End Function

Private Sub ClearAttachmentExistenceCache()
    If mAttachmentExistsCache Is Nothing Then Exit Sub

    On Error Resume Next
    mAttachmentExistsCache.RemoveAll
    If Err.Number <> 0 Then
        Set mAttachmentExistsCache = Nothing
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Private Sub ClearTemplateControls(ByRef txtTO As MSForms.TextBox, _
                                  ByRef txtCC As MSForms.TextBox, _
                                  ByRef lstAT As MSForms.ListBox, _
                                  ByRef txtSubj As MSForms.TextBox, _
                                  ByRef txtBody As MSForms.TextBox, _
                                  ByRef txtSignature As MSForms.TextBox)
    AssignTemplateTextBoxValue txtTO, vbNullString
    AssignTemplateTextBoxValue txtCC, vbNullString
    ClearListBoxItems lstAT
    AssignTemplateTextBoxValue txtSubj, vbNullString
    AssignTemplateTextBoxValue txtBody, vbNullString
    AssignTemplateTextBoxValue txtSignature, vbNullString
End Sub

Private Sub AssignAttachmentList(ByRef target As MSForms.ListBox, ByVal entries As Collection)
    AssignListBoxItems target, entries
End Sub

Private Sub AssignTemplateTextBoxValue(ByRef target As MSForms.TextBox, ByVal value As String)
    If target Is Nothing Then Exit Sub
    target.Value = value
End Sub

Private Sub ClearListBoxItems(ByRef target As MSForms.ListBox)
    If target Is Nothing Then Exit Sub
    target.Clear
End Sub

Private Sub AssignListBoxItems(ByRef target As MSForms.ListBox, ByVal entries As Collection)
    Dim entry As Variant

    If target Is Nothing Then Exit Sub

    target.Clear

    If entries Is Nothing Then Exit Sub

    For Each entry In entries
        On Error Resume Next
        target.AddItem CStr(entry)
        On Error GoTo 0
    Next entry
End Sub

Private Sub TraceTemplateColumnMismatch(ByVal templateKey As String)
    If Not ENABLE_TEMPLATE_TRACE Then Exit Sub

    Debug.Print "[TemplateLoad] Unable to resolve template column for key '" & templateKey & "'."
End Sub

Private Sub TraceTemplateLoad(ByVal templateKey As String, _
                              ByVal toValue As String, _
                              ByVal ccValue As String, _
                              ByVal subjValue As String, _
                              ByVal bodyValue As String, _
                              ByVal signatureValue As String, _
                              ByVal attachments As Collection)
    Dim attachmentCount As Long

    If Not ENABLE_TEMPLATE_TRACE Then Exit Sub

    If Not attachments Is Nothing Then
        On Error Resume Next
        attachmentCount = attachments.Count
        On Error GoTo 0
    End If

    Debug.Print "[TemplateLoad] Key='" & templateKey & "' To='" & toValue & "' CC='" & ccValue & _
                "' Subject='" & subjValue & "' Attachments=" & attachmentCount & _
                " BodyLen=" & Len(bodyValue) & " SignatureLen=" & Len(signatureValue)
End Sub

Private Function CombineAttachmentCollections(ByVal templateEntries As Collection, _
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

    Set CombineAttachmentCollections = combined
End Function

'-------------------------------------------------------------------------------
' Procedure: GetAvailableTemplateKeys
' Purpose  : List all template column headers on the email template worksheet.
' Parameters: None.
' Returns  : Collection containing each distinct template key string, preserving sheet order.
' Side Effects:
'   None.
'-------------------------------------------------------------------------------
Public Function GetAvailableTemplateKeys() As Collection
    Dim ws As Worksheet
    Dim keys As Collection
    Dim headerValue As String

    Set ws = ResolveTemplateWorksheet()
    If ws Is Nothing Then Exit Function

    Set keys = New Collection

    headerValue = ResolveTemplateColumnHeader(ws)

    If LenB(headerValue) > 0 Then
        keys.Add headerValue
    Else
        keys.Add DEFAULT_TEMPLATE_KEY
    End If

    Set GetAvailableTemplateKeys = keys
End Function

Private Function BuildBodyValue(ByVal greetingValue As String, _
                                ByVal bodyValue As String, _
                                ByVal signatureValue As String) As String
    BuildBodyValue = modEmailPlaceholders.CombineTemplateSections(greetingValue, bodyValue, signatureValue)
End Function

'-------------------------------------------------------------------------------
' Procedure: AppendTemplateAttachments
' Purpose  : Add validated attachment entries to the template column without duplicating
'            existing records.
' Parameters:
'   templateKey - Column header identifying the template to update.
'   attachmentPaths - Collection of fully qualified file paths chosen by the user.
' Returns  : String representing the final serialized attachment entries written to the sheet.
' Side Effects:
'   Writes the updated attachment entry string back to the template worksheet; raises
'   descriptive errors when the worksheet or template column cannot be located.
'-------------------------------------------------------------------------------
Public Function AppendTemplateAttachments(ByVal templateKey As String, _
                                          ByVal attachmentPaths As Collection) As String
    Dim ws As Worksheet
    Dim templateColumn As Long
    Dim existingValue As String
    Dim existingEntries As Collection
    Dim resultEntries As Collection
    Dim seen As Object
    Dim idx As Long
    Dim entry As Variant
    Dim normalizedKey As String
    Dim selectedPath As Variant
    Dim newEntry As String
    Dim resolvedPath As String
    Dim displayName As String

    Set ws = ResolveTemplateWorksheet()
    If ws Is Nothing Then
        Err.Raise vbObjectError + 513, "modEmailTemplates.AppendTemplateAttachments", _
                  "Email template worksheet could not be found."
    End If

    templateKey = NormaliseTemplateKey(templateKey, ws)

    templateColumn = ResolveTemplateColumn(ws, templateKey)
    If templateColumn = 0 Then
        Err.Raise vbObjectError + 514, "modEmailTemplates.AppendTemplateAttachments", _
                  "The selected template could not be found on the EmailTemplate worksheet."
    End If

    existingValue = Trim$(CStrSafe(ws.Cells(EMAIL_ROW_ATTACHMENTS, templateColumn).Value))

    Set existingEntries = ParseAttachmentEntries(existingValue)
    Set resultEntries = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    If Not seen Is Nothing Then
        On Error Resume Next
        seen.CompareMode = vbTextCompare
        On Error GoTo 0
    End If

    For idx = 1 To existingEntries.Count
        entry = existingEntries(idx)
        normalizedKey = NormalizeAttachmentKey(CStr(entry))
        If LenB(normalizedKey) = 0 Then
            normalizedKey = UCase$(CStr(entry))
        End If
        If seen Is Nothing Then
            resultEntries.Add CStr(entry)
        ElseIf Not seen.Exists(normalizedKey) Then
            seen(normalizedKey) = True
            resultEntries.Add CStr(entry)
        End If
    Next idx

    If Not attachmentPaths Is Nothing Then
        For Each selectedPath In attachmentPaths
            resolvedPath = Trim$(CStr(selectedPath))
            displayName = vbNullString

            ' Ignore selections that fail validation; error messages are emitted upstream.
            If Not CheckIfAttachmentExists(displayName, resolvedPath) Then GoTo NextSelection

            normalizedKey = NormalizeAttachmentPath(resolvedPath)
            If LenB(normalizedKey) = 0 Then GoTo NextSelection
            If Not seen Is Nothing Then
                If seen.Exists(normalizedKey) Then GoTo NextSelection
                seen(normalizedKey) = True
            End If
            newEntry = BuildAttachmentEntryFromComponents(displayName, resolvedPath)
            If LenB(newEntry) = 0 Then GoTo NextSelection
            resultEntries.Add newEntry
NextSelection:
        Next selectedPath
    End If

    AppendTemplateAttachments = JoinAttachmentEntries(resultEntries)
    ws.Cells(EMAIL_ROW_ATTACHMENTS, templateColumn).Value = AppendTemplateAttachments
    ClearAttachmentExistenceCache
End Function

'-------------------------------------------------------------------------------
' Procedure: GetTemplateAttachmentEntries
' Purpose  : Parse a serialized attachment entry string into individual collection items.
' Parameters:
'   rawValue - Line- or semicolon-delimited attachment entry text.
' Returns  : Collection of parsed attachment entry strings (possibly empty).
' Side Effects:
'   None.
'-------------------------------------------------------------------------------
Public Function GetTemplateAttachmentEntries(ByVal rawValue As String) As Collection
    Set GetTemplateAttachmentEntries = ParseAttachmentEntries(rawValue)
End Function

'-------------------------------------------------------------------------------
' Procedure: NormalizeTemplateAttachmentPath
' Purpose  : Provide a consistent normalized representation of an attachment path.
' Parameters:
'   filePath - Raw attachment path.
' Returns  : Normalized attachment path string.
' Side Effects:
'   None.
'-------------------------------------------------------------------------------
Public Function NormalizeTemplateAttachmentPath(ByVal filePath As String) As String
    NormalizeTemplateAttachmentPath = NormalizeAttachmentPath(filePath)
End Function

'-------------------------------------------------------------------------------
' Procedure: BuildTemplateAttachmentEntry
' Purpose  : Convert a file path into a serialized attachment entry with display name.
' Parameters:
'   filePath - Fully qualified attachment path.
' Returns  : Attachment entry string suitable for persisting in the template sheet.
' Side Effects:
'   None.
'-------------------------------------------------------------------------------
Public Function BuildTemplateAttachmentEntry(ByVal filePath As String) As String
    BuildTemplateAttachmentEntry = BuildAttachmentEntry(filePath)
End Function

'-------------------------------------------------------------------------------
' Procedure: JoinTemplateAttachmentEntries
' Purpose  : Serialize attachment entries into the worksheet storage format.
' Parameters:
'   entries - Collection of attachment entry strings.
' Returns  : String that joins entries using newline delimiters.
' Side Effects:
'   None.
'-------------------------------------------------------------------------------
Public Function JoinTemplateAttachmentEntries(ByVal entries As Collection) As String
    JoinTemplateAttachmentEntries = JoinAttachmentEntries(entries)
End Function

'-------------------------------------------------------------------------------
' Procedure: NormalizeTemplateAttachmentEntry
' Purpose  : Produce a comparison-friendly key for attachment entries.
' Parameters:
'   entry - Serialized attachment entry.
' Returns  : Normalized entry string used for dictionary lookups.
' Side Effects:
'   None.
'-------------------------------------------------------------------------------
Public Function NormalizeTemplateAttachmentEntry(ByVal entry As String) As String
    NormalizeTemplateAttachmentEntry = NormalizeAttachmentKey(entry)
End Function

'-------------------------------------------------------------------------------
' Procedure: GetTemplateAttachmentEntriesForKey
' Purpose  : Retrieve and validate template-managed attachment entries for the given key.
' Parameters:
'   templateKey - Template column identifier.
' Returns  : Collection of attachment file paths that exist on disk.
' Side Effects:
'   Refreshes the worksheet value when validation cleans or normalizes entries.
'-------------------------------------------------------------------------------
Public Function GetTemplateAttachmentEntriesForKey(ByVal templateKey As String) As Collection
    Dim ws As Worksheet
    Dim templateColumn As Long
    Dim attachmentValue As String

    Set ws = ResolveTemplateWorksheet()
    If ws Is Nothing Then Exit Function

    templateKey = NormaliseTemplateKey(templateKey, ws)

    templateColumn = ResolveTemplateColumn(ws, templateKey)
    If templateColumn = 0 Then Exit Function

    attachmentValue = Trim$(CStrSafe(ws.Cells(EMAIL_ROW_ATTACHMENTS, templateColumn).Value))
    Set GetTemplateAttachmentEntriesForKey = ValidateTemplateAttachmentPaths(ws, templateColumn, attachmentValue)
End Function

'-------------------------------------------------------------------------------
' Procedure: GetUserAttachmentEntries
' Purpose  : Read persisted user-managed attachments associated with the template key.
' Parameters:
'   templateKey - Template column identifier.
' Returns  : Collection of attachment entry strings read from the worksheet.
' Side Effects:
'   None.
'-------------------------------------------------------------------------------
Public Function GetUserAttachmentEntries(ByVal templateKey As String) As Collection
    Dim ws As Worksheet
    Dim templateColumn As Long

    Set ws = ResolveTemplateWorksheet()
    If ws Is Nothing Then Exit Function

    templateKey = NormaliseTemplateKey(templateKey, ws)

    templateColumn = ResolveTemplateColumn(ws, templateKey)
    If templateColumn = 0 Then Exit Function

    Set GetUserAttachmentEntries = ReadUserAttachmentEntriesFromWorksheet(ws, templateColumn)
End Function

'-------------------------------------------------------------------------------
' Procedure: WriteUserAttachmentEntries
' Purpose  : Persist user-managed attachment entries into the template worksheet.
' Parameters:
'   templateKey - Template column identifier.
'   userEntries - Collection containing combined display/path entries supplied by the user.
' Returns  : None.
' Side Effects:
'   Writes file name and path components into dedicated worksheet rows.
'-------------------------------------------------------------------------------
Public Sub WriteUserAttachmentEntries(ByVal templateKey As String, _
                                      ByVal userEntries As Collection)
    Dim ws As Worksheet
    Dim templateColumn As Long
    Dim fileNames As Collection
    Dim filePaths As Collection

    Set ws = ResolveTemplateWorksheet()
    If ws Is Nothing Then Exit Sub

    templateKey = NormaliseTemplateKey(templateKey, ws)

    templateColumn = ResolveTemplateColumn(ws, templateKey)
    If templateColumn = 0 Then Exit Sub

    AppendAttachmentComponentCollections userEntries, fileNames, filePaths

    ws.Cells(EMAIL_ROW_USER_ATTACHMENT_NAMES, templateColumn).Value = JoinCollectionValues(fileNames)
    ws.Cells(EMAIL_ROW_USER_ATTACHMENT_PATHS, templateColumn).Value = JoinCollectionValues(filePaths)
    ClearAttachmentExistenceCache
End Sub

'-------------------------------------------------------------------------------
' Procedure: WriteTemplateAttachmentEntries
' Purpose  : Replace the template-managed attachment entry set for the given key.
' Parameters:
'   templateKey - Template column identifier.
'   attachmentEntries - Collection of validated attachment entry strings.
' Returns  : String written into the worksheet (joined entries).
' Side Effects:
'   Writes serialized attachment string to the Email Template worksheet and raises
'   descriptive errors if the sheet or column cannot be found.
'-------------------------------------------------------------------------------
Public Function WriteTemplateAttachmentEntries(ByVal templateKey As String, _
                                               ByVal attachmentEntries As Collection) As String
    Dim ws As Worksheet
    Dim templateColumn As Long
    Dim finalValue As String

    finalValue = JoinAttachmentEntries(attachmentEntries)

    Set ws = ResolveTemplateWorksheet()
    If ws Is Nothing Then
        Err.Raise vbObjectError + 515, "modEmailTemplates.WriteTemplateAttachmentEntries", _
                  "Email template worksheet could not be found."
    End If

    templateKey = NormaliseTemplateKey(templateKey, ws)

    templateColumn = ResolveTemplateColumn(ws, templateKey)
    If templateColumn = 0 Then
        Err.Raise vbObjectError + 516, "modEmailTemplates.WriteTemplateAttachmentEntries", _
                  "The selected template could not be found on the EmailTemplate worksheet."
    End If

    ws.Cells(EMAIL_ROW_ATTACHMENTS, templateColumn).Value = finalValue
    ClearAttachmentExistenceCache
    WriteTemplateAttachmentEntries = finalValue
End Function

'-------------------------------------------------------------------------------
' Procedure: GetValidatedTemplateAttachmentPaths
' Purpose  : Validate stored template attachment entries and return existing file paths.
' Parameters:
'   templateKey - Template column identifier.
' Returns  : Collection of file paths for attachments that currently exist.
' Side Effects:
'   Cleans worksheet values when entries are normalized or removed, ensuring downstream
'   consumers do not see stale or invalid references.
'-------------------------------------------------------------------------------
Public Function GetValidatedTemplateAttachmentPaths(ByVal templateKey As String) As Collection
    Dim ws As Worksheet
    Dim templateColumn As Long
    Dim originalValue As String
    Dim entries As Collection
    Dim validatedEntries As Collection
    Dim attachments As Collection
    Dim entry As Variant
    Dim entryValue As String
    Dim fileName As String
    Dim filePath As String
    Dim updatedEntry As String
    Dim finalValue As String
    Dim attachmentExists As Boolean
    Dim changed As Boolean

    Set ws = ResolveTemplateWorksheet()
    If ws Is Nothing Then Exit Function

    templateKey = NormaliseTemplateKey(templateKey, ws)

    templateColumn = ResolveTemplateColumn(ws, templateKey)
    If templateColumn = 0 Then Exit Function

    originalValue = Trim$(CStrSafe(ws.Cells(EMAIL_ROW_ATTACHMENTS, templateColumn).Value))
    If LenB(originalValue) = 0 Then Exit Function

    Set entries = ParseAttachmentEntries(originalValue)
    If entries Is Nothing Then Exit Function
    If entries.Count = 0 Then Exit Function

    Set validatedEntries = New Collection
    Set attachments = New Collection

    For Each entry In entries
        entryValue = CStr(entry)
        filePath = ExtractAttachmentPath(entryValue)
        If LenB(filePath) = 0 Then
            filePath = Trim$(entryValue)
        End If
        If LenB(filePath) = 0 Then
            changed = True
            GoTo NextEntry
        End If

        fileName = ExtractAttachmentEntryName(entryValue)
        If LenB(fileName) = 0 Then
            fileName = ExtractAttachmentFileName(filePath)
        End If

        ' Keep the path even if the file is missing so the worksheet can be corrected.
        attachmentExists = CheckIfAttachmentExists(fileName, filePath)
        If LenB(filePath) = 0 Then
            changed = True
            GoTo NextEntry
        End If

        If LenB(fileName) = 0 Then
            fileName = ExtractAttachmentFileName(filePath)
        End If

        updatedEntry = UpdateAttachmentEntry(entryValue, fileName, filePath)
        If StrComp(updatedEntry, entryValue, vbBinaryCompare) <> 0 Then
            changed = True
        End If
        validatedEntries.Add updatedEntry

        If attachmentExists Then
            attachments.Add filePath
        End If
NextEntry:
    Next entry

    If validatedEntries.Count > 0 Then
        finalValue = JoinAttachmentEntries(validatedEntries)
    Else
        finalValue = vbNullString
    End If

    If changed Or StrComp(originalValue, finalValue, vbBinaryCompare) <> 0 Then
        ws.Cells(EMAIL_ROW_ATTACHMENTS, templateColumn).Value = finalValue
    End If

    If attachments.Count > 0 Then
        Set GetValidatedTemplateAttachmentPaths = attachments
    End If
End Function

'-------------------------------------------------------------------------------
' Procedure: ResolveAttachmentPathsFromEntries
' Purpose  : Extract existing file paths from serialized attachment entries.
' Parameters:
'   entries - Collection of attachment entry strings.
' Returns  : Collection of unique, validated file paths corresponding to existing files.
' Side Effects:
'   None, although invalid entries are skipped and therefore excluded from the result.
'-------------------------------------------------------------------------------
Public Function ResolveAttachmentPathsFromEntries(ByVal entries As Collection) As Collection
    Dim attachments As Collection
    Dim entry As Variant
    Dim entryValue As String
    Dim fileName As String
    Dim filePath As String
    Dim replacementEntry As String
    Dim attachmentExists As Boolean
    Dim normalizedKey As String
    Dim seen As Object

    If entries Is Nothing Then Exit Function
    If entries.Count = 0 Then Exit Function

    Set attachments = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    If Not seen Is Nothing Then
        On Error Resume Next
        seen.CompareMode = vbTextCompare
        On Error GoTo 0
    End If

    For Each entry In entries
        entryValue = CStr(entry)
        filePath = ExtractAttachmentPath(entryValue)
        If LenB(filePath) = 0 Then
            filePath = Trim$(entryValue)
        End If
        If LenB(filePath) = 0 Then GoTo NextEntry

        fileName = ExtractAttachmentEntryName(entryValue)
        If LenB(fileName) = 0 Then
            fileName = ExtractAttachmentFileName(filePath)
        End If

        attachmentExists = AttachmentFileExists(filePath)

        If Not attachmentExists Then
            replacementEntry = vbNullString
            If HandleMissingAttachment(filePath, replacementEntry) Then
                If LenB(replacementEntry) > 0 Then
                    filePath = ExtractAttachmentPath(replacementEntry)
                    If LenB(filePath) = 0 Then
                        filePath = Trim$(replacementEntry)
                    End If
                    fileName = ExtractAttachmentEntryName(replacementEntry)
                    If LenB(fileName) = 0 Then
                        fileName = ExtractAttachmentFileName(filePath)
                    End If
                    attachmentExists = AttachmentFileExists(filePath)
                End If
            End If
        End If

        If Not attachmentExists Then GoTo NextEntry

        normalizedKey = NormalizeAttachmentPath(filePath)
        If Not seen Is Nothing Then
            If seen.Exists(normalizedKey) Then GoTo NextEntry
            seen(normalizedKey) = True
        End If

        attachments.Add filePath
NextEntry:
    Next entry

    If attachments.Count > 0 Then
        Set ResolveAttachmentPathsFromEntries = attachments
    End If
End Function

'-------------------------------------------------------------------------------
' Procedure: CheckIfAttachmentExists
' Purpose  : Confirm whether a supplied attachment path can be resolved, optionally
'            updating it when a replacement is available.
' Parameters:
'   fileName - In/out parameter holding the display name for the attachment.
'   filePath - In/out parameter holding the attachment path to validate.
' Returns  : True when the attachment can be resolved; False when it cannot.
' Side Effects:
'   May modify fileName and filePath to reflect recovered attachment information.
'-------------------------------------------------------------------------------
Public Function CheckIfAttachmentExists(ByRef fileName As String, _
                                        ByRef filePath As String) As Boolean
    Dim replacementEntry As String
    Dim newPath As String
    Dim newName As String

    fileName = Trim$(fileName)
    filePath = Trim$(filePath)

    If LenB(filePath) = 0 Then Exit Function

    If AttachmentFileExists(filePath) Then
        If LenB(fileName) = 0 Then
            fileName = ExtractAttachmentFileName(filePath)
        End If
        CheckIfAttachmentExists = True
        Exit Function
    End If

    If HandleMissingAttachment(filePath, replacementEntry) Then
        If LenB(replacementEntry) > 0 Then
            newPath = ExtractAttachmentPath(replacementEntry)
            If LenB(newPath) > 0 Then
                filePath = newPath
                newName = ExtractAttachmentFileName(newPath)
                If LenB(newName) > 0 Then
                    fileName = newName
                ElseIf LenB(fileName) = 0 Then
                    fileName = newPath
                End If
                If AttachmentFileExists(filePath) Then
                    CheckIfAttachmentExists = True
                End If
            End If
        End If
    Else
        fileName = vbNullString
        filePath = vbNullString
    End If
End Function

Private Function ResolveTemplateColumn(ByVal ws As Worksheet, ByVal templateKey As String) As Long
    ResolveTemplateColumn = ResolveTemplateColumnIndex(ws, templateKey)
End Function

Private Function ReadUserAttachmentEntriesFromWorksheet(ByVal ws As Worksheet, _
                                                        ByVal templateColumn As Long) As Collection
    Dim nameValues As Collection
    Dim pathValues As Collection
    Dim entryName As String
    Dim entryPath As String
    Dim entry As String
    Dim maxCount As Long
    Dim idx As Long

    If ws Is Nothing Then Exit Function
    If templateColumn <= 0 Then Exit Function

    Set nameValues = CollectTemplateAttachmentValues(CStrSafe(ws.Cells(EMAIL_ROW_USER_ATTACHMENT_NAMES, templateColumn).Value))
    Set pathValues = CollectTemplateAttachmentValues(CStrSafe(ws.Cells(EMAIL_ROW_USER_ATTACHMENT_PATHS, templateColumn).Value))

    maxCount = pathValues.Count
    If nameValues.Count > maxCount Then
        maxCount = nameValues.Count
    End If

    If maxCount = 0 Then Exit Function

    Set ReadUserAttachmentEntriesFromWorksheet = New Collection

    For idx = 1 To maxCount
        entryName = vbNullString
        entryPath = vbNullString
        If idx <= nameValues.Count Then entryName = CStr(nameValues(idx))
        If idx <= pathValues.Count Then entryPath = CStr(pathValues(idx))
        entryName = Trim$(entryName)
        entryPath = Trim$(entryPath)
        If LenB(entryPath) = 0 And LenB(entryName) = 0 Then GoTo NextEntry
        If LenB(entryPath) = 0 Then entryPath = entryName
        entry = BuildAttachmentEntryFromComponents(entryName, entryPath)
        If LenB(entry) > 0 Then
            ReadUserAttachmentEntriesFromWorksheet.Add entry
        End If
NextEntry:
    Next idx
End Function

Private Function ParseAttachmentEntries(ByVal rawValue As String) As Collection
    Dim entries As Collection
    Dim normalized As String
    Dim parts() As String
    Dim part As Variant
    Dim subParts() As String
    Dim subPart As Variant

    Set entries = New Collection
    If LenB(rawValue) = 0 Then
        Set ParseAttachmentEntries = entries
        Exit Function
    End If

    normalized = Replace(rawValue, vbCrLf, vbLf)
    normalized = Replace(normalized, vbCr, vbLf)
    parts = Split(normalized, vbLf)

    For Each part In parts
        part = Trim$(CStr(part))
        If LenB(part) = 0 Then GoTo NextPart
        If InStr(part, ";") > 0 Then
            subParts = Split(part, ";")
            For Each subPart In subParts
                subPart = Trim$(CStr(subPart))
                If LenB(subPart) > 0 Then entries.Add CStr(subPart)
            Next subPart
        Else
            entries.Add CStr(part)
        End If
NextPart:
    Next part

    If entries.Count = 0 Then
        If InStr(rawValue, ";") > 0 Then
            subParts = Split(rawValue, ";")
            For Each subPart In subParts
                subPart = Trim$(CStr(subPart))
                If LenB(subPart) > 0 Then entries.Add CStr(subPart)
            Next subPart
        ElseIf LenB(Trim$(rawValue)) > 0 Then
            entries.Add Trim$(rawValue)
        End If
    End If

    Set ParseAttachmentEntries = entries
End Function

Private Function CollectTemplateAttachmentValues(ByVal rawValue As String) As Collection
    Dim items As Collection
    Dim normalized As String
    Dim parts() As String
    Dim part As Variant

    Set items = New Collection

    rawValue = Trim$(rawValue)
    If LenB(rawValue) = 0 Then
        Set CollectTemplateAttachmentValues = items
        Exit Function
    End If

    normalized = Replace(rawValue, vbCrLf, vbLf)
    normalized = Replace(normalized, vbCr, vbLf)
    normalized = Replace(normalized, ";", vbLf)
    normalized = Replace(normalized, ",", vbLf)

    parts = Split(normalized, vbLf)
    For Each part In parts
        part = Trim$(CStr(part))
        If LenB(part) > 0 Then items.Add CStr(part)
    Next part

    Set CollectTemplateAttachmentValues = items
End Function

Private Function JoinCollectionValues(ByVal items As Collection) As String
    Dim arr() As String
    Dim idx As Long

    If items Is Nothing Then Exit Function
    If items.Count = 0 Then Exit Function

    ReDim arr(1 To items.Count)
    For idx = 1 To items.Count
        arr(idx) = Trim$(CStr(items(idx)))
    Next idx

    JoinCollectionValues = Join(arr, vbCrLf)
End Function

Private Sub AppendAttachmentComponentCollections(ByVal entries As Collection, _
                                                 ByRef fileNames As Collection, _
                                                 ByRef filePaths As Collection)
    Dim entry As Variant
    Dim entryValue As String
    Dim fileName As String
    Dim filePath As String

    If entries Is Nothing Then Exit Sub

    For Each entry In entries
        entryValue = CStr(entry)
        filePath = ExtractAttachmentPath(entryValue)
        If LenB(filePath) = 0 Then
            filePath = Trim$(entryValue)
        End If
        If LenB(filePath) = 0 Then GoTo NextEntry

        fileName = ExtractAttachmentEntryName(entryValue)
        If LenB(fileName) = 0 Then
            fileName = ExtractAttachmentFileName(filePath)
        End If

        If fileNames Is Nothing Then Set fileNames = New Collection
        If filePaths Is Nothing Then Set filePaths = New Collection

        fileNames.Add fileName
        filePaths.Add filePath
NextEntry:
    Next entry
End Sub

Private Function NormalizeAttachmentKey(ByVal entry As String) As String
    Dim pathValue As String

    pathValue = ExtractAttachmentPath(entry)
    If LenB(pathValue) > 0 Then
        NormalizeAttachmentKey = NormalizeAttachmentPath(pathValue)
    Else
        NormalizeAttachmentKey = UCase$(Trim$(entry))
    End If
End Function

Private Function NormalizeAttachmentPath(ByVal filePath As String) As String
    NormalizeAttachmentPath = UCase$(Trim$(filePath))
End Function

Private Function BuildAttachmentEntry(ByVal filePath As String) As String
    Dim fileName As String

    fileName = ExtractAttachmentFileName(filePath)
    If LenB(fileName) > 0 Then
        BuildAttachmentEntry = fileName & " | " & Trim$(filePath)
    Else
        BuildAttachmentEntry = Trim$(filePath)
    End If
End Function

'-------------------------------------------------------------------------------
' Procedure: BuildAttachmentEntryFromComponents
' Purpose  : Combine the provided display name and path into the standard entry format.
' Parameters:
'   fileName - Friendly name shown in the UI (optional).
'   filePath - Fully qualified file path (required).
' Returns  : Serialized attachment entry string or empty when the path is blank.
' Side Effects:
'   None.
'-------------------------------------------------------------------------------
Public Function BuildAttachmentEntryFromComponents(ByVal fileName As String, _
                                                   ByVal filePath As String) As String
    fileName = Trim$(fileName)
    filePath = Trim$(filePath)

    If LenB(filePath) = 0 Then Exit Function

    If LenB(fileName) = 0 Then
        BuildAttachmentEntryFromComponents = BuildAttachmentEntry(filePath)
    Else
        BuildAttachmentEntryFromComponents = fileName & " | " & filePath
    End If
End Function

Private Function ExtractAttachmentFileName(ByVal filePath As String) As String
    Dim pos As Long

    filePath = Trim$(filePath)
    If LenB(filePath) = 0 Then Exit Function

    pos = InStrRev(filePath, Application.PathSeparator)
    If pos > 0 Then
        ExtractAttachmentFileName = Mid$(filePath, pos + 1)
    Else
        ExtractAttachmentFileName = filePath
    End If
End Function

Private Function ExtractAttachmentPath(ByVal entry As String) As String
    Dim separatorPos As Long

    entry = Trim$(entry)
    If LenB(entry) = 0 Then Exit Function

    separatorPos = InStr(entry, "|")
    If separatorPos > 0 Then
        ExtractAttachmentPath = Trim$(Mid$(entry, separatorPos + 1))
        Exit Function
    End If

    separatorPos = InStr(entry, ": ")
    If separatorPos > 0 Then
        ExtractAttachmentPath = Trim$(Mid$(entry, separatorPos + 2))
        Exit Function
    End If

    If InStr(entry, Application.PathSeparator) > 0 Or InStr(entry, ":") > 0 Then
        ExtractAttachmentPath = entry
    End If
End Function

Private Function ExtractAttachmentEntryName(ByVal entry As String) As String
    Dim separatorPos As Long

    entry = Trim$(entry)
    If LenB(entry) = 0 Then Exit Function

    separatorPos = InStr(entry, "|")
    If separatorPos > 0 Then
        ExtractAttachmentEntryName = Trim$(Left$(entry, separatorPos - 1))
        Exit Function
    End If

    separatorPos = InStr(entry, ": ")
    If separatorPos > 0 Then
        ExtractAttachmentEntryName = Trim$(Left$(entry, separatorPos - 1))
        Exit Function
    End If

    If InStr(entry, Application.PathSeparator) = 0 And InStr(entry, ":") = 0 Then
        ExtractAttachmentEntryName = entry
    End If
End Function

Private Function UpdateAttachmentEntry(ByVal originalEntry As String, _
                                       ByVal fileName As String, _
                                       ByVal filePath As String) As String
    Dim separatorPos As Long
    Dim trimmedEntry As String

    trimmedEntry = Trim$(originalEntry)
    fileName = Trim$(fileName)
    filePath = Trim$(filePath)

    If LenB(filePath) = 0 Then Exit Function

    separatorPos = InStr(trimmedEntry, "|")
    If separatorPos > 0 Then
        If LenB(fileName) = 0 Then
            fileName = ExtractAttachmentFileName(filePath)
        End If
        UpdateAttachmentEntry = fileName & " | " & filePath
        Exit Function
    End If

    separatorPos = InStr(trimmedEntry, ": ")
    If separatorPos > 0 Then
        If LenB(fileName) = 0 Then
            fileName = ExtractAttachmentFileName(filePath)
        End If
        UpdateAttachmentEntry = fileName & ": " & filePath
        Exit Function
    End If

    If LenB(fileName) > 0 Then
        UpdateAttachmentEntry = fileName & " | " & filePath
    Else
        UpdateAttachmentEntry = BuildAttachmentEntry(filePath)
    End If
End Function

Private Function JoinAttachmentEntries(ByVal entries As Collection) As String
    Dim arr() As String
    Dim idx As Long

    If entries Is Nothing Then Exit Function
    If entries.Count = 0 Then Exit Function

    ReDim arr(1 To entries.Count)
    For idx = 1 To entries.Count
        arr(idx) = entries(idx)
    Next idx

    JoinAttachmentEntries = Join(arr, vbCrLf)
End Function

Private Function ValidateTemplateAttachmentPaths(ByVal ws As Worksheet, _
                                                ByVal templateColumn As Long, _
                                                ByVal rawValue As String) As Collection
    Dim entries As Collection
    Dim updatedEntries As Collection
    Dim entry As Variant
    Dim pathValue As String
    Dim newEntry As String
    Dim resultValue As String
    Dim changed As Boolean

    Set entries = ParseAttachmentEntries(rawValue)
    If entries Is Nothing Then Exit Function

    Set updatedEntries = New Collection

    For Each entry In entries
        pathValue = ExtractAttachmentPath(CStr(entry))
        If LenB(pathValue) = 0 Then
            updatedEntries.Add CStr(entry)
        ElseIf AttachmentFileExists(pathValue) Then
            updatedEntries.Add CStr(entry)
        ElseIf HandleMissingAttachment(pathValue, newEntry) Then
            ' Allow custom recovery hooks to substitute a new location without losing the entry.
            If LenB(newEntry) > 0 Then
                updatedEntries.Add newEntry
                changed = True
            Else
                updatedEntries.Add CStr(entry)
            End If
        Else
            changed = True
        End If
    Next entry

    If updatedEntries Is Nothing Then
        resultValue = vbNullString
    Else
        resultValue = JoinAttachmentEntries(updatedEntries)
    End If

    If changed Or StrComp(resultValue, rawValue, vbBinaryCompare) <> 0 Then
        ws.Cells(EMAIL_ROW_ATTACHMENTS, templateColumn).Value = resultValue
    End If

    Set ValidateTemplateAttachmentPaths = updatedEntries
End Function

Private Function AttachmentFileExists(ByVal filePath As String) As Boolean
    Dim resolvedPath As String
    Dim normalizedPath As String
    Dim cache As Object

    filePath = Trim$(filePath)
    If LenB(filePath) = 0 Then Exit Function

    normalizedPath = NormalizeAttachmentPath(filePath)
    Set cache = GetAttachmentExistsCache()

    If Not cache Is Nothing Then
        On Error Resume Next
        If cache.Exists(normalizedPath) Then
            AttachmentFileExists = CBool(cache(normalizedPath))
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0
    End If

    On Error Resume Next
    resolvedPath = Dir(filePath, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive)
    On Error GoTo 0

    AttachmentFileExists = LenB(resolvedPath) > 0

    If Not cache Is Nothing Then
        On Error Resume Next
        cache(normalizedPath) = AttachmentFileExists
        On Error GoTo 0
    End If
End Function

Private Function ResolveActiveTemplateKeyFromForm() As String
    Dim frm As Object
    Dim resolved As String

    For Each frm In VBA.UserForms
        If StrComp(TypeName(frm), "EmailForm", vbTextCompare) = 0 Then
            resolved = ResolveEmailFormActiveKey(frm)
            Exit For
        End If
    Next frm

    resolved = Trim$(resolved)

    If LenB(resolved) = 0 Then
        Dim ws As Worksheet
        Set ws = ResolveTemplateWorksheet()
        If ws Is Nothing Then
            resolved = DEFAULT_TEMPLATE_KEY
        Else
            resolved = NormaliseTemplateKey(resolved, ws)
        End If
    End If

    ResolveActiveTemplateKeyFromForm = resolved
End Function

Private Function ResolveEmailFormActiveKey(ByVal frm As Object) As String
    Dim resolved As String

    If frm Is Nothing Then Exit Function

    On Error Resume Next
    resolved = Trim$(CStr(CallByName(frm, "ActiveTemplateKey", VbMethod, False)))
    If LenB(resolved) = 0 Then
        resolved = Trim$(CStr(CallByName(frm, "ActiveTemplateKey", VbMethod, True)))
    End If
    On Error GoTo 0

    If LenB(resolved) = 0 Then
        resolved = TryGetFormControlValue(frm, "txtTEMP")
    End If

    ResolveEmailFormActiveKey = resolved
End Function

Private Function TryGetFormControlValue(ByVal targetForm As Object, _
                                        ByVal controlName As String) As String
    Dim ctrl As Object
    Dim resolved As String

    controlName = Trim$(controlName)
    If targetForm Is Nothing Then Exit Function
    If LenB(controlName) = 0 Then Exit Function

    On Error Resume Next
    Set ctrl = targetForm.Controls(controlName)
    If Err.Number <> 0 Then
        Err.Clear
        Set ctrl = Nothing
    End If
    On Error GoTo 0

    If ctrl Is Nothing Then Exit Function

    On Error Resume Next
    resolved = Trim$(CStr(ctrl.Value))
    If Err.Number <> 0 Then
        Err.Clear
        resolved = Trim$(CStr(ctrl.Text))
    End If
    On Error GoTo 0

    TryGetFormControlValue = resolved
End Function

Private Sub UpdateWorksheetAttachmentsForReplacement(ByVal missingPath As String, _
                                                     ByVal replacementPath As String)
    Dim templateKey As String
    Dim ws As Worksheet
    Dim templateColumn As Long
    Dim normalizedMissing As String
    Dim replacementName As String
    Dim attachmentEntries As Collection
    Dim updatedEntries As Collection
    Dim entry As Variant
    Dim entryValue As String
    Dim entryPath As String
    Dim candidateKey As String
    Dim templateUpdated As Boolean
    Dim nameValues As Collection
    Dim pathValues As Collection
    Dim updatedNameValues As Collection
    Dim updatedPathValues As Collection
    Dim maxCount As Long
    Dim idx As Long
    Dim entryName As String
    Dim userUpdated As Boolean

    replacementPath = Trim$(replacementPath)
    If LenB(replacementPath) = 0 Then Exit Sub

    templateKey = ResolveActiveTemplateKeyFromForm()

    Set ws = ResolveTemplateWorksheet()
    If ws Is Nothing Then Exit Sub

    templateKey = NormaliseTemplateKey(templateKey, ws)

    templateColumn = ResolveTemplateColumn(ws, templateKey)
    If templateColumn = 0 Then Exit Sub

    normalizedMissing = NormalizeAttachmentPath(missingPath)
    If LenB(normalizedMissing) = 0 Then Exit Sub

    replacementName = ExtractAttachmentFileName(replacementPath)
    If LenB(replacementName) = 0 Then
        replacementName = replacementPath
    End If

    Set attachmentEntries = ParseAttachmentEntries(CStrSafe(ws.Cells(EMAIL_ROW_ATTACHMENTS, templateColumn).Value))
    If Not attachmentEntries Is Nothing Then
        If attachmentEntries.Count > 0 Then
            Set updatedEntries = New Collection
            For Each entry In attachmentEntries
                entryValue = CStr(entry)
                entryPath = ExtractAttachmentPath(entryValue)
                If LenB(entryPath) = 0 Then
                    entryPath = Trim$(entryValue)
                End If
                candidateKey = NormalizeAttachmentPath(entryPath)
                If LenB(candidateKey) > 0 And _
                   StrComp(candidateKey, normalizedMissing, vbTextCompare) = 0 Then
                    updatedEntries.Add BuildAttachmentEntry(replacementPath)
                    templateUpdated = True
                Else
                    updatedEntries.Add entryValue
                End If
            Next entry
            If templateUpdated Then
                ws.Cells(EMAIL_ROW_ATTACHMENTS, templateColumn).Value = JoinAttachmentEntries(updatedEntries)
            End If
        End If
    End If

    Set nameValues = CollectTemplateAttachmentValues(CStrSafe(ws.Cells(EMAIL_ROW_USER_ATTACHMENT_NAMES, templateColumn).Value))
    Set pathValues = CollectTemplateAttachmentValues(CStrSafe(ws.Cells(EMAIL_ROW_USER_ATTACHMENT_PATHS, templateColumn).Value))

    maxCount = pathValues.Count
    If nameValues.Count > maxCount Then
        maxCount = nameValues.Count
    End If

    If maxCount = 0 Then Exit Sub

    Set updatedNameValues = New Collection
    Set updatedPathValues = New Collection

    For idx = 1 To maxCount
        entryName = vbNullString
        entryPath = vbNullString

        If idx <= nameValues.Count Then entryName = Trim$(CStr(nameValues(idx)))
        If idx <= pathValues.Count Then entryPath = Trim$(CStr(pathValues(idx)))

        If LenB(entryPath) = 0 And LenB(entryName) > 0 Then
            entryPath = entryName
        End If

        If LenB(entryPath) = 0 And LenB(entryName) = 0 Then GoTo NextEntry

        candidateKey = NormalizeAttachmentPath(entryPath)
        If LenB(candidateKey) > 0 And _
           StrComp(candidateKey, normalizedMissing, vbTextCompare) = 0 Then
            entryPath = replacementPath
            entryName = replacementName
            userUpdated = True
        ElseIf LenB(entryName) = 0 Then
            entryName = ExtractAttachmentFileName(entryPath)
            If LenB(entryName) = 0 Then
                entryName = entryPath
            End If
        End If

        updatedNameValues.Add entryName
        updatedPathValues.Add entryPath
NextEntry:
    Next idx

    If userUpdated Then
        ws.Cells(EMAIL_ROW_USER_ATTACHMENT_NAMES, templateColumn).Value = JoinCollectionValues(updatedNameValues)
        ws.Cells(EMAIL_ROW_USER_ATTACHMENT_PATHS, templateColumn).Value = JoinCollectionValues(updatedPathValues)
    End If

    If templateUpdated Or userUpdated Then
        ClearAttachmentExistenceCache
    End If
End Sub

Private Function HandleMissingAttachment(ByVal missingPath As String, _
                                         ByRef replacementEntry As String) As Boolean
    Dim response As VbMsgBoxResult
    Dim fd As FileDialog
    Dim selectedPath As String

    replacementEntry = vbNullString

    response = modUIHelpers.ShowDecisionMessage( _
        "AUTO_SPCL couldn't find the attachment '" & missingPath & "'." & vbCrLf & vbCrLf & _
        "Select Yes to locate the file or No to remove it from the template.", _
        vbYesNo + vbQuestion, _
        "Attachment Not Found")

    If response = vbYes Then
        On Error Resume Next
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        On Error GoTo 0
        If fd Is Nothing Then
            HandleMissingAttachment = True
            Exit Function
        End If

        With fd
            .AllowMultiSelect = False
            .Title = "Select replacement attachment"
            On Error Resume Next
            .InitialFileName = missingPath
            On Error GoTo 0
            .Filters.Clear
            .Filters.Add "All Files", "*.*"
            If .Show = -1 Then
                If .SelectedItems.Count > 0 Then
                    selectedPath = Trim$(CStr(.SelectedItems(1)))
                End If
            End If
        End With

        If LenB(selectedPath) > 0 Then
            replacementEntry = BuildAttachmentEntry(selectedPath)
            UpdateWorksheetAttachmentsForReplacement missingPath, selectedPath
        End If

        Set fd = Nothing
        HandleMissingAttachment = True
    ElseIf response = vbNo Then
        HandleMissingAttachment = False
    End If
End Function

Private Function CStrSafe(ByVal value As Variant) As String
    If IsError(value) Then Exit Function
    If IsNull(value) Then Exit Function
    If IsEmpty(value) Then Exit Function
    CStrSafe = Trim$(CStr(value))
End Function
