Attribute VB_Name = "modEmailTemplates"
Option Explicit

Private Const TEMPLATE_SHEET_NAME_PRIMARY As String = "EmailTemplate"
Private Const TEMPLATE_SHEET_NAME_ALT As String = "EmailTemplates"
Private Const TEMPLATE_SHEET_NAME_ALT2 As String = "Email Templates"

Private Const EMAIL_ROW_TO As Long = 2
Private Const EMAIL_ROW_CC As Long = 3
Private Const EMAIL_ROW_SUBJECT As Long = 4
Private Const EMAIL_ROW_BODY As Long = 5
Private Const EMAIL_ROW_GREETING As Long = 6
Private Const EMAIL_ROW_SIGNATURE As Long = 7
Private Const EMAIL_ROW_ATTACHMENTS As Long = 9

Public Sub LoadEmailTemplateData(ByVal templateKey As String, _
                                 ByRef txtTO As MSForms.TextBox, _
                                 ByRef txtCC As MSForms.TextBox, _
                                 ByRef txtAT As MSForms.TextBox, _
                                 ByRef txtSubj As MSForms.TextBox, _
                                 ByRef txtBody As MSForms.TextBox, _
                                 ByRef txtSignature As MSForms.TextBox)
    Dim ws As Worksheet
    Dim lastCol As Long
    Dim colIndex As Long
    Dim templateColumn As Long
    Dim headerValue As String
    Dim toValue As String
    Dim ccValue As String
    Dim subjValue As String
    Dim bodyValue As String
    Dim greetingValue As String
    Dim signatureValue As String
    Dim attachmentValue As String

    If LenB(templateKey) = 0 Then Exit Sub

    Set ws = ResolveTemplateWorksheet()
    If ws Is Nothing Then Exit Sub

    ClearTemplateControls txtTO, txtCC, txtAT, txtSubj, txtBody, txtSignature

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then Exit Sub

    For colIndex = 1 To lastCol
        headerValue = Trim$(CStrSafe(ws.Cells(1, colIndex).Value))
        If LenB(headerValue) > 0 Then
            If StrComp(headerValue, templateKey, vbTextCompare) = 0 Then
                templateColumn = colIndex
                Exit For
            End If
        End If
    Next colIndex

    If templateColumn = 0 Then Exit Sub

    toValue = Trim$(CStrSafe(ws.Cells(EMAIL_ROW_TO, templateColumn).Value))
    ccValue = Trim$(CStrSafe(ws.Cells(EMAIL_ROW_CC, templateColumn).Value))
    subjValue = Trim$(CStrSafe(ws.Cells(EMAIL_ROW_SUBJECT, templateColumn).Value))
    bodyValue = Trim$(CStrSafe(ws.Cells(EMAIL_ROW_BODY, templateColumn).Value))
    greetingValue = Trim$(CStrSafe(ws.Cells(EMAIL_ROW_GREETING, templateColumn).Value))
    signatureValue = Trim$(CStrSafe(ws.Cells(EMAIL_ROW_SIGNATURE, templateColumn).Value))
    attachmentValue = Trim$(CStrSafe(ws.Cells(EMAIL_ROW_ATTACHMENTS, templateColumn).Value))
    attachmentValue = ValidateTemplateAttachmentPaths(ws, templateColumn, attachmentValue)

    AssignTextBoxValue txtTO, toValue
    AssignTextBoxValue txtCC, ccValue
    AssignTextBoxValue txtSubj, subjValue
    AssignTextBoxValue txtSignature, signatureValue
    AssignTextBoxValue txtAT, attachmentValue
    AssignTextBoxValue txtBody, BuildBodyValue(greetingValue, bodyValue)
End Sub

Private Function ResolveTemplateWorksheet() As Worksheet
    On Error Resume Next
    Set ResolveTemplateWorksheet = ThisWorkbook.Worksheets(TEMPLATE_SHEET_NAME_PRIMARY)
    If ResolveTemplateWorksheet Is Nothing Then
        Set ResolveTemplateWorksheet = ThisWorkbook.Worksheets(TEMPLATE_SHEET_NAME_ALT)
    End If
    If ResolveTemplateWorksheet Is Nothing Then
        Set ResolveTemplateWorksheet = ThisWorkbook.Worksheets(TEMPLATE_SHEET_NAME_ALT2)
    End If
    On Error GoTo 0
End Function

Private Sub ClearTemplateControls(ByRef txtTO As MSForms.TextBox, _
                                  ByRef txtCC As MSForms.TextBox, _
                                  ByRef txtAT As MSForms.TextBox, _
                                  ByRef txtSubj As MSForms.TextBox, _
                                  ByRef txtBody As MSForms.TextBox, _
                                  ByRef txtSignature As MSForms.TextBox)
    AssignTextBoxValue txtTO, vbNullString
    AssignTextBoxValue txtCC, vbNullString
    AssignTextBoxValue txtAT, vbNullString
    AssignTextBoxValue txtSubj, vbNullString
    AssignTextBoxValue txtBody, vbNullString
    AssignTextBoxValue txtSignature, vbNullString
End Sub

Private Sub AssignTextBoxValue(ByRef target As MSForms.TextBox, ByVal value As String)
    If target Is Nothing Then Exit Sub
    target.Value = value
End Sub

Private Function BuildBodyValue(ByVal greetingValue As String, _
                                ByVal bodyValue As String) As String
    If LenB(greetingValue) = 0 Then
        BuildBodyValue = bodyValue
        Exit Function
    End If

    BuildBodyValue = greetingValue
    If LenB(bodyValue) > 0 Then
        BuildBodyValue = BuildBodyValue & vbCrLf & vbCrLf & bodyValue
    End If
End Function

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

    If LenB(templateKey) = 0 Then Exit Function

    Set ws = ResolveTemplateWorksheet()
    If ws Is Nothing Then
        Err.Raise vbObjectError + 513, "modEmailTemplates.AppendTemplateAttachments", _
                  "Email template worksheet could not be found."
    End If

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
            normalizedKey = NormalizeAttachmentPath(CStr(selectedPath))
            If LenB(normalizedKey) = 0 Then GoTo NextSelection
            If Not seen Is Nothing Then
                If seen.Exists(normalizedKey) Then GoTo NextSelection
                seen(normalizedKey) = True
            End If
            newEntry = BuildAttachmentEntry(CStr(selectedPath))
            resultEntries.Add newEntry
NextSelection:
        Next selectedPath
    End If

    AppendTemplateAttachments = JoinAttachmentEntries(resultEntries)
    ws.Cells(EMAIL_ROW_ATTACHMENTS, templateColumn).Value = AppendTemplateAttachments
End Function

Public Function GetTemplateAttachmentEntries(ByVal rawValue As String) As Collection
    Set GetTemplateAttachmentEntries = ParseAttachmentEntries(rawValue)
End Function

Public Function NormalizeTemplateAttachmentPath(ByVal filePath As String) As String
    NormalizeTemplateAttachmentPath = NormalizeAttachmentPath(filePath)
End Function

Public Function BuildTemplateAttachmentEntry(ByVal filePath As String) As String
    BuildTemplateAttachmentEntry = BuildAttachmentEntry(filePath)
End Function

Public Function JoinTemplateAttachmentEntries(ByVal entries As Collection) As String
    JoinTemplateAttachmentEntries = JoinAttachmentEntries(entries)
End Function

Public Function NormalizeTemplateAttachmentEntry(ByVal entry As String) As String
    NormalizeTemplateAttachmentEntry = NormalizeAttachmentKey(entry)
End Function

Public Function WriteTemplateAttachmentEntries(ByVal templateKey As String, _
                                               ByVal attachmentEntries As Collection) As String
    Dim ws As Worksheet
    Dim templateColumn As Long
    Dim finalValue As String

    finalValue = JoinAttachmentEntries(attachmentEntries)

    If LenB(templateKey) = 0 Then
        WriteTemplateAttachmentEntries = finalValue
        Exit Function
    End If

    Set ws = ResolveTemplateWorksheet()
    If ws Is Nothing Then
        Err.Raise vbObjectError + 515, "modEmailTemplates.WriteTemplateAttachmentEntries", _
                  "Email template worksheet could not be found."
    End If

    templateColumn = ResolveTemplateColumn(ws, templateKey)
    If templateColumn = 0 Then
        Err.Raise vbObjectError + 516, "modEmailTemplates.WriteTemplateAttachmentEntries", _
                  "The selected template could not be found on the EmailTemplate worksheet."
    End If

    ws.Cells(EMAIL_ROW_ATTACHMENTS, templateColumn).Value = finalValue
    WriteTemplateAttachmentEntries = finalValue
End Function

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

    If LenB(templateKey) = 0 Then Exit Function

    Set ws = ResolveTemplateWorksheet()
    If ws Is Nothing Then Exit Function

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
    Dim lastCol As Long
    Dim colIndex As Long
    Dim headerValue As String

    If ws Is Nothing Then Exit Function
    If LenB(templateKey) = 0 Then Exit Function

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then Exit Function

    For colIndex = 1 To lastCol
        headerValue = Trim$(CStrSafe(ws.Cells(1, colIndex).Value))
        If LenB(headerValue) = 0 Then GoTo NextColumn
        If StrComp(headerValue, templateKey, vbTextCompare) = 0 Then
            ResolveTemplateColumn = colIndex
            Exit Function
        End If
NextColumn:
    Next colIndex
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
                                                ByVal rawValue As String) As String
    Dim entries As Collection
    Dim updatedEntries As Collection
    Dim entry As Variant
    Dim pathValue As String
    Dim newEntry As String
    Dim resultValue As String
    Dim changed As Boolean

    Set entries = ParseAttachmentEntries(rawValue)
    If entries Is Nothing Then
        ValidateTemplateAttachmentPaths = rawValue
        Exit Function
    End If

    Set updatedEntries = New Collection

    For Each entry In entries
        pathValue = ExtractAttachmentPath(CStr(entry))
        If LenB(pathValue) = 0 Then
            updatedEntries.Add CStr(entry)
        ElseIf AttachmentFileExists(pathValue) Then
            updatedEntries.Add CStr(entry)
        ElseIf HandleMissingAttachment(pathValue, newEntry) Then
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

    ValidateTemplateAttachmentPaths = resultValue
End Function

Private Function AttachmentFileExists(ByVal filePath As String) As Boolean
    Dim resolvedPath As String

    filePath = Trim$(filePath)
    If LenB(filePath) = 0 Then Exit Function

    On Error Resume Next
    resolvedPath = Dir(filePath, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive)
    On Error GoTo 0

    AttachmentFileExists = LenB(resolvedPath) > 0
End Function

Private Function HandleMissingAttachment(ByVal missingPath As String, _
                                         ByRef replacementEntry As String) As Boolean
    Dim response As VbMsgBoxResult
    Dim fd As FileDialog
    Dim selectedPath As String

    replacementEntry = vbNullString

    response = MsgBox("The attachment '" & missingPath & "' could not be found." & _
                      vbCrLf & vbCrLf & _
                      "Would you like to locate the file?" & vbCrLf & _
                      "Click Yes to find the file or No to remove it from the template.", _
                      vbYesNo + vbQuestion, "Attachment Not Found")

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
