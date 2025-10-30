Attribute VB_Name = "modEmail"
Option Explicit

Public Sub CreateDraftsFromID(Optional ByVal allowedMembers As Variant, _
                              Optional ByVal templateKey As String = vbNullString, _
                              Optional ByVal userAttachments As Variant)
    Dim wsID As Worksheet, wsElig As Worksheet
    Dim lastRow As Long, r As Long
    Dim personName As String, toList As String, eligNote As String
    Dim olApp As Object, olMail As Object  ' Outlook.Application / MailItem (late bound)
    Dim createdCount As Long, skippedCount As Long
    Dim whitelist As Object
    Dim hasWhitelist As Boolean
    Dim memberIndex As Long
    Dim skipNote As String
    Dim templateAttachments As Collection
    Dim userAttachmentPaths As Collection
    Dim attachmentPath As Variant
    
    On Error GoTo CleanFail
    
    Set wsID = ThisWorkbook.Worksheets("ID")
    Set wsElig = ThisWorkbook.Worksheets("Eligibles RED Board")

    If Not IsMissing(allowedMembers) Then
        Set whitelist = NormalizeDraftWhitelist(allowedMembers)
        hasWhitelist = Not whitelist Is Nothing
    End If
    
    lastRow = wsID.Cells(wsID.Rows.Count, "B").End(xlUp).row
    If lastRow < 2 Then
        MsgBox "No data rows found on 'ID' (need names in column B).", vbExclamation
        Exit Sub
    End If
    
    ' Get or start Outlook
    On Error Resume Next
    Set olApp = GetObject(Class:="Outlook.Application")
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    On Error GoTo CleanFail
    If olApp Is Nothing Then
        MsgBox "Unable to start Outlook.", vbCritical
        Exit Sub
    End If

    If LenB(templateKey) > 0 Then
        Set templateAttachments = GetValidatedTemplateAttachmentPaths(templateKey)
    End If

    If Not IsMissing(userAttachments) Then
        If IsObject(userAttachments) Then
            On Error Resume Next
            Set userAttachmentPaths = userAttachments
            If Err.Number <> 0 Then
                Err.Clear
                Set userAttachmentPaths = Nothing
            End If
            On Error GoTo CleanFail
        End If
    End If

    Application.ScreenUpdating = False

    For r = 2 To lastRow
        memberIndex = r - 1
        personName = Trim$(wsID.Cells(r, "B").Value)

        If hasWhitelist Then
            If Not DraftWhitelistAllowsMember(memberIndex, personName, whitelist) Then
                skippedCount = skippedCount + 1
                GoTo nextRow
            End If
        End If

        If Len(personName) = 0 Then
            skippedCount = skippedCount + 1
            GoTo nextRow
        End If
        
        ' Build To: from columns C:F (semicolon-separated)
        toList = BuildEmailList(wsID, r, "C", "F")
        If Len(toList) = 0 Then
            ' No valid email addresses found for this row
            skippedCount = skippedCount + 1
            GoTo nextRow
        End If
        
        ' Lookup note from Eligibles col A -> take col C
        eligNote = GetEligiblesNote(wsElig, personName)
        
        ' Create the draft (hidden; saved to Drafts)
        Set olMail = olApp.CreateItem(0) ' olMailItem = 0
        With olMail
            .To = toList
            .CC = CC_LIST  ' hard-coded CCs (modify above)
            .Subject = Replace(SUBJECT_TEMPLATE, "{Name}", personName)
            .Body = BuildBody(personName, eligNote)
            If Not templateAttachments Is Nothing Then
                For Each attachmentPath In templateAttachments
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
nextRow:
    Next r
    
    Application.ScreenUpdating = True
    If hasWhitelist Then skipNote = " (including members not marked as Draft)"
    MsgBox "Draft creation complete." & vbCrLf & _
           "Created: " & createdCount & vbCrLf & _
           "Skipped (no name, no emails, or filtered out): " & skippedCount & skipNote, vbInformation
    Exit Sub
    
CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

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

' Build the email body by replacing placeholders in BODY_TEMPLATE.
Private Function BuildBody(ByVal personName As String, ByVal eligNote As String) As String
    Dim bodyText As String
    Dim noteText As String
    Dim replacements As Variant

    bodyText = BODY_TEMPLATE

    If LenB(eligNote) > 0 Then
        noteText = eligNote
    Else
        noteText = "(no note found)"
    End If

    replacements = Array( _
        "Name", personName, _
        "EligiblesNote", noteText, _
        "ISSUES", noteText, _
        "ISSUE", noteText _
    )

    bodyText = ReplacePlaceholdersArray(bodyText, replacements)
    BuildBody = bodyText
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




