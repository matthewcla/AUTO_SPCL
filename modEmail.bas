Attribute VB_Name = "modEmail"
Option Explicit

' ==== CONFIGURE THESE ====
Private Const CC_LIST As String = "alex.d.schneider.mil@us.navy.mil"  ' hard-code CCs here (semicolon-separated)
Private Const SUBJECT_TEMPLATE As String = "CDR CMD Record Review"

' Use {Name} and {EligiblesNote} placeholders in the template.
' Add any other fixed lines you need. Use vbCrLf for new lines.
Private Const BODY_TEMPLATE As String = _
    "Greetings from Millington!" & vbCrLf & vbCrLf & _
    "We are in the process of doing our initial round of record reviews for the FY-27 CDR CMD board taking place from 8-12 December 2025.  Below are items that require your immediate attention from our internal review; however, we also request you review your personal OSR/PSR in BOL for accuracy and to ensure alignment with anything we may find missing or inaccurate." & vbCrLf & vbCrLf & _
    "Noted issues: " & vbCrLf & vbCrLf & _
    "{EligiblesNote}" & vbCrLf & vbCrLf & _
    "I'm here to assist you in resolving the above discrepancies.  Let me know if any information is incorrect and we will work to resolve the issues before the board in December." & vbCrLf & vbCrLf & _
    "If you have missing AQDs I am able to add those to your record, but will require appropriate documentation (with the exception of Joint-related AQDs).  For any FITREP related issues, you will need to work with MNCC (askmncc@navy.mil) to resolve those issues and anything that can’t be resolved before the board must be submitted via BOL as a letter to the board (LTB) – if in doubt, just submit a LTB as backup to be sure.  As a detailer, I am unable to fix FITREP issues, or upload FITREPs into your OMPF/record." & vbCrLf & vbCrLf & _
    "Keep in mind that any items submitted as a LTB will not permanently be fixed in your record but only temporarily for that particular board.  LTBs must be submitted via BOL ESSBD NLT 28 November to be accepted, and request you also email me a copy of your submitted LTB for our internal records since we do not have access to your LTB as detailers." & vbCrLf & vbCrLf & _
    "Also be aware that we do not know what awards you've received throughout your career, so this is an item we are unable to verify on your behalf.  If any awards are missing from your OSR just submit as a LTB to fix for the board." & vbCrLf & vbCrLf & _
    "MyNavyHR contains a plethora of helpful information regarding all things career and record related.  I highly recommend you utilize MyNavyHR to answer any immediate questions you may have.  The below links may be particularly useful.  The second link takes you to the PERS-41 homepage --> Click the [Officer Record Management Guide] hyperlink on the right side of the page for a great record management tool." & vbCrLf & vbCrLf & _
    "https://www.mynavyhr.navy.mil/Career-Management/Detailing/Officer/Pers-41-SWO/Detailers/410-411/" & vbCrLf & vbCrLf & _
    "https://www.mynavyhr.navy.mil/Career-Management/Detailing/Officer/Pers-41-SWO/" & vbCrLf & vbCrLf & _
    "Let me know if you have any questions or require further guidance and/or assistance.  Thank you, and have a great day!" & vbCrLf & _
    "Very respectfully," & vbCrLf & _
    "Alex"

Public Sub CreateDraftsFromID()
    Dim wsID As Worksheet, wsElig As Worksheet
    Dim lastRow As Long, r As Long
    Dim personName As String, toList As String, eligNote As String
    Dim olApp As Object, olMail As Object  ' Outlook.Application / MailItem (late bound)
    Dim createdCount As Long, skippedCount As Long
    
    On Error GoTo CleanFail
    
    Set wsID = ThisWorkbook.Worksheets("ID")
    Set wsElig = ThisWorkbook.Worksheets("Eligibles RED Board")
    
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
    
    Application.ScreenUpdating = False
    
    For r = 2 To lastRow
        personName = Trim$(wsID.Cells(r, "B").Value)
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
            .Save            ' <-- creates draft in Outlook Drafts
            ' .Display       ' (intentionally NOT displayed to keep it hidden)
        End With
        createdCount = createdCount + 1
nextRow:
    Next r
    
    Application.ScreenUpdating = True
    MsgBox "Draft creation complete." & vbCrLf & _
           "Created: " & createdCount & vbCrLf & _
           "Skipped (no name or no emails): " & skippedCount, vbInformation
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


