Option Explicit
'========================================================================================
' Module: modRecordReview
' Purpose: Drives the OAIS record review pipeline and writes outputs to:
'          - "Eligibles Status Board" (wsSB)
'          - "Eligibles RED Board"   (wsRB)
'          Shows a modeless progress form via modProgressUI and DOES NOT create drafts.
'
' External dependencies expected elsewhere in project:
'   - ConnectToRunningOAIS, ChangeScreen, entText, HitF8, iCS.GetText, ParseYYYYMMDD
'   - Progress_Show, Progress_Update, Progress_Log, Progress_WaitIfPaused,
'     Progress_Cancelled, Progress_Close   (from modProgressUI)
'   - ResetRunState, LastRun* variables                              (from modReviewState)
'
' Notes:
'   - This module keeps the loop index "i" at MODULE scope so worker subs can see it.
'   - "CreateDraftsFromID" is NOT called here (you will trigger it from another form later).
'========================================================================================

'-------------------------
' Public run metrics
'-------------------------
Public processed As Long
Public total As Long
Public BoardType As String      ' kept for compatibility if used elsewhere

'-------------------------
' Module state used across worker subs
'-------------------------
Private issues As String
Private nIssue As String
Private IssueCAT As String
Private currentIssues As Collection

Private MasLook As String
Private BacLook As String

Private arrayID As Variant          ' 2D array [row, 1..2] => ID, Name
Private wsSB As Worksheet           ' "Eligibles Status Board"
Private wsRB As Worksheet           ' "Eligibles RED Board"

Private i As Long                   ' current candidate index (module-scoped on purpose)
Private vo5FY As Integer            ' promotion FY derived in lookINFO

'========================================================================================
' ENTRY POINT
'========================================================================================
Public Sub A_Record_Review(Optional ByVal Reserved As Boolean = False)
    ' Runs the full pipeline for all IDs present on sheet "ID".
    ' Shows modeless progress UI, supports pause/cancel, and records run-state for later.

    Dim iLo As Long, iHi As Long
    Dim nm As String, id As String
    Dim runWasCancelled As Boolean
    Dim progressClosed As Boolean

    On Error GoTo CleanFail
    ResetRunState

    processed = 0
    total = 0

    ' Progress UI (show immediately so it renders before any heavy work)
    Progress_Show total, "Record Review Progress"
    Progress_Log "Loading IDs..."

    Application.ScreenUpdating = False
    Application.EnableCancelKey = xlErrorHandler

    ConnectToRunningOAIS

    Set wsRB = ThisWorkbook.Worksheets("Eligibles RED Board")
    Set wsSB = ThisWorkbook.Worksheets("Eligibles Status Board")

    ' Load ID/Name pairs into memory
    SetID

    If IsEmpty(arrayID) Then
        total = 0
    Else
        iLo = LBound(arrayID, 1)
        iHi = UBound(arrayID, 1)
        total = IIf(iHi >= iLo, iHi - iLo + 1, 0)
    End If

    processed = 0

    ' Refresh the totals now that they are known
    Progress_Update processed, total, IIf(total = 1, "Loaded 1 ID.", "Loaded " & total & " IDs.")

    If total = 0 Then
        Progress_Log "No IDs available for processing."
        GoTo NoWork
    End If

    Set currentIssues = Nothing

    ' Main review loop (module-level i is intentional so workers can use it)
    For i = iLo To iHi
        If Not Progress_WaitIfPaused() Then Exit For
        If Progress_Cancelled() Then Exit For  ' This exits the whole loop if cancelled

        ' Process each record...
        id = Trim$(CStr(arrayID(i, 1)))
        nm = Trim$(CStr(arrayID(i, 2)))
        Set currentIssues = New Collection
        Progress_Log "Starting: " & IIf(Len(nm) > 0, nm, id) & "  [" & id & "]"

        '=== Pipeline ===
        If Progress_Cancelled() Then Exit For
        lookINFO
        If Progress_Cancelled() Then Exit For
        lookMASTER
        If Progress_Cancelled() Then Exit For
        lookBACHELOR
        If Progress_Cancelled() Then Exit For
        lookAQD
        If Progress_Cancelled() Then Exit For
        lookFITREP
        '================
    
        processed = processed + 1
        Progress_Update processed, total, BuildFinishedStatus(IIf(Len(nm) > 0, nm, id))
        Set currentIssues = Nothing
        DoEvents
    Next i

    Set currentIssues = Nothing

NoWork:
    runWasCancelled = Progress_Cancelled()

    If runWasCancelled Then
        Progress_Close "Cancelled by user."
        progressClosed = True
    Else
        Progress_Update processed, total, "Review complete."
        Progress_Close "Review complete.", True
        progressClosed = True
    End If

CleanOK:
    Application.ScreenUpdating = True

    ' Record run-state for your follow-on userform (which will trigger CreateDraftsFromID)
    LastRunProcessed = processed
    LastRunTotal = total
    If Not progressClosed Then
        runWasCancelled = Progress_Cancelled()
        Progress_Close "Terminated due to error."
        progressClosed = True
    End If

    LastRunWasCancelled = runWasCancelled
    LastRunCompleted = Not LastRunWasCancelled
    Exit Sub

CleanFail:
    Progress_Log "ERROR: " & Err.Number & " - " & Err.Description
    Resume CleanOK
End Sub

'========================================================================================
' INPUT LOADING
'========================================================================================
Private Sub SetID()
    ' Loads the "ID" sheet Column A:B into arrayID as a 1-based 2D array:
    '   arrayID(r, 1) = ID
    '   arrayID(r, 2) = NAME

    Dim ws As Worksheet
    Dim eRow As Long

    Set ws = ThisWorkbook.Sheets("ID")

    eRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    If eRow < 2 Then
        arrayID = Empty
        Exit Sub
    End If

    arrayID = ws.Range("A2:B" & eRow).Value
End Sub

'========================================================================================
' WORKER: INFO (PER1/PER2)
'========================================================================================
Private Sub lookINFO()
    ' Pulls identification, screening, flag, command, date/email info from PER1
    ' and CDR promotion FY from PER2. Writes to wsSB Row (i+1) and logs to wsRB as needed.

    Dim vYG As String, vDESIG As String, vSCRN As String
    Dim vFLG As String, vRRET As String
    Dim vcCMD As String, vCBIL As String, vuCMD As String
    Dim vEdd As String, vEda As String, vAUIC As String
    Dim NFAAS As String, oNSIPS As String, hNFAAS As String, hNSIPS As String

    Dim id As String
    Dim lk As String                          ' selection bucket derived at the end
    Dim dT As String
    Dim j As Long
    Dim yearPart As Integer, monthPart As Integer
    Dim fiscalYear As Integer
    Dim eddYear As Integer, eddMonth As Integer
    Dim edaYear As Integer, edaMonth As Integer

    id = CStr(arrayID(i, 1))
    
    If Progress_Cancelled() Then Exit Sub
    
    ChangeScreen "PER1"

    ' Write SSN to PER1
    entText 4, 41, id

    ' YEAR GROUP
    vYG = Format$(Trim$(iCS.GetText(6, 21, 3)), "@")
    wsSB.Cells(i + 1, 3).Value = vYG

    ' DESIGNATOR
    vDESIG = Format$(Trim$(iCS.GetText(4, 73, 5)), "@")
    wsSB.Cells(i + 1, 4).Value = vDESIG

    ' SCREENING CODE
    vSCRN = Format$(Trim$(iCS.GetText(7, 41, 5)), "@")
    wsSB.Cells(i + 1, 5).Value = vSCRN

    ' PERS-8 FLAGS
    If InStr(Format$(Trim$(iCS.GetText(5, 71, 8)), "@"), "8") > 0 Then
        vFLG = "Y"
        wsSB.Cells(i + 1, 6).Value = vFLG

        nIssue = Format$(Trim$(iCS.GetText(5, 71, 8)), "@")
        IssueCAT = "PERS-8 Flags"
        writeRB
    Else
        vFLG = "N"
        wsSB.Cells(i + 1, 6).Value = vFLG
    End If

    ' EDD -> as first day of YYMM month (expanded to full year)
    vEdd = Format$(Trim$(iCS.GetText(12, 73, 4)), "@")
    If vEdd <> vbNullString Then
        eddYear = ExpandTwoDigitYear(CInt(VBA.Left$(vEdd, 2)))
        eddMonth = CInt(Right$(vEdd, 2))
        wsSB.Cells(i + 1, 11).Value = DateSerial(eddYear, eddMonth, 1)
    End If

    ' RESIGNATION/RETIREMENT indicator
    If Format$(Trim$(iCS.GetText(12, 17, 16)), "@") = "SEPARATIONS TPPH" Then
        vRRET = "Y"
        wsSB.Cells(i + 1, 7).Value = vRRET

        nIssue = "SEPARATIONS TPPH"
        IssueCAT = "RESIG/RETIRE"
        writeRB
    Else
        vRRET = "N"
        wsSB.Cells(i + 1, 7).Value = vRRET
    End If

    ' CURRENT CMD/ BILLET / ULTIMATE CMD / AUIC
    vcCMD = Format$(Trim$(iCS.GetText(9, 17, 18)), "@")
    vAUIC = Format$(Trim$(iCS.GetText(9, 11, 5)), "@")
    wsSB.Cells(i + 1, 9).Value = vcCMD

    vCBIL = Format$(Trim$(iCS.GetText(10, 17, 18)), "@")
    wsSB.Cells(i + 1, 10).Value = vCBIL

    vuCMD = Format$(Trim$(iCS.GetText(12, 17, 16)), "@")
    wsSB.Cells(i + 1, 12).Value = vuCMD

    ' EDA -> as first day of YYMM month (expanded to full year)
    vEda = Format$(Trim$(iCS.GetText(12, 58, 4)), "@")
    If vEda <> vbNullString Then
        edaYear = ExpandTwoDigitYear(CInt(VBA.Left$(vEda, 2)))
        edaMonth = CInt(Right$(vEda, 2))
        wsSB.Cells(i + 1, 13).Value = DateSerial(edaYear, edaMonth, 1)
    End If

    ' EMAILS (also written back to "ID" sheet cols C:F)
    NFAAS = LCase$(Format$(Trim$(iCS.GetText(16, 8, 60)), "@"))
    ThisWorkbook.Worksheets("ID").Cells(i + 1, 3).Value = NFAAS

    oNSIPS = LCase$(Format$(Trim$(iCS.GetText(18, 17, 60)), "@"))
    ThisWorkbook.Worksheets("ID").Cells(i + 1, 4).Value = oNSIPS

    HitF8

    hNFAAS = LCase$(Format$(Trim$(iCS.GetText(17, 13, 60)), "@"))
    ThisWorkbook.Worksheets("ID").Cells(i + 1, 5).Value = hNFAAS

    hNSIPS = LCase$(Format$(Trim$(iCS.GetText(18, 13, 60)), "@"))
    ThisWorkbook.Worksheets("ID").Cells(i + 1, 6).Value = hNSIPS

    ' O-5 PROMOTION FY from PER2
    ChangeScreen "PER2"
    dT = vbNullString
    For j = 7 To 11
        If Format$(Trim$(iCS.GetText(j, 4, 3)), "@") = "CDR" Then
            dT = Format$(Trim$(iCS.GetText(j, 12, 6)), "@") ' YYMMDD (we only care about YYMM)
            Exit For
        End If
    Next j

    If Len(dT) >= 4 Then
        yearPart = CInt(VBA.Left$(dT, 2))
        monthPart = CInt(Mid$(dT, 3, 2))

        ' Fiscal year: months 10..12 roll to next year
        If monthPart >= 10 Then
            fiscalYear = yearPart + 1
        Else
            fiscalYear = yearPart
        End If

        vo5FY = fiscalYear
        wsSB.Cells(i + 1, 8).Value = vo5FY
    End If

    ' Selection bucket "lk" from screening/command/FY
    Dim f As Range
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("CMD_UIC")

    Set f = ws.Columns("A").Find(What:=vAUIC, LookIn:=xlValues, LookAt:=xlWhole, _
                                 SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)

    If f Is Nothing Then
        ' Not in MAJ CMD / FLT UP
        If Right$(vSCRN, 4) = "AKFZ" Then lk = "BNK"
        If Right$(vSCRN, 4) = "AKNZ" Then lk = "BNK"
        If Right$(vSCRN, 4) = "AKPZ" Then lk = "BNK"
        If Right$(vSCRN, 4) = "AKAZ" Then lk = "BNK"
    Else
        lk = "SEQ"
    End If

    If lk = vbNullString Then
        ' Fall back by year-from-promo-FY relative to current year (two-digit logic)
        If CInt(vo5FY) + 5 = Val(Right$(CStr(Year(Date)), 2)) Then lk = "1L"
        If CInt(vo5FY) + 5 = Val(Right$(CStr(Year(Date)), 2)) - 1 Then lk = "2L"
        If CInt(vo5FY) + 5 = Val(Right$(CStr(Year(Date)), 2)) - 2 Then lk = "3L"
    End If
    wsSB.Cells(i + 1, 2).Value = lk
End Sub

'========================================================================================
' WORKER: MASTERS (ACED)
'========================================================================================
Private Sub lookMASTER()
    ' Looks for Masters degree code on ACED; logs RED BOARD deficiency if missing.
    Dim j As Long
    
    If Progress_Cancelled() Then Exit Sub
    
    If Not (Format$(Trim$(iCS.GetText(1, 2, 4)), "@") = "ACED") Then ChangeScreen "ACED"

    For j = 12 To 15
        If Trim$(iCS.GetText(j, 30, 1)) = "8" Then
            MasLook = "Y"
            wsSB.Cells(i + 1, 15).Value = MasLook
            Exit Sub
        End If
    Next j

    wsSB.Cells(i + 1, 15).Value = "N"
    nIssue = "Not Documented"
    IssueCAT = "Masters Degree"
    writeRB
End Sub

'========================================================================================
' WORKER: BACHELOR (ACED)
'========================================================================================
Private Sub lookBACHELOR()
    ' Looks for Bachelor degree code on ACED; logs RED BOARD deficiency if missing.
    Dim j As Long
    
    If Progress_Cancelled() Then Exit Sub
    
    If Not (Format$(Trim$(iCS.GetText(1, 2, 4)), "@") = "ACED") Then ChangeScreen "ACED"

    For j = 12 To 15
        If Trim$(iCS.GetText(j, 30, 1)) = "6" Then
            BacLook = "Y"
            wsSB.Cells(i + 1, 14).Value = BacLook
            Exit Sub
        End If
    Next j

    wsSB.Cells(i + 1, 14).Value = "N"
    nIssue = "Not Documented"
    IssueCAT = "Bachelors Degree"
    writeRB
End Sub

'========================================================================================
' WORKER: AQDs (PRQS)
'========================================================================================
Private Sub lookAQD()
    ' Scans PRQS page rows 11..19 for AQDs and writes Y/N flags to wsSB.
    ' Logs RED BOARD lines if required quals are missing.

    Dim r As Long, cT As Long
    Dim rowText As String
    Dim rowCache() As String

    Dim WTI As Boolean, JI As Boolean, JII As Boolean, JT As Boolean
    Dim PCC As Boolean, PCA As Boolean, CQ As Boolean, SWCLA As Boolean
    Dim TAO As Boolean, EOOW As Boolean

    Dim codesWTI As Variant
    Dim codesJoint As Variant
    Dim codesTAO As Variant
    Dim codesEOOW As Variant
    
    If Progress_Cancelled() Then Exit Sub

    If Not Trim$(iCS.GetText(1, 2, 4)) = "PRQS" Then ChangeScreen "PRQS"

    ReDim rowCache(11 To 19)
    For r = 11 To 19
        rowCache(r) = Format$(Trim$(iCS.GetText(r, 1, 80)), "@")
    Next r

    codesWTI = Array(" KW1 ", " KW2 ", " KW3 ", " KW4 ", " KWC ")
    codesJoint = Array(" JS1 ", " JS2 ")
    codesTAO = Array(" LF6 ", " LF7 ")
    codesEOOW = Array(" LC1 ", " LC2 ", " LC3 ", " LC6 ", " LC7 ", " LC8 ", " LC9 ", " KD2 ")

    For r = 11 To 19
        rowText = rowCache(r)

        If Not WTI Then
            For cT = LBound(codesWTI) To UBound(codesWTI)
                If InStr(rowText, CStr(codesWTI(cT))) > 0 Then
                    WTI = True
                    Exit For
                End If
            Next cT
        End If

        If Not JI Then If InStr(rowText, " JS7 ") > 0 Then JI = True
        If Not JII Then If InStr(rowText, " JS8 ") > 0 Then JII = True

        If Not JT Then
            For cT = LBound(codesJoint) To UBound(codesJoint)
                If InStr(rowText, CStr(codesJoint(cT))) > 0 Then
                    JT = True
                    Exit For
                End If
            Next cT
        End If

        If Not PCC Then If InStr(rowText, " LN3 ") > 0 Then PCC = True
        If Not PCA Then If InStr(rowText, " LN4 ") > 0 Then PCA = True
        If Not CQ Then If InStr(rowText, " LN7 ") > 0 Then CQ = True
        If Not SWCLA Then If InStr(rowText, " 2D1 ") > 0 Then SWCLA = True

        If Not TAO Then
            For cT = LBound(codesTAO) To UBound(codesTAO)
                If InStr(rowText, CStr(codesTAO(cT))) > 0 Then
                    TAO = True
                    Exit For
                End If
            Next cT
        End If

        If Not EOOW Then
            For cT = LBound(codesEOOW) To UBound(codesEOOW)
                If InStr(rowText, CStr(codesEOOW(cT))) > 0 Then
                    EOOW = True
                    Exit For
                End If
            Next cT
        End If

        If WTI And JI And JII And JT And PCC And PCA And CQ And SWCLA And TAO And EOOW Then Exit For
    Next r

    ' Write Y/N results and log deficiencies
    wsSB.Cells(i + 1, 16).Value = IIf(WTI, "Y", "N")

    If JI Then
        wsSB.Cells(i + 1, 17).Value = "Y"
    Else
        wsSB.Cells(i + 1, 17).Value = "N": nIssue = "No Record of Completion": IssueCAT = "JPME Phase I": writeRB
    End If

    If JII Then
        wsSB.Cells(i + 1, 18).Value = "Y"
    Else
        wsSB.Cells(i + 1, 18).Value = "N": nIssue = "No Record of Completion": IssueCAT = "JPME Phase II": writeRB
    End If

    If JT Then
        wsSB.Cells(i + 1, 19).Value = "Y"
    Else
        wsSB.Cells(i + 1, 19).Value = "N": nIssue = "No Record of Completion": IssueCAT = "22 month Joint Req": writeRB
    End If

    If PCC Then
        wsSB.Cells(i + 1, 20).Value = "Y"
    Else
        wsSB.Cells(i + 1, 20).Value = "N": nIssue = "No Record of Completion": IssueCAT = "CDR CMD Tour": writeRB
    End If

    If PCA Then
        wsSB.Cells(i + 1, 21).Value = "Y"
    Else
        wsSB.Cells(i + 1, 21).Value = "N": nIssue = "No Record of Completion": IssueCAT = "CAPT CMD Tour": writeRB
    End If

    If CQ Then
        wsSB.Cells(i + 1, 22).Value = "Y"
    Else
        wsSB.Cells(i + 1, 22).Value = "N": nIssue = "No Record of Completion": IssueCAT = "Command Qual": writeRB
    End If

    If SWCLA Then
        wsSB.Cells(i + 1, 23).Value = "Y"
    Else
        wsSB.Cells(i + 1, 23).Value = "N": nIssue = "No Record of Completion": IssueCAT = "SWCLA": writeRB
    End If

    If TAO Then
        wsSB.Cells(i + 1, 24).Value = "Y"
    Else
        wsSB.Cells(i + 1, 24).Value = "N": nIssue = "No Record of Completion": IssueCAT = "TAO Qual": writeRB
    End If

    If EOOW Then
        wsSB.Cells(i + 1, 25).Value = "Y"
    Else
        wsSB.Cells(i + 1, 25).Value = "N": nIssue = "No Record of Completion": IssueCAT = "EOOW Qual": writeRB
    End If
End Sub

'========================================================================================
' WORKER: FITREP (OFT2)
'========================================================================================
Private Sub lookFITREP()
    ' Navigates OFT2, reads recent FITREP date blocks, detects >30-day gaps, flags OCT FITREP,
    ' and writes indicators to wsSB plus RED BOARD entries when gaps found.

    Dim wsID As Worksheet
    Dim gap As Boolean
    Dim outRow As Long, oRow As Long
    Dim eRow As Long
    Dim first As Boolean

    Dim vF1 As String, vT1 As String, vF2 As String, vT2 As String, vF3 As String, vT3 As String
    Dim vF4 As String, vT4 As String

    Dim dF1 As Double, dT1 As Double, dF2 As Double, dT2 As Double, dF3 As Double, dT3 As Double
    Dim dF4 As Double, dT4 As Double

    ConnectToRunningOAIS
    Set wsID = thisworkbook.Worksheets("ID")
    
    If Progress_Cancelled() Then Exit Sub
    
    outRow = 2
    oRow = 2 ' BASE_ROW external if you use it elsewhere; ensure consistent usage

    eRow = wsID.Cells(wsID.Rows.Count, "A").End(xlUp).Row

    If Not Trim$(iCS.GetText(1, 2, 4)) = "OFT2" Then ChangeScreen "OFT2"

    Application.ScreenUpdating = False

    first = True
    Do
        nIssue = vbNullString
        IssueCAT = vbNullString
        gap = False

        If first Then
            vF1 = Trim$(iCS.GetText(8, 12, 8))
            vT1 = Trim$(iCS.GetText(8, 21, 8))
            vF2 = Trim$(iCS.GetText(12, 12, 8))
            vT2 = Trim$(iCS.GetText(12, 21, 8))
            vF3 = Trim$(iCS.GetText(16, 12, 8))
            vT3 = Trim$(iCS.GetText(16, 21, 8))

            dF1 = ParseYYYYMMDD(CStr(vF1)): dT1 = ParseYYYYMMDD(CStr(vT1))
            dF2 = ParseYYYYMMDD(CStr(vF2)): dT2 = ParseYYYYMMDD(CStr(vT2))
            dF3 = ParseYYYYMMDD(CStr(vF3)): dT3 = ParseYYYYMMDD(CStr(vT3))

            ' OCT FITREP (example rule): mark col 27 "Y" if most recent To-date is in Oct of current year
            If dT1 > 0 Then
                Dim dtMostRecent As Date
                dtMostRecent = CDate(dT1)
                If Year(dtMostRecent) = Year(Date) And Month(dtMostRecent) = 10 Then
                    wsSB.Cells(i + 1, 27).Value = "Y"
                Else
                    wsSB.Cells(i + 1, 27).Value = "N"
                End If
                wsSB.Cells(i + 1, 26).Value = Format$(dtMostRecent, "DMMMYY")
            Else
                wsSB.Cells(i + 1, 27).Value = "N"
                wsSB.Cells(i + 1, 26).ClearContents
            End If
            first = False

            If dF1 - dT2 > 30 Then gap = True: IssueCAT = "FITREP Gap > 30 days: (" & CInt(dF1 - dT2) & " days)": nIssue = vT2 & " to " & vF1: writeRB
            If dF2 - dT3 > 30 And dT3 > 0 Then gap = True: IssueCAT = "FITREP Gap > 30 days: (" & CInt(dF2 - dT3) & " days)": nIssue = vT3 & " to " & vF2: writeRB

        Else
            vF4 = vF3: vT4 = vT3
            vF1 = Trim$(iCS.GetText(8, 12, 8))
            vT1 = Trim$(iCS.GetText(8, 21, 8))
            vF2 = Trim$(iCS.GetText(12, 12, 8))
            vT2 = Trim$(iCS.GetText(12, 21, 8))
            vF3 = Trim$(iCS.GetText(16, 12, 8))
            vT3 = Trim$(iCS.GetText(16, 21, 8))

            dF1 = ParseYYYYMMDD(CStr(vF1)): dT1 = ParseYYYYMMDD(CStr(vT1))
            dF2 = ParseYYYYMMDD(CStr(vF2)): dT2 = ParseYYYYMMDD(CStr(vT2))
            dF3 = ParseYYYYMMDD(CStr(vF3)): dT3 = ParseYYYYMMDD(CStr(vT3))
            dF4 = ParseYYYYMMDD(CStr(vF4)): dT4 = ParseYYYYMMDD(CStr(vT4))

            If dF4 - dT1 > 30 Then gap = True: IssueCAT = "FITREP Gap > 30 days (" & CInt(dF4 - dT1) & " days)": nIssue = vT1 & " to " & vF4: writeRB
            If dF1 - dT2 > 30 Then gap = True: IssueCAT = "FITREP Gap > 30 days (" & CInt(dF1 - dT2) & " days)": nIssue = vT2 & " to " & vF1: writeRB
            If dF2 - dT3 > 30 And dT3 > 0 Then gap = True: IssueCAT = "FITREP Gap > 30 days (" & CInt(dF2 - dT3) & " days)": nIssue = vT3 & " to " & vF2: writeRB
        End If

        wsSB.Cells(i + 1, 28).Value = IIf(gap, "Y", "N")

        ' Look for "8=FORWard" in footer line 23 to advance pages
        Dim checkStr As String
        checkStr = Trim$(iCS.GetText(23, 1, 79))
        If InStr(1, checkStr, "8=FORWard", vbTextCompare) = 0 Then Exit Do

        HitF8
        DoEvents
    Loop

    Application.ScreenUpdating = True
End Sub

' Expands a two-digit year to four digits using a 2000-2029 pivot, otherwise 1900s.
Private Function ExpandTwoDigitYear(ByVal twoDigitYear As Integer) As Integer
    Dim normalized As Integer

    normalized = twoDigitYear Mod 100
    If normalized <= 29 Then
        ExpandTwoDigitYear = 2000 + normalized
    Else
        ExpandTwoDigitYear = 1900 + normalized
    End If
End Function

' Parses "YYYYMMDD" (preferred) or "YYMMDD" into a VBA date serial (Double).
' Returns 0 if the input is blank/invalid.
Public Function ParseYYYYMMDD(ByVal s As String) As Double
    Dim Y As Long, m As Long, d As Long
    s = Trim$(s)
    If s = vbNullString Then
        ParseYYYYMMDD = 0#
        Exit Function
    End If

    If Len(s) = 8 And IsNumeric(s) Then
        ' YYYYMMDD
        Y = CLng(VBA.Left$(s, 4))
        m = CLng(Mid$(s, 5, 2))
        d = CLng(Right$(s, 2))
    ElseIf Len(s) = 6 And IsNumeric(s) Then
        ' YYMMDD  -> pivot: 00?29 => 2000?2029, else 1900?1999
        Y = CLng(VBA.Left$(s, 2))
        If Y <= 29 Then
            Y = 2000 + Y
        Else
            Y = 1900 + Y
        End If
        m = CLng(Mid$(s, 3, 2))
        d = CLng(Right$(s, 2))
    Else
        ParseYYYYMMDD = 0#
        Exit Function
    End If

    On Error GoTo BadDate
    ParseYYYYMMDD = CDbl(DateSerial(Y, m, d))
    Exit Function

BadDate:
    ' Out-of-range (e.g., 20251340) -> treat as missing
    ParseYYYYMMDD = 0#
End Function

'========================================================================================
' RED BOARD LOGGING HELPERS
'========================================================================================
Private Sub writeRB()
    ' Appends or creates a RED BOARD issue line for the current member.
    ' Table13 layout assumed:
    '   Col1 = Name
    '   Col2 = (your lk/sequence code)
    '   Col3 = Issues (multiple lines, each prefixed with "_<Category>: <detail>")
    Dim rw As Long
    Dim nm As String
    Dim cT As Long
    
    On Error Resume Next
    
    nm = CStr(arrayID(i, 2))
    rw = FindRow(nm)

    ' If name not found (rw = 0), allocate a new row
    If rw = 0 Then
        rw = Application.WorksheetFunction.CountA(wsRB.Range("A:A")) + 1 ' Allocate a new row if name is not found
    End If

    'ToggleThisWorkbookVisibility
    
    ' If rw = 0 after trying to find a new row, then append a new row.
    If rw = 0 Then
        Dim newRow As ListRow
        Set newRow = wsRB.ListObjects("RED_Board").ListRows.Add
        rw = newRow.Range.row
    End If

    ' If first issue, seed col 1/2 and start line; else append with newline
    Dim issueLine As String
    issueLine = "_" & IssueCAT & ": " & nIssue

    cT = CountUnderscores(CStr(wsRB.Cells(rw, 3).Value))
    If cT = 0 Then
        wsRB.Cells(rw, 1).Value = nm
        wsRB.Cells(rw, 2).Value = wsSB.Cells(i + 1, 2).Value
        wsRB.Cells(rw, 3).Value = issueLine
    Else
        issues = CStr(wsRB.Cells(rw, 3).Value)
        wsRB.Cells(rw, 3).Value = issues & vbNewLine & issueLine
    End If

    If Not currentIssues Is Nothing Then
        currentIssues.Add issueLine
    End If
End Sub

Private Function BuildFinishedStatus(ByVal displayName As String) As String
    Dim statusText As String
    Dim issueSummary As String

    statusText = "Finished: " & displayName

    issueSummary = BuildIssueSummary(displayName)
    If Len(issueSummary) > 0 Then
        statusText = statusText & "  --  " & issueSummary
    End If

    BuildFinishedStatus = statusText
End Function

Private Function BuildIssueSummary(ByVal displayName As String) As String
    Dim summary As String
    Dim idx As Long
    Dim issueText As String

    If currentIssues Is Nothing Then Exit Function
    If currentIssues.Count = 0 Then Exit Function

    For idx = 1 To currentIssues.Count
        issueText = CStr(currentIssues(idx))
        If Len(summary) = 0 Then
            summary = displayName & ": " & issueText
        Else
            summary = summary & "; " & issueText
        End If
    Next idx

    BuildIssueSummary = summary
End Function

Private Function CountUnderscores(inputText As String) As Long
    ' Returns the number of "_" characters in the given string.
    Dim lengthAfterRemove As Long
    lengthAfterRemove = Len(Replace(inputText, "_", ""))
    CountUnderscores = Len(inputText) - lengthAfterRemove
End Function

Private Function FindRow(nmSearch As String) As Long
    ' Finds (or allocates) the worksheet row in RED BOARD "Table13" for the given NAME.
    ' Returns the row number to write into.
    Dim lo As ListObject
    Dim nameCol As Range
    Dim nameCell As Range
    Dim newRow As ListRow
    
    On Error Resume Next
    
    FindRow = FoundCell(wsRB, "RED_Board", nmSearch)
    If FindRow <> 0 Then Exit Function

    Set lo = wsRB.ListObjects("RED_Table")

    ' If the table already has rows, look for the first blank name slot to reuse.
    If Not lo.DataBodyRange Is Nothing Then
        Set nameCol = lo.ListColumns(1).DataBodyRange
        For Each nameCell In nameCol
            If Len(Trim$(CStr(nameCell.Value))) = 0 Then
                FindRow = nameCell.row
                Exit Function
            End If
        Next nameCell
    End If

    ' No reusable row found; append a new row to the table and return its worksheet row.
    Set newRow = lo.ListRows.Add
    FindRow = newRow.Range.row
End Function

Private Function FoundCell(ws As Worksheet, TableName As String, _
                           target As String, Optional ByRef tableIndexOut As Long = 0) As Long
    ' Looks for target in the FIRST column of the specified ListObject.
    ' Returns the WORKSHEET ROW of the found cell, or 0 if not found.
    Dim lo As ListObject
    Dim col As Range
    Dim pos As Variant
    
    On Error Resume Next
    
    Set lo = ws.ListObjects(TableName)

    If lo.DataBodyRange Is Nothing Then
        FoundCell = 0
        Exit Function
    End If

    Set col = lo.ListColumns(1).DataBodyRange ' Column 1 in the table is "Name"
    pos = Application.Match(target, col, 0)

    If IsError(pos) Then
        FoundCell = 0
    Else
        tableIndexOut = CLng(pos)                           ' 1-based index within table
        FoundCell = col.Rows(pos).row                       ' actual worksheet row number
    End If
End Function





