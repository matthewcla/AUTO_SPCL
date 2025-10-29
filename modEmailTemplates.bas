Attribute VB_Name = "modEmailTemplates"
Option Explicit

Private Const TEMPLATE_SHEET_NAME As String = "EmailTemplates"
Private Const TEMPLATE_SHEET_NAME_ALT As String = "Email Templates"

Public Sub LoadEmailTemplateData(ByVal templateKey As String, _
                                 ByRef txtTO As MSForms.TextBox, _
                                 ByRef txtCC As MSForms.TextBox, _
                                 ByRef txtAT As MSForms.TextBox, _
                                 ByRef txtSubj As MSForms.TextBox, _
                                 ByRef txtBody As MSForms.TextBox, _
                                 ByRef txtSignature As MSForms.TextBox)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim headerRow As Range
    Dim dataRange As Range
    Dim rowIndex As Long
    Dim keyColumn As Long
    Dim toColumn As Long
    Dim ccColumn As Long
    Dim attachColumn As Long
    Dim subjColumn As Long
    Dim bodyColumn As Long
    Dim signatureColumn As Long

    If LenB(templateKey) = 0 Then Exit Sub

    Set ws = ResolveTemplateWorksheet()
    If ws Is Nothing Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then Exit Sub

    Set headerRow = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
    Set dataRange = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol))

    DetermineTemplateColumns headerRow, keyColumn, toColumn, ccColumn, _
                              attachColumn, subjColumn, bodyColumn, signatureColumn

    If keyColumn = 0 Then keyColumn = headerRow.Cells(1, 1).Column
    If toColumn = 0 Then toColumn = keyColumn + 1
    If ccColumn = 0 Then ccColumn = toColumn + 1
    If attachColumn = 0 Then attachColumn = ccColumn + 1
    If subjColumn = 0 Then subjColumn = attachColumn + 1
    If bodyColumn = 0 Then bodyColumn = subjColumn + 1
    If signatureColumn = 0 Then signatureColumn = bodyColumn + 1

    For rowIndex = 1 To dataRange.Rows.Count
        If StrComp(templateKey, Trim$(CStrSafe(ws.Cells(rowIndex + 1, keyColumn).Value)), vbTextCompare) = 0 Then
            txtTO.Value = CStrSafe(ws.Cells(rowIndex + 1, toColumn).Value)
            txtCC.Value = CStrSafe(ws.Cells(rowIndex + 1, ccColumn).Value)
            txtAT.Value = CStrSafe(ws.Cells(rowIndex + 1, attachColumn).Value)
            txtSubj.Value = CStrSafe(ws.Cells(rowIndex + 1, subjColumn).Value)
            txtBody.Value = CStrSafe(ws.Cells(rowIndex + 1, bodyColumn).Value)
            txtSignature.Value = CStrSafe(ws.Cells(rowIndex + 1, signatureColumn).Value)
            Exit Sub
        End If
    Next rowIndex
End Sub

Private Sub DetermineTemplateColumns(ByVal headerRow As Range, _
                                     ByRef keyColumn As Long, _
                                     ByRef toColumn As Long, _
                                     ByRef ccColumn As Long, _
                                     ByRef attachColumn As Long, _
                                     ByRef subjColumn As Long, _
                                     ByRef bodyColumn As Long, _
                                     ByRef signatureColumn As Long)
    Dim headerCell As Range
    Dim headerValue As String

    For Each headerCell In headerRow.Cells
        headerValue = LCase$(Trim$(CStrSafe(headerCell.Value)))
        Select Case headerValue
            Case "template key", "template", "key", "template id", "id"
                If keyColumn = 0 Then keyColumn = headerCell.Column
            Case "to", "to list", "tolist", "recipients", "recipient"
                If toColumn = 0 Then toColumn = headerCell.Column
            Case "cc", "cc list", "cclist"
                If ccColumn = 0 Then ccColumn = headerCell.Column
            Case "attachments", "attachment", "attach", "att"
                If attachColumn = 0 Then attachColumn = headerCell.Column
            Case "subject", "subj", "title"
                If subjColumn = 0 Then subjColumn = headerCell.Column
            Case "body", "message"
                If bodyColumn = 0 Then bodyColumn = headerCell.Column
            Case "signature", "signoff"
                If signatureColumn = 0 Then signatureColumn = headerCell.Column
        End Select
    Next headerCell
End Sub

Private Function ResolveTemplateWorksheet() As Worksheet
    On Error Resume Next
    Set ResolveTemplateWorksheet = ThisWorkbook.Worksheets(TEMPLATE_SHEET_NAME)
    If ResolveTemplateWorksheet Is Nothing Then
        Set ResolveTemplateWorksheet = ThisWorkbook.Worksheets(TEMPLATE_SHEET_NAME_ALT)
    End If
    On Error GoTo 0
End Function

Private Function CStrSafe(ByVal value As Variant) As String
    If IsError(value) Then Exit Function
    If IsNull(value) Then Exit Function
    If IsEmpty(value) Then Exit Function
    CStrSafe = Trim$(CStr(value))
End Function
