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

Private Function CStrSafe(ByVal value As Variant) As String
    If IsError(value) Then Exit Function
    If IsNull(value) Then Exit Function
    If IsEmpty(value) Then Exit Function
    CStrSafe = Trim$(CStr(value))
End Function
