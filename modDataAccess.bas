Option Explicit

'-----------------------------------------------------------------------
' Module: modDataAccess
' Purpose: Provide guarded helpers for interacting with the ID worksheet.
'-----------------------------------------------------------------------

Public Function GetIDSheet() As Worksheet
    Dim wsID As Worksheet

    On Error Resume Next
    Set wsID = ThisWorkbook.Worksheets("ID")
    If Err.Number <> 0 Then
        Debug.Print "[modDataAccess] GetIDSheet error " & Err.Number & ": " & Err.Description
        Err.Clear
        Set wsID = Nothing
    End If
    On Error GoTo 0

    If wsID Is Nothing Then
        Debug.Print "[modDataAccess] GetIDSheet: worksheet 'ID' is unavailable."
    End If

    Set GetIDSheet = wsID
End Function

Private Function TryGetLastIdRow(ByVal wsID As Worksheet, ByRef lastRow As Long) As Boolean
    If wsID Is Nothing Then Exit Function

    On Error Resume Next
    lastRow = wsID.Cells(wsID.Rows.Count, "B").End(xlUp).Row
    If Err.Number <> 0 Then
        Debug.Print "[modDataAccess] TryGetLastIdRow: failed to read last row. Error " & _
                    Err.Number & ": " & Err.Description
        Err.Clear
        lastRow = 0
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    If lastRow < 1 Then Exit Function

    TryGetLastIdRow = True
End Function

Private Function ValidateRowIndex(ByVal wsID As Worksheet, _
                                  ByVal rowIndex As Long, _
                                  ByVal callerName As String) As Boolean
    Dim lastRow As Long

    If rowIndex < 1 Then
        Debug.Print "[modDataAccess] " & callerName & ": invalid rowIndex " & rowIndex
        Exit Function
    End If

    If Not TryGetLastIdRow(wsID, lastRow) Then Exit Function

    If rowIndex > lastRow Then
        Debug.Print "[modDataAccess] " & callerName & ": requested row " & rowIndex & _
                    " exceeds last row " & lastRow
        Exit Function
    End If

    ValidateRowIndex = True
End Function

Public Function FindIdRowByName(ByVal fullName As String) As Long
    Dim wsID As Worksheet
    Dim normalizedSearch As String
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim candidate As String

    normalizedSearch = NormalizeValue(fullName)
    If LenB(normalizedSearch) = 0 Then
        Debug.Print "[modDataAccess] FindIdRowByName: fullName argument is empty."
        Exit Function
    End If

    Set wsID = GetIDSheet()
    If wsID Is Nothing Then Exit Function

    If Not TryGetLastIdRow(wsID, lastRow) Then Exit Function

    For rowIndex = 1 To lastRow
        candidate = NormalizeValue(wsID.Cells(rowIndex, "B").Value)
        If LenB(candidate) = 0 Then GoTo NextRow

        If StrComp(candidate, normalizedSearch, vbTextCompare) = 0 Then
            FindIdRowByName = rowIndex
            Exit Function
        End If

NextRow:
    Next rowIndex

    Debug.Print "[modDataAccess] FindIdRowByName: name '" & fullName & "' not found."
End Function

Public Function GetSsnByName(ByVal fullName As String) As String
    Dim wsID As Worksheet
    Dim rowIndex As Long
    Dim value As String

    rowIndex = FindIdRowByName(fullName)
    If rowIndex = 0 Then Exit Function

    Set wsID = GetIDSheet()
    If wsID Is Nothing Then Exit Function

    On Error Resume Next
    value = NormalizeValue(wsID.Cells(rowIndex, "A").Value)
    If Err.Number <> 0 Then
        Debug.Print "[modDataAccess] GetSsnByName: failed to read column A. Error " & Err.Number & ": " & Err.Description
        Err.Clear
        value = vbNullString
    End If
    On Error GoTo 0

    GetSsnByName = value
End Function

Public Function GetSsnByRow(ByVal rowIndex As Long) As String
    Dim wsID As Worksheet
    Dim value As String

    Set wsID = GetIDSheet()
    If wsID Is Nothing Then Exit Function

    If Not ValidateRowIndex(wsID, rowIndex, "GetSsnByRow") Then Exit Function

    On Error Resume Next
    value = NormalizeValue(wsID.Cells(rowIndex, "A").Value)
    If Err.Number <> 0 Then
        Debug.Print "[modDataAccess] GetSsnByRow: failed to read column A. Error " & Err.Number & ": " & Err.Description
        Err.Clear
        value = vbNullString
    End If
    On Error GoTo 0

    GetSsnByRow = value
End Function

Public Function GetNameByRow(ByVal rowIndex As Long) As String
    Dim wsID As Worksheet
    Dim value As String

    Set wsID = GetIDSheet()
    If wsID Is Nothing Then Exit Function

    If Not ValidateRowIndex(wsID, rowIndex, "GetNameByRow") Then Exit Function

    On Error Resume Next
    value = NormalizeValue(wsID.Cells(rowIndex, "B").Value)
    If Err.Number <> 0 Then
        Debug.Print "[modDataAccess] GetNameByRow: failed to read column B. Error " & Err.Number & ": " & Err.Description
        Err.Clear
        value = vbNullString
    End If
    On Error GoTo 0

    GetNameByRow = value
End Function

Public Function GetEmailsByName(ByVal fullName As String) As String
    Dim wsID As Worksheet
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim part As String
    Dim recipients As String

    rowIndex = FindIdRowByName(fullName)
    If rowIndex = 0 Then Exit Function

    Set wsID = GetIDSheet()
    If wsID Is Nothing Then Exit Function

    ' Columns C through F hold up to four email addresses for each ID row.
    For columnIndex = 3 To 6
        On Error Resume Next
        part = NormalizeValue(wsID.Cells(rowIndex, columnIndex).Value)
        If Err.Number <> 0 Then
            Debug.Print "[modDataAccess] GetEmailsByName: failed to read column " & columnIndex & _
                        ". Error " & Err.Number & ": " & Err.Description
            Err.Clear
            part = vbNullString
        End If
        On Error GoTo 0

        If LenB(part) > 0 Then
            If LenB(recipients) > 0 Then recipients = recipients & ";"
            recipients = recipients & part
        End If
    Next columnIndex

    GetEmailsByName = recipients
End Function

Public Function GetEmailsByRow(ByVal rowIndex As Long) As String
    Dim wsID As Worksheet
    Dim columnIndex As Long
    Dim part As String
    Dim recipients As String

    Set wsID = GetIDSheet()
    If wsID Is Nothing Then Exit Function

    If Not ValidateRowIndex(wsID, rowIndex, "GetEmailsByRow") Then Exit Function

    For columnIndex = 3 To 6
        On Error Resume Next
        part = NormalizeValue(wsID.Cells(rowIndex, columnIndex).Value)
        If Err.Number <> 0 Then
            Debug.Print "[modDataAccess] GetEmailsByRow: failed to read column " & columnIndex & _
                        ". Error " & Err.Number & ": " & Err.Description
            Err.Clear
            part = vbNullString
        End If
        On Error GoTo 0

        If LenB(part) > 0 Then
            If LenB(recipients) > 0 Then recipients = recipients & ";"
            recipients = recipients & part
        End If
    Next columnIndex

    GetEmailsByRow = recipients
End Function

Private Function NormalizeValue(ByVal value As Variant) As String
    If IsError(value) Then Exit Function
    If IsNull(value) Then Exit Function
    If IsEmpty(value) Then Exit Function

    On Error Resume Next
    NormalizeValue = Trim$(CStr(value))
    If Err.Number <> 0 Then
        Err.Clear
        NormalizeValue = vbNullString
    End If
    On Error GoTo 0
End Function
