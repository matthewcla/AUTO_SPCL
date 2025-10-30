Attribute VB_Name = "modEmailPlaceholders"
Option Explicit

Public Function ReplacePlaceholders(ByVal template As String, ParamArray placeholderPairs()) As String
    ReplacePlaceholders = ReplacePlaceholdersArray(template, placeholderPairs)
End Function

Public Function ReplacePlaceholdersArray(ByVal template As String, ByRef placeholderPairs As Variant) As String
    Dim result As String
    Dim lower As Long
    Dim upper As Long
    Dim i As Long
    Dim placeholderName As String
    Dim placeholderValue As String

    result = template

    If LenB(template) = 0 Then
        ReplacePlaceholdersArray = template
        Exit Function
    End If

    If Not IsArray(placeholderPairs) Then
        ReplacePlaceholdersArray = result
        Exit Function
    End If

    On Error Resume Next
    lower = LBound(placeholderPairs)
    upper = UBound(placeholderPairs)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        ReplacePlaceholdersArray = result
        Exit Function
    End If
    On Error GoTo 0

    If upper < lower Then
        ReplacePlaceholdersArray = result
        Exit Function
    End If

    For i = lower To upper Step 2
        If i + 1 > upper Then Exit For

        placeholderName = Trim$(SafePlaceholderText(placeholderPairs(i)))
        If LenB(placeholderName) = 0 Then
            GoTo NextPlaceholder
        End If

        placeholderValue = SafePlaceholderText(placeholderPairs(i + 1))
        result = Replace(result, "{" & placeholderName & "}", placeholderValue, , , vbTextCompare)
NextPlaceholder:
    Next i

    ReplacePlaceholdersArray = result
End Function

Private Function SafePlaceholderText(ByVal value As Variant) As String
    If IsError(value) Then Exit Function
    If IsNull(value) Then Exit Function
    SafePlaceholderText = CStr(value)
End Function
