Attribute VB_Name = "modEmailPlaceholders"
Option Explicit

'-------------------------------------------------------------------------------
' Procedure: ReplacePlaceholders
' Purpose  : Substitute placeholder tokens within a template string using supplied
'            name/value pairs.
' Parameters:
'   template - Source text containing placeholder tokens such as {Name}.
'   placeholderPairs - ParamArray of alternating placeholder names and values.
' Returns  : String with placeholders replaced; returns template when no replacements apply.
' Side Effects:
'   None.
'-------------------------------------------------------------------------------
Public Function ReplacePlaceholders(ByVal template As String, ParamArray placeholderPairs()) As String
    Dim normalizedPairs As Variant
    Dim upper As Long

    On Error GoTo NoArguments
    upper = UBound(placeholderPairs)
    On Error GoTo 0

    If upper < 0 Then
        ReplacePlaceholders = template
        Exit Function
    End If

    normalizedPairs = NormaliseParamArrayPairs(placeholderPairs, upper)
    ReplacePlaceholders = ReplacePlaceholdersArray(template, normalizedPairs)
    Exit Function

NoArguments:
    On Error GoTo 0
    ReplacePlaceholders = template
End Function

Private Function NormaliseParamArrayPairs(ByRef placeholderPairs() As Variant, ByVal upper As Long) As Variant
    Dim pairCount As Long
    Dim result() As Variant
    Dim index As Long
    Dim targetRow As Long

    pairCount = (upper + 2) \/ 2

    If pairCount = 0 Then
        NormaliseParamArrayPairs = result
        Exit Function
    End If

    ReDim result(0 To pairCount - 1, 0 To 1)

    For index = 0 To upper Step 2
        targetRow = index \ 2
        result(targetRow, 0) = placeholderPairs(index)
        If index + 1 <= upper Then
            result(targetRow, 1) = placeholderPairs(index + 1)
        Else
            result(targetRow, 1) = vbNullString
        End If
    Next index

    NormaliseParamArrayPairs = result
End Function

'-------------------------------------------------------------------------------
' Procedure: ReplacePlaceholdersArray
' Purpose  : Perform placeholder substitution using a caller-provided array of values.
' Parameters:
'   template - Source text containing placeholder tokens.
'   placeholderPairs - Array (1-D or 2-D) containing alternating placeholder names
'                     and values or a dictionary keyed by placeholder name.
' Returns  : String with placeholders replaced; returns template when no replacements apply.
' Side Effects:
'   None.
'-------------------------------------------------------------------------------
Public Function ReplacePlaceholdersArray(ByVal template As String, ByRef placeholderPairs As Variant) As String
    Dim result As String

    result = template

    If LenB(template) = 0 Then
        ReplacePlaceholdersArray = template
        Exit Function
    End If

    If IsObject(placeholderPairs) Then
        result = ReplaceUsingDictionaryObject(result, placeholderPairs)
        ReplacePlaceholdersArray = result
        Exit Function
    End If

    If IsArray(placeholderPairs) Then
        result = ReplaceUsingVariantArray(result, placeholderPairs)
    End If

    ReplacePlaceholdersArray = result
End Function

Private Function ReplaceUsingVariantArray(ByVal template As String, ByRef placeholderPairs As Variant) As String
    Dim dimensionCount As Long
    Dim result As String

    result = template

    dimensionCount = ArrayDimensionCount(placeholderPairs)

    Select Case dimensionCount
        Case 0
            ' Uninitialised array; nothing to replace.
        Case 1
            result = ReplaceUsingFlatVariantArray(result, placeholderPairs)
        Case Else
            result = ReplaceUsingMatrixVariantArray(result, placeholderPairs)
    End Select

    ReplaceUsingVariantArray = result
End Function

Private Function ArrayDimensionCount(ByRef values As Variant) As Long
    Dim dimIndex As Long
    Dim currentUpper As Long

    On Error GoTo ExitRoutine

    For dimIndex = 1 To 60
        currentUpper = UBound(values, dimIndex)
        ArrayDimensionCount = dimIndex
    Next dimIndex

ExitRoutine:
    On Error GoTo 0
End Function

Private Function ReplaceUsingFlatVariantArray(ByVal template As String, ByRef placeholderPairs As Variant) As String
    Dim lower As Long
    Dim upper As Long
    Dim i As Long
    Dim result As String

    result = template

    On Error Resume Next
    lower = LBound(placeholderPairs)
    upper = UBound(placeholderPairs)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        ReplaceUsingFlatVariantArray = result
        Exit Function
    End If
    On Error GoTo 0

    If upper < lower Then
        ReplaceUsingFlatVariantArray = result
        Exit Function
    End If

    For i = lower To upper Step 2
        If i + 1 > upper Then Exit For
        result = ApplyPlaceholderReplacement(result, placeholderPairs(i), placeholderPairs(i + 1))
    Next i

    ReplaceUsingFlatVariantArray = result
End Function

Private Function ReplaceUsingMatrixVariantArray(ByVal template As String, ByRef placeholderPairs As Variant) As String
    Dim rowLower As Long
    Dim rowUpper As Long
    Dim colLower As Long
    Dim colUpper As Long
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim result As String

    result = template

    On Error Resume Next
    rowLower = LBound(placeholderPairs, 1)
    rowUpper = UBound(placeholderPairs, 1)
    colLower = LBound(placeholderPairs, 2)
    colUpper = UBound(placeholderPairs, 2)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        ReplaceUsingMatrixVariantArray = result
        Exit Function
    End If
    On Error GoTo 0

    If rowUpper < rowLower Or colUpper < colLower Then
        ReplaceUsingMatrixVariantArray = result
        Exit Function
    End If

    For rowIndex = rowLower To rowUpper
        For columnIndex = colLower To colUpper Step 2
            If columnIndex + 1 > colUpper Then Exit For
            result = ApplyPlaceholderReplacement(result, placeholderPairs(rowIndex, columnIndex), _
                                               placeholderPairs(rowIndex, columnIndex + 1))
        Next columnIndex
    Next rowIndex

    ReplaceUsingMatrixVariantArray = result
End Function

Private Function ReplaceUsingDictionaryObject(ByVal template As String, ByRef placeholderPairs As Variant) As String
    Dim result As String
    Dim keys As Variant
    Dim key As Variant

    result = template

    If placeholderPairs Is Nothing Then
        ReplaceUsingDictionaryObject = result
        Exit Function
    End If

    On Error GoTo Cleanup
    keys = placeholderPairs.keys

    If IsArray(keys) Then
        For Each key In keys
            result = ApplyPlaceholderReplacement(result, key, placeholderPairs(key))
        Next key
    End If

Cleanup:
    On Error GoTo 0
    ReplaceUsingDictionaryObject = result
End Function

Private Function ApplyPlaceholderReplacement(ByVal template As String, _
                                             ByVal placeholderNameValue As Variant, _
                                             ByVal placeholderRawValue As Variant) As String
    Dim placeholderName As String
    Dim placeholderValue As String
    Dim result As String

    result = template

    placeholderName = Trim$(SafePlaceholderText(placeholderNameValue))
    If LenB(placeholderName) = 0 Then
        ApplyPlaceholderReplacement = result
        Exit Function
    End If

    placeholderValue = SafePlaceholderText(placeholderRawValue)
    result = Replace(result, "{" & placeholderName & "}", placeholderValue, , , vbTextCompare)

    ApplyPlaceholderReplacement = result
End Function

Private Function SafePlaceholderText(ByVal value As Variant) As String
    If IsError(value) Then Exit Function
    If IsNull(value) Then Exit Function
    SafePlaceholderText = CStr(value)
End Function
