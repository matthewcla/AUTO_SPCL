Attribute VB_Name = "modRedBoardRecords"
Option Explicit

Private Const RED_BOARD_TABLE_NAME As String = "RED_Board"

Public Function NormalizeRedBoardFieldKey(ByVal fieldName As String) As String
    Dim cleaned As String

    If IsNull(fieldName) Then Exit Function

    cleaned = Trim$(CStr(fieldName))
    If LenB(cleaned) = 0 Then Exit Function

    cleaned = UCase$(cleaned)
    cleaned = Replace$(cleaned, " ", vbNullString)
    cleaned = Replace$(cleaned, "_", vbNullString)
    cleaned = Replace$(cleaned, "-", vbNullString)
    cleaned = Replace$(cleaned, ".", vbNullString)
    cleaned = Replace$(cleaned, ":", vbNullString)
    cleaned = Replace$(cleaned, "/", vbNullString)
    cleaned = Replace$(cleaned, "\", vbNullString)
    cleaned = Replace$(cleaned, "(", vbNullString)
    cleaned = Replace$(cleaned, ")", vbNullString)

    NormalizeRedBoardFieldKey = cleaned
End Function

Private Function TryGetRedBoardTable() As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set lo = ws.ListObjects(RED_BOARD_TABLE_NAME)
        On Error GoTo 0

        If Not lo Is Nothing Then
            Set TryGetRedBoardTable = lo
            Exit Function
        End If
    Next ws
End Function

Public Function GetRedBoardCount() As Long
    Dim lo As ListObject

    On Error GoTo CleanFail

    Set lo = TryGetRedBoardTable()
    If lo Is Nothing Then GoTo CleanExit
    If lo.DataBodyRange Is Nothing Then GoTo CleanExit

    GetRedBoardCount = lo.DataBodyRange.Rows.Count

CleanExit:
    Exit Function

CleanFail:
    Err.Clear
    Resume CleanExit
End Function

Public Function GetRedBoardRecord(ByVal recordIndex As Long) As Object
    Dim lo As ListObject
    Dim rowRange As Range
    Dim dict As Object
    Dim column As ListColumn
    Dim value As Variant
    Dim columnKey As String
    Dim normalizedKey As String

    If recordIndex < 1 Then Exit Function

    Set lo = TryGetRedBoardTable()
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If recordIndex > lo.DataBodyRange.Rows.Count Then Exit Function

    On Error GoTo CleanFail
    Set rowRange = lo.DataBodyRange.Rows(recordIndex)
    On Error GoTo 0
    If rowRange Is Nothing Then Exit Function

    Set dict = CreateObject("Scripting.Dictionary")
    If dict Is Nothing Then Exit Function

    On Error Resume Next
    dict.CompareMode = vbTextCompare
    On Error GoTo CleanFail

    For Each column In lo.ListColumns
        value = rowRange.Columns(column.Index).Value
        columnKey = Trim$(CStr(column.Name))
        If LenB(columnKey) > 0 Then
            dict(columnKey) = value
            normalizedKey = NormalizeRedBoardFieldKey(columnKey)
            If LenB(normalizedKey) > 0 Then
                dict(normalizedKey) = value
            End If
        End If
        dict(CStr(column.Index)) = value
    Next column

    dict("_RowIndex") = recordIndex
    dict("_WorksheetRow") = rowRange.Row
    Set dict("_RowRange") = rowRange

    Set GetRedBoardRecord = dict
    Exit Function

CleanFail:
    Err.Clear
End Function

