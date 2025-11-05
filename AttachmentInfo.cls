Attribute VB_Name = "AttachmentInfo"
Option Explicit

Private mFileName As String
Private mFilePath As String
Private mDisplayName As String

Public Property Get FileName() As String
    FileName = mFileName
End Property

Public Property Let FileName(ByVal value As String)
    mFileName = value
End Property

Public Property Get FilePath() As String
    FilePath = mFilePath
End Property

Public Property Let FilePath(ByVal value As String)
    mFilePath = value
End Property

Public Property Get DisplayName() As String
    DisplayName = mDisplayName
End Property

Public Property Let DisplayName(ByVal value As String)
    mDisplayName = value
End Property
