Attribute VB_Name = "modReviewState"
Option Explicit

Public LastRunCompleted As Boolean
Public LastRunWasCancelled As Boolean
Public LastRunProcessed As Long
Public LastRunTotal As Long

Public Sub ResetRunState()
    LastRunCompleted = False
    LastRunWasCancelled = False
    LastRunProcessed = 0
    LastRunTotal = 0
End Sub

