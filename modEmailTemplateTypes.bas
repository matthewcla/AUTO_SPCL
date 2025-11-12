Attribute VB_Name = "modEmailTemplateTypes"
Option Explicit

Public Type EmailTemplate
    TemplateName As String
    Cc As String
    Subject As String
    Body As String
    Attachments As String
End Type
