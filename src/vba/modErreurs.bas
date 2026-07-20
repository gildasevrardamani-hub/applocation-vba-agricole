Attribute VB_Name = "modErreurs"
Option Explicit

Public Sub SignalerErreur(ByVal moduleSource As String, ByVal message As String, Optional ByVal procedureSource As String = "")
    Dim messageComplet As String
    messageComplet = MessageErreur(moduleSource, procedureSource, message)
    Journaliser moduleSource, "Erreur", messageComplet, "ERREUR", moduleSource, vbNullString
    MsgBox messageComplet, vbExclamation, APP_NAME
End Sub

Public Sub GererErreur(ByVal moduleSource As String, ByVal procedureSource As String, ByVal numeroErreur As Long, ByVal descriptionErreur As String)
    SignalerErreur moduleSource, "Erreur " & CStr(numeroErreur) & " : " & descriptionErreur, procedureSource
End Sub

Private Function MessageErreur(ByVal moduleSource As String, ByVal procedureSource As String, ByVal message As String) As String
    If EstTexteRenseigne(procedureSource) Then
        MessageErreur = moduleSource & "." & procedureSource & " - " & message
    Else
        MessageErreur = message
    End If
End Function
