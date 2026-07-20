Attribute VB_Name = "modApp"
Option Explicit

Public Const APP_NAME As String = "AGRIECO PRO"
Public Const APP_VERSION As String = "V1-Sprint3"

Public Sub DemarrerApplication()
    InitialiserJournal
    Journaliser "Application", "Demarrage", "Ouverture de " & APP_NAME & " " & APP_VERSION, "INFO"
    AllerAccueil
End Sub

Public Sub InitialiserApplication()
    DemarrerApplication
End Sub
