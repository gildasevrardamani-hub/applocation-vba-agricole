Attribute VB_Name = "modNavigation"
Option Explicit

Public Sub AllerAccueil()
    ActiverFeuille "ACCUEIL"
End Sub

Public Sub OuvrirModule(ByVal nomFeuille As String)
    If FeuilleExiste(nomFeuille) Then
        ActiverFeuille nomFeuille
        Journaliser "Navigation", "OuvrirModule", nomFeuille, "INFO"
    Else
        SignalerErreur "Navigation", "Feuille introuvable : " & nomFeuille
    End If
End Sub

Private Sub ActiverFeuille(ByVal nomFeuille As String)
    ThisWorkbook.Worksheets(nomFeuille).Activate
End Sub

Private Function FeuilleExiste(ByVal nomFeuille As String) As Boolean
    Dim ws As Worksheet
    FeuilleExiste = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = nomFeuille Then
            FeuilleExiste = True
            Exit Function
        End If
    Next ws
End Function
