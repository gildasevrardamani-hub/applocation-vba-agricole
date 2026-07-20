Attribute VB_Name = "modIdentifiants"
Option Explicit

Public Function NouvelIdentifiant(ByVal prefixe As String) As String
    Randomize
    NouvelIdentifiant = UCase$(prefixe) & "-" & Format$(Now, "yyyymmddhhmmss") & "-" & Format$(CLng(Rnd() * 1000000), "000000")
End Function

Public Function NouvelIdentifiantUnique(ByVal prefixe As String, ByVal nomTableau As String, ByVal nomColonneID As String) As String
    Dim tbl As ListObject
    Dim candidat As String
    Dim tentative As Long

    Set tbl = ObtenirTableau(nomTableau)
    For tentative = 1 To 100
        candidat = NouvelIdentifiant(prefixe)
        If Not IdentifiantExiste(tbl, nomColonneID, candidat) Then
            NouvelIdentifiantUnique = candidat
            Exit Function
        End If
        DoEvents
    Next tentative

    Err.Raise vbObjectError + 1100, "modIdentifiants.NouvelIdentifiantUnique", "Impossible de generer un identifiant unique pour " & nomTableau
End Function
