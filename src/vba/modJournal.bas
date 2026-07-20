Attribute VB_Name = "modJournal"
Option Explicit

Private Const TABLE_JOURNAL As String = "tbl_JOURNAL"

Public Sub InitialiserJournal()
    If ObtenirTableau(TABLE_JOURNAL).ListColumns.Count = 0 Then Exit Sub
End Sub

Public Sub Journaliser(ByVal moduleSource As String, ByVal action As String, ByVal message As String, ByVal niveau As String, Optional ByVal entite As String = "", Optional ByVal identifiantEntite As String = "", Optional ByVal ancienEtat As String = "", Optional ByVal nouvelEtat As String = "")
    Dim tbl As ListObject
    Dim ligne As ListRow

    Set tbl = ObtenirTableau(TABLE_JOURNAL)
    Set ligne = LigneDisponible(tbl, "JournalID")

    EcrireValeurLigne ligne, "JournalID", NouvelIdentifiantUnique("LOG", TABLE_JOURNAL, "JournalID")
    EcrireValeurLigne ligne, "DateAction", Now
    EcrireValeurLigne ligne, "Utilisateur", Environ$("Username")
    EcrireValeurLigne ligne, "Module", moduleSource
    EcrireValeurLigne ligne, "Action", action
    EcrireValeurLigne ligne, "Entite", entite
    EcrireValeurLigne ligne, "IdentifiantEntite", identifiantEntite
    EcrireValeurLigne ligne, "Message", message
    EcrireValeurLigne ligne, "Niveau", niveau
    EcrireValeurSiColonneExiste ligne, "AncienEtat", ancienEtat
    EcrireValeurSiColonneExiste ligne, "NouvelEtat", nouvelEtat
    EcrireValeurSiColonneExiste ligne, "DateHeure", Now
End Sub

Private Sub EcrireValeurSiColonneExiste(ByVal ligne As ListRow, ByVal nomColonne As String, ByVal valeur As Variant)
    On Error Resume Next
    EcrireValeurLigne ligne, nomColonne, valeur
    On Error GoTo 0
End Sub
