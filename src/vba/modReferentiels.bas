Attribute VB_Name = "modReferentiels"
Option Explicit

Private Const TABLE_REFERENTIELS As String = "tbl_REFERENTIELS"

' Charge une liste active du referentiel. Les valeurs de secours ne sont utilisees
' que si le classeur ne contient pas encore le referentiel attendu.
Public Sub ChargerListeDepuisReferentiel(ByVal liste As Object, ByVal domaine As String, ParamArray valeursSecours() As Variant)
    Dim tbl As ListObject
    Dim ligne As ListRow
    Dim nbAjouts As Long

    liste.Clear
    Set tbl = ObtenirTableau(TABLE_REFERENTIELS)

    For Each ligne In tbl.ListRows
        If UCase$(CStr(ValeurLigne(ligne, "Domaine"))) = UCase$(domaine) _
           And EstValeurActive(ValeurLigne(ligne, "Actif")) Then
            liste.AddItem CStr(ValeurLigne(ligne, "Libelle"))
            nbAjouts = nbAjouts + 1
        End If
    Next ligne

    If nbAjouts = 0 Then AjouterValeursSecours liste, valeursSecours
End Sub

Private Sub AjouterValeursSecours(ByVal liste As Object, ByVal valeursSecours As Variant)
    Dim index As Long

    On Error GoTo AucuneValeur
    For index = LBound(valeursSecours) To UBound(valeursSecours)
        liste.AddItem CStr(valeursSecours(index))
    Next index
AucuneValeur:
End Sub

Public Function EstValeurActive(ByVal valeur As Variant) As Boolean
    Dim texte As String
    texte = UCase$(Trim$(CStr(valeur)))
    EstValeurActive = (texte = "" Or texte = "TRUE" Or texte = "VRAI" Or texte = "1" Or texte = "OUI")
End Function
