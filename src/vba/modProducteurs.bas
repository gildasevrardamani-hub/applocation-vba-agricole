Attribute VB_Name = "modProducteurs"
Option Explicit

Private Const MODULE_PRODUCTEURS As String = "Producteurs"
Private Const TABLE_PRODUCTEURS As String = "tbl_PRODUCTEURS"
Private Const TABLE_EXPLOITATIONS As String = "tbl_EXPLOITATIONS"
Private Const COL_PRODUCTEUR_ID As String = "ProducteurID"

' Point d'entree appele depuis Excel pour ouvrir le module Producteurs.
Public Sub OuvrirGestionProducteurs()
    On Error GoTo GestionErreur
    frmProducteur.Show
    Exit Sub
GestionErreur:
    GererErreur MODULE_PRODUCTEURS, "OuvrirGestionProducteurs", Err.Number, Err.Description
End Sub

' Cree un producteur dans le tableau structure PRODUCTEURS apres validation minimale.
Public Function AjouterProducteur(ByVal organisationID As String, ByVal codeProducteur As String, ByVal nomProducteur As String, ByVal contact As String, ByVal localisation As String, ByVal typeProducteur As String) As String
    On Error GoTo GestionErreur
    Dim tbl As ListObject
    Dim ligne As ListRow
    Dim nouvelID As String
    Dim nouvelEtat As String

    If Not ValiderProducteur(nomProducteur, localisation, typeProducteur) Then Exit Function

    Set tbl = TableProducteurs()
    nouvelID = NouvelIdentifiantUnique("PROD", TABLE_PRODUCTEURS, COL_PRODUCTEUR_ID)
    Set ligne = LigneDisponible(tbl, COL_PRODUCTEUR_ID)

    EcrireLigneProducteur ligne, nouvelID, organisationID, codeProducteur, nomProducteur, contact, localisation, typeProducteur, True, Now
    nouvelEtat = EtatProducteur(ligne)
    AjouterProducteur = nouvelID
    Journaliser MODULE_PRODUCTEURS, "Creation", "Producteur cree : " & nouvelID & " - " & nomProducteur, "INFO", "Producteur", nouvelID, vbNullString, nouvelEtat
    Exit Function
GestionErreur:
    GererErreur MODULE_PRODUCTEURS, "AjouterProducteur", Err.Number, Err.Description
End Function

' Met a jour un producteur existant sans modifier son identifiant ni sa date de creation.
Public Function ModifierProducteur(ByVal producteurID As String, ByVal organisationID As String, ByVal codeProducteur As String, ByVal nomProducteur As String, ByVal contact As String, ByVal localisation As String, ByVal typeProducteur As String, ByVal actif As Boolean) As Boolean
    On Error GoTo GestionErreur
    Dim ligne As ListRow
    Dim ancienEtat As String
    Dim nouvelEtat As String

    If Not EstTexteRenseigne(producteurID) Then
        SignalerErreur MODULE_PRODUCTEURS, "Aucun producteur selectionne pour modification.", "ModifierProducteur"
        Exit Function
    End If
    If Not ValiderProducteur(nomProducteur, localisation, typeProducteur) Then Exit Function

    Set ligne = LigneProducteur(producteurID)
    If ligne Is Nothing Then
        SignalerErreur MODULE_PRODUCTEURS, "Producteur introuvable : " & producteurID, "ModifierProducteur"
        Exit Function
    End If

    ancienEtat = EtatProducteur(ligne)
    EcrireLigneProducteur ligne, producteurID, organisationID, codeProducteur, nomProducteur, contact, localisation, typeProducteur, actif, ValeurLigne(ligne, "DateCreation")
    nouvelEtat = EtatProducteur(ligne)
    ModifierProducteur = True
    Journaliser MODULE_PRODUCTEURS, "Modification", "Producteur modifie : " & producteurID & " - " & nomProducteur, "INFO", "Producteur", producteurID, ancienEtat, nouvelEtat
    Exit Function
GestionErreur:
    GererErreur MODULE_PRODUCTEURS, "ModifierProducteur", Err.Number, Err.Description
End Function

' Supprime un producteur uniquement s'il ne possede aucune exploitation rattachee.
Public Function SupprimerProducteur(ByVal producteurID As String) As Boolean
    On Error GoTo GestionErreur
    Dim ligne As ListRow
    Dim ancienEtat As String

    If Not EstTexteRenseigne(producteurID) Then
        SignalerErreur MODULE_PRODUCTEURS, "Aucun producteur selectionne pour suppression.", "SupprimerProducteur"
        Exit Function
    End If

    If ProducteurADependances(producteurID) Then
        SignalerErreur MODULE_PRODUCTEURS, "Suppression impossible : ce producteur possede au moins une exploitation rattachee.", "SupprimerProducteur"
        Exit Function
    End If

    Set ligne = LigneProducteur(producteurID)
    If ligne Is Nothing Then
        SignalerErreur MODULE_PRODUCTEURS, "Producteur introuvable : " & producteurID, "SupprimerProducteur"
        Exit Function
    End If

    ancienEtat = EtatProducteur(ligne)
    ligne.Delete
    SupprimerProducteur = True
    Journaliser MODULE_PRODUCTEURS, "Suppression", "Producteur supprime : " & producteurID, "INFO", "Producteur", producteurID, ancienEtat, vbNullString
    Exit Function
GestionErreur:
    GererErreur MODULE_PRODUCTEURS, "SupprimerProducteur", Err.Number, Err.Description
End Function

Public Function ProducteurADependances(ByVal producteurID As String) As Boolean
    On Error GoTo GestionErreur
    Dim tbl As ListObject

    Set tbl = ObtenirTableau(TABLE_EXPLOITATIONS)
    ProducteurADependances = IdentifiantExiste(tbl, COL_PRODUCTEUR_ID, producteurID)
    Exit Function
GestionErreur:
    GererErreur MODULE_PRODUCTEURS, "ProducteurADependances", Err.Number, Err.Description
End Function

' Alimente la liste du formulaire avec un filtre simple sur nom, code ou contact.
Public Sub ChargerProducteursDansListe(ByVal liste As Object, Optional ByVal filtre As String = "")
    On Error GoTo GestionErreur
    Dim tbl As ListObject
    Dim ligne As ListRow
    Dim texteRecherche As String
    Dim producteurID As String
    Dim nomProducteur As String
    Dim codeProducteur As String
    Dim contact As String

    Set tbl = TableProducteurs()
    texteRecherche = UCase$(Trim$(filtre))
    liste.Clear
    liste.ColumnCount = 5
    liste.ColumnWidths = "100 pt;100 pt;160 pt;100 pt;120 pt"

    For Each ligne In tbl.ListRows
        producteurID = CStr(ValeurLigne(ligne, COL_PRODUCTEUR_ID))
        nomProducteur = CStr(ValeurLigne(ligne, "NomProducteur"))
        codeProducteur = CStr(ValeurLigne(ligne, "CodeProducteur"))
        contact = CStr(ValeurLigne(ligne, "Contact"))

        If EstTexteRenseigne(producteurID) And _
           (texteRecherche = vbNullString Or InStr(1, UCase$(nomProducteur & " " & codeProducteur & " " & contact), texteRecherche, vbTextCompare) > 0) Then
            liste.AddItem producteurID
            liste.List(liste.ListCount - 1, 1) = codeProducteur
            liste.List(liste.ListCount - 1, 2) = nomProducteur
            liste.List(liste.ListCount - 1, 3) = contact
            liste.List(liste.ListCount - 1, 4) = CStr(ValeurLigne(ligne, "Localisation"))
        End If
    Next ligne
    Exit Sub
GestionErreur:
    GererErreur MODULE_PRODUCTEURS, "ChargerProducteursDansListe", Err.Number, Err.Description
End Sub

Public Sub ChargerProducteurDansFormulaire(ByVal producteurID As String, ByVal formulaire As Object)
    On Error GoTo GestionErreur
    Dim ligne As ListRow
    Set ligne = LigneProducteur(producteurID)
    If ligne Is Nothing Then Exit Sub

    formulaire.txtProducteurID.Value = CStr(ValeurLigne(ligne, COL_PRODUCTEUR_ID))
    formulaire.txtOrganisationID.Value = CStr(ValeurLigne(ligne, "OrganisationID"))
    formulaire.txtCodeProducteur.Value = CStr(ValeurLigne(ligne, "CodeProducteur"))
    formulaire.txtNomProducteur.Value = CStr(ValeurLigne(ligne, "NomProducteur"))
    formulaire.txtContact.Value = CStr(ValeurLigne(ligne, "Contact"))
    formulaire.txtLocalisation.Value = CStr(ValeurLigne(ligne, "Localisation"))
    formulaire.cboTypeProducteur.Value = CStr(ValeurLigne(ligne, "TypeProducteur"))
    formulaire.chkActif.Value = CBool(ValeurLigne(ligne, "Actif"))
    Exit Sub
GestionErreur:
    GererErreur MODULE_PRODUCTEURS, "ChargerProducteurDansFormulaire", Err.Number, Err.Description
End Sub

Private Function ValiderProducteur(ByVal nomProducteur As String, ByVal localisation As String, ByVal typeProducteur As String) As Boolean
    If Not EstTexteRenseigne(nomProducteur) Then
        SignalerErreur MODULE_PRODUCTEURS, "Le nom du producteur est obligatoire.", "ValiderProducteur"
        Exit Function
    End If
    If Not EstTexteRenseigne(localisation) Then
        SignalerErreur MODULE_PRODUCTEURS, "La localisation est obligatoire.", "ValiderProducteur"
        Exit Function
    End If
    If Not EstTexteRenseigne(typeProducteur) Then
        SignalerErreur MODULE_PRODUCTEURS, "Le type de producteur est obligatoire.", "ValiderProducteur"
        Exit Function
    End If
    ValiderProducteur = True
End Function

Private Sub EcrireLigneProducteur(ByVal ligne As ListRow, ByVal producteurID As String, ByVal organisationID As String, ByVal codeProducteur As String, ByVal nomProducteur As String, ByVal contact As String, ByVal localisation As String, ByVal typeProducteur As String, ByVal actif As Boolean, ByVal dateCreation As Variant)
    EcrireValeurLigne ligne, COL_PRODUCTEUR_ID, producteurID
    EcrireValeurLigne ligne, "OrganisationID", organisationID
    EcrireValeurLigne ligne, "CodeProducteur", codeProducteur
    EcrireValeurLigne ligne, "NomProducteur", nomProducteur
    EcrireValeurLigne ligne, "Contact", contact
    EcrireValeurLigne ligne, "Localisation", localisation
    EcrireValeurLigne ligne, "TypeProducteur", typeProducteur
    EcrireValeurLigne ligne, "Actif", actif
    EcrireValeurLigne ligne, "DateCreation", dateCreation
End Sub

Private Function EtatProducteur(ByVal ligne As ListRow) As String
    EtatProducteur = "ProducteurID=" & CStr(ValeurLigne(ligne, COL_PRODUCTEUR_ID)) & _
        ";NomProducteur=" & CStr(ValeurLigne(ligne, "NomProducteur")) & _
        ";Contact=" & CStr(ValeurLigne(ligne, "Contact")) & _
        ";Localisation=" & CStr(ValeurLigne(ligne, "Localisation")) & _
        ";TypeProducteur=" & CStr(ValeurLigne(ligne, "TypeProducteur")) & _
        ";Actif=" & CStr(ValeurLigne(ligne, "Actif"))
End Function

Private Function LigneProducteur(ByVal producteurID As String) As ListRow
    Set LigneProducteur = LigneParIdentifiant(TableProducteurs(), COL_PRODUCTEUR_ID, producteurID)
End Function

Private Function TableProducteurs() As ListObject
    Set TableProducteurs = ObtenirTableau(TABLE_PRODUCTEURS)
End Function
