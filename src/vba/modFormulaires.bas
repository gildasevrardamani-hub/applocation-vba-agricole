Attribute VB_Name = "modFormulaires"
Option Explicit

Public Const MODE_FORM_NOUVEAU As String = "NOUVEAU"
Public Const MODE_FORM_SELECTION As String = "SELECTION"

' Applique un comportement commun aux formulaires CRUD simples.
Public Sub AppliquerEtatBoutonsCrud(ByVal formulaire As Object, ByVal modeFormulaire As String)
    Dim selectionActive As Boolean
    selectionActive = (modeFormulaire = MODE_FORM_SELECTION)

    DefinirBouton formulaire, "cmdEnregistrer", True
    DefinirBouton formulaire, "cmdModifier", selectionActive
    DefinirBouton formulaire, "cmdSupprimer", selectionActive
    DefinirBouton formulaire, "cmdRechercher", True
End Sub

Private Sub DefinirBouton(ByVal formulaire As Object, ByVal nomControle As String, ByVal actif As Boolean)
    On Error Resume Next
    formulaire.Controls(nomControle).Enabled = actif
    On Error GoTo 0
End Sub
