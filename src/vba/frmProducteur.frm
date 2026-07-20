VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin {C62A69F0-16A6-11CE-9E98-00AA00574A4F} frmProducteur
   Caption         =   "AGRIECO PRO - Gestion des producteurs"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9300
   StartUpPosition =   1  'CenterOwner
   Begin MSForms.TextBox txtRecherche
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   300
      Width           =   3000
   End
   Begin MSForms.CommandButton cmdRechercher
      Caption         =   "Rechercher"
      Height          =   330
      Left            =   3240
      TabIndex        =   1
      Top             =   285
      Width           =   1200
   End
   Begin MSForms.ListBox lstProducteurs
      Height          =   2400
      Left            =   120
      TabIndex        =   2
      Top             =   780
      Width           =   9000
   End
   Begin MSForms.TextBox txtProducteurID
      Height          =   300
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   3480
      Width           =   3000
   End
   Begin MSForms.TextBox txtOrganisationID
      Height          =   300
      Left            =   1800
      TabIndex        =   4
      Top             =   3900
      Width           =   3000
   End
   Begin MSForms.TextBox txtCodeProducteur
      Height          =   300
      Left            =   1800
      TabIndex        =   5
      Top             =   4320
      Width           =   3000
   End
   Begin MSForms.TextBox txtNomProducteur
      Height          =   300
      Left            =   1800
      TabIndex        =   6
      Top             =   4740
      Width           =   3000
   End
   Begin MSForms.TextBox txtContact
      Height          =   300
      Left            =   6120
      TabIndex        =   7
      Top             =   3480
      Width           =   3000
   End
   Begin MSForms.TextBox txtLocalisation
      Height          =   300
      Left            =   6120
      TabIndex        =   8
      Top             =   3900
      Width           =   3000
   End
   Begin MSForms.ComboBox cboTypeProducteur
      Height          =   300
      Left            =   6120
      TabIndex        =   9
      Top             =   4320
      Width           =   3000
   End
   Begin MSForms.CheckBox chkActif
      Caption         =   "Actif"
      Height          =   300
      Left            =   6120
      TabIndex        =   10
      Top             =   4740
      Width           =   1000
   End
   Begin MSForms.CommandButton cmdNouveau
      Caption         =   "Nouveau"
      Height          =   360
      Left            =   120
      TabIndex        =   11
      Top             =   6360
      Width           =   1200
   End
   Begin MSForms.CommandButton cmdEnregistrer
      Caption         =   "Enregistrer"
      Height          =   360
      Left            =   1440
      TabIndex        =   12
      Top             =   6360
      Width           =   1200
   End
   Begin MSForms.CommandButton cmdModifier
      Caption         =   "Modifier"
      Height          =   360
      Left            =   2760
      TabIndex        =   13
      Top             =   6360
      Width           =   1200
   End
   Begin MSForms.CommandButton cmdSupprimer
      Caption         =   "Supprimer"
      Height          =   360
      Left            =   4080
      TabIndex        =   14
      Top             =   6360
      Width           =   1200
   End
   Begin MSForms.CommandButton cmdFermer
      Caption         =   "Fermer"
      Height          =   360
      Left            =   7920
      TabIndex        =   15
      Top             =   6360
      Width           =   1200
   End
   Begin MSForms.Label lblProducteurID
      Caption         =   "Identifiant"
      Height          =   240
      Left            =   120
      Top             =   3540
      Width           =   1500
   End
   Begin MSForms.Label lblOrganisationID
      Caption         =   "Organisation"
      Height          =   240
      Left            =   120
      Top             =   3960
      Width           =   1500
   End
   Begin MSForms.Label lblCodeProducteur
      Caption         =   "Code"
      Height          =   240
      Left            =   120
      Top             =   4380
      Width           =   1500
   End
   Begin MSForms.Label lblNomProducteur
      Caption         =   "Nom *"
      Height          =   240
      Left            =   120
      Top             =   4800
      Width           =   1500
   End
   Begin MSForms.Label lblContact
      Caption         =   "Contact"
      Height          =   240
      Left            =   5040
      Top             =   3540
      Width           =   1000
   End
   Begin MSForms.Label lblLocalisation
      Caption         =   "Localisation *"
      Height          =   240
      Left            =   5040
      Top             =   3960
      Width           =   1200
   End
   Begin MSForms.Label lblTypeProducteur
      Caption         =   "Type *"
      Height          =   240
      Left            =   5040
      Top             =   4380
      Width           =   1000
   End
End
Attribute VB_Name = "frmProducteur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    InitialiserTypesProducteur
    ViderFormulaire
    ChargerProducteursDansListe lstProducteurs
    AppliquerEtatBoutonsCrud Me, MODE_FORM_NOUVEAU
End Sub

Private Sub cmdNouveau_Click()
    ViderFormulaire
End Sub

Private Sub cmdEnregistrer_Click()
    Dim producteurID As String
    producteurID = AjouterProducteur(txtOrganisationID.Value, txtCodeProducteur.Value, txtNomProducteur.Value, txtContact.Value, txtLocalisation.Value, cboTypeProducteur.Value)
    If Len(producteurID) > 0 Then
        txtProducteurID.Value = producteurID
        ChargerProducteursDansListe lstProducteurs, txtRecherche.Value
        AppliquerEtatBoutonsCrud Me, MODE_FORM_SELECTION
        MsgBox "Producteur enregistre avec succes.", vbInformation, APP_NAME
    End If
End Sub

Private Sub cmdModifier_Click()
    If ModifierProducteur(txtProducteurID.Value, txtOrganisationID.Value, txtCodeProducteur.Value, txtNomProducteur.Value, txtContact.Value, txtLocalisation.Value, cboTypeProducteur.Value, chkActif.Value) Then
        ChargerProducteursDansListe lstProducteurs, txtRecherche.Value
        AppliquerEtatBoutonsCrud Me, MODE_FORM_SELECTION
        MsgBox "Producteur modifie avec succes.", vbInformation, APP_NAME
    End If
End Sub

Private Sub cmdSupprimer_Click()
    If MsgBox("Confirmer la suppression du producteur selectionne ?", vbQuestion + vbYesNo, APP_NAME) = vbNo Then Exit Sub
    If SupprimerProducteur(txtProducteurID.Value) Then
        ViderFormulaire
        ChargerProducteursDansListe lstProducteurs, txtRecherche.Value
        AppliquerEtatBoutonsCrud Me, MODE_FORM_NOUVEAU
        MsgBox "Producteur supprime avec succes.", vbInformation, APP_NAME
    End If
End Sub

Private Sub cmdRechercher_Click()
    ChargerProducteursDansListe lstProducteurs, txtRecherche.Value
End Sub

Private Sub cmdFermer_Click()
    Unload Me
End Sub

Private Sub lstProducteurs_Click()
    If lstProducteurs.ListIndex >= 0 Then
        ChargerProducteurDansFormulaire CStr(lstProducteurs.List(lstProducteurs.ListIndex, 0)), Me
        AppliquerEtatBoutonsCrud Me, MODE_FORM_SELECTION
    End If
End Sub

Private Sub ViderFormulaire()
    txtProducteurID.Value = vbNullString
    txtOrganisationID.Value = vbNullString
    txtCodeProducteur.Value = vbNullString
    txtNomProducteur.Value = vbNullString
    txtContact.Value = vbNullString
    txtLocalisation.Value = vbNullString
    If cboTypeProducteur.ListCount > 0 Then cboTypeProducteur.ListIndex = 0
    chkActif.Value = True
    AppliquerEtatBoutonsCrud Me, MODE_FORM_NOUVEAU
    txtNomProducteur.SetFocus
End Sub

Private Sub InitialiserTypesProducteur()
    ChargerListeDepuisReferentiel cboTypeProducteur, "TYPE_PRODUCTEUR"
End Sub
