Attribute VB_Name = "modTables"
Option Explicit

' Service transversal d'acces aux tableaux structures Excel.
Public Function ObtenirTableau(ByVal nomTableau As String, Optional ByVal nomFeuille As String = "") As ListObject
    Dim ws As Worksheet
    Dim tbl As ListObject

    If EstTexteRenseigne(nomFeuille) Then
        Set ObtenirTableau = ThisWorkbook.Worksheets(nomFeuille).ListObjects(nomTableau)
        Exit Function
    End If

    For Each ws In ThisWorkbook.Worksheets
        For Each tbl In ws.ListObjects
            If UCase$(tbl.Name) = UCase$(nomTableau) Then
                Set ObtenirTableau = tbl
                Exit Function
            End If
        Next tbl
    Next ws

    Err.Raise vbObjectError + 1000, "modTables.ObtenirTableau", "Tableau introuvable : " & nomTableau
End Function

Public Function ColonneTable(ByVal tbl As ListObject, ByVal nomColonne As String) As Long
    ColonneTable = tbl.ListColumns(nomColonne).Index
End Function

Public Function LigneDisponible(ByVal tbl As ListObject, ByVal nomColonneID As String) As ListRow
    Dim ligne As ListRow
    Dim colID As Long

    colID = ColonneTable(tbl, nomColonneID)
    For Each ligne In tbl.ListRows
        If Not EstTexteRenseigne(ligne.Range.Cells(1, colID).Value) Then
            Set LigneDisponible = ligne
            Exit Function
        End If
    Next ligne

    Set LigneDisponible = tbl.ListRows.Add
End Function

Public Function LigneParIdentifiant(ByVal tbl As ListObject, ByVal nomColonneID As String, ByVal identifiant As String) As ListRow
    Dim ligne As ListRow
    Dim colID As Long

    If Not EstTexteRenseigne(identifiant) Then Exit Function

    colID = ColonneTable(tbl, nomColonneID)
    For Each ligne In tbl.ListRows
        If CStr(ligne.Range.Cells(1, colID).Value) = identifiant Then
            Set LigneParIdentifiant = ligne
            Exit Function
        End If
    Next ligne
End Function

Public Function ValeurLigne(ByVal ligne As ListRow, ByVal nomColonne As String) As Variant
    ValeurLigne = ligne.Range.Cells(1, ColonneTable(ligne.Parent, nomColonne)).Value
End Function

Public Sub EcrireValeurLigne(ByVal ligne As ListRow, ByVal nomColonne As String, ByVal valeur As Variant)
    ligne.Range.Cells(1, ColonneTable(ligne.Parent, nomColonne)).Value = valeur
End Sub

Public Function IdentifiantExiste(ByVal tbl As ListObject, ByVal nomColonneID As String, ByVal identifiant As String) As Boolean
    Dim ligne As ListRow
    Set ligne = LigneParIdentifiant(tbl, nomColonneID, identifiant)
    IdentifiantExiste = Not (ligne Is Nothing)
End Function
