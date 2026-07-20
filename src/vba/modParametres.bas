Attribute VB_Name = "modParametres"
Option Explicit

Private Const FEUILLE_PARAMETRES As String = "PARAMETRES"

Public Function LireParametre(ByVal domaine As String, ByVal cle As String, Optional ByVal valeurDefaut As String = "") As String
    Dim tbl As ListObject
    Dim ligne As ListRow
    Set tbl = ThisWorkbook.Worksheets(FEUILLE_PARAMETRES).ListObjects(1)
    LireParametre = valeurDefaut
    For Each ligne In tbl.ListRows
        If CStr(ligne.Range.Cells(1, 2).Value) = domaine And CStr(ligne.Range.Cells(1, 3).Value) = cle Then
            LireParametre = CStr(ligne.Range.Cells(1, 4).Value)
            Exit Function
        End If
    Next ligne
End Function
