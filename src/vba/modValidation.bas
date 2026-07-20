Attribute VB_Name = "modValidation"
Option Explicit

Public Function EstTexteRenseigne(ByVal valeur As Variant) As Boolean
    EstTexteRenseigne = Len(Trim$(CStr(valeur))) > 0
End Function

Public Function EstNombrePositif(ByVal valeur As Variant, Optional ByVal strictementPositif As Boolean = False) As Boolean
    If Not IsNumeric(valeur) Then
        EstNombrePositif = False
    ElseIf strictementPositif Then
        EstNombrePositif = CDbl(valeur) > 0
    Else
        EstNombrePositif = CDbl(valeur) >= 0
    End If
End Function

Public Function EstDateValide(ByVal valeur As Variant) As Boolean
    EstDateValide = IsDate(valeur)
End Function
