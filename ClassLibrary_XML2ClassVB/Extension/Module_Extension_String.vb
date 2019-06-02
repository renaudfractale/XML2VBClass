Imports System.Runtime.CompilerServices

Public Module Module_Extension_String
    <Extension>
    Public Function GString(txt As String) As String
        Return Chr(34) + txt + Chr(34)
    End Function
    <Extension>
    Public Function Normalisation(txt As String) As String
        Return txt.ToLowerInvariant.Trim({" "c, "/"c, "-"c, "\"c, ":"c, ";"c})
    End Function
End Module
