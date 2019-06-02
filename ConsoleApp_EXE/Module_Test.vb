Imports System.IO
Imports ClassLibrary_XML2ClassVB
Module Module_Test

    Sub Main()

        Dim Path_In As String = "Path Of Rep Files XML"
        Dim Path_Out As String = Path.GetTempPath
        'Déclaration des Variables Global
        ClassLibrary_XML2ClassVB.Déclaration_XML2VB.Déclaration()

        For Each Path_FileXML In Directory.GetFiles(Path_In, "*.xml")
            ClassLibrary_XML2ClassVB.Class_XML2VB.AnalyzeFile(Path_FileXML)
        Next
        'On simplifit les structures ////// IMPORTANT \\\\\\
        ClassLibrary_XML2ClassVB.Class_XML2VB.Simplification()

        Dim TimeCode = Now.ToString("yyyy-MM-dd HH-mm-ss ffff")
        Dim PathDir = Path_Out + TimeCode + "\"
        My.Computer.FileSystem.CreateDirectory(PathDir)

        ClassLibrary_XML2ClassVB.Class_XML2VB.GenerateVBFiles(PathDir)

        Dim P As New Process()
        P.Start(PathDir)

    End Sub



End Module
