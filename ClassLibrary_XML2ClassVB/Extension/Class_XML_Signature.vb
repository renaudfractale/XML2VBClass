<Serializable>
Public Class Class_XML_Signature
    Property Liste_Attribue As New List(Of String)
    ' -------------------------------- Dictionary(Of LocalName, Siganture) ------Si IsFirstLevel ==TRUE
    'Property Liste_ElementIfFirstLevelTrue As New Dictionary(Of String, String)

    ' -------------------------------- Dictionary(Of LocalName, List(Of(Siganture) ------Si IsFirstLevel ==TRUE
    'Property Liste_ElementsIfFirstLevelTrue As New Dictionary(Of String, List(Of String))


    ' -------------------------------- ---------    List(Of(LocalName) ------Si IsFirstLevel ==FALSE
    Property Liste_Element As New List(Of String)

    ' -------------------------------- ---------    List(Of(LocalName) ------Si IsFirstLevel ==FALSE
    Property Liste_Elements As New List(Of String)

    Property LocalName As String
    Property ElementMain As String
    Property IsFirstLevel As Boolean = False
    Public Sub New()

    End Sub
    Public Sub New(ElementMain As XElement, Optional IsFirstLevel As Boolean = True)
        Me.ElementMain = ElementMain.ToString
        Me.IsFirstLevel = IsFirstLevel
        Me.LocalName = ElementMain.Name.LocalName

        '----------  Analyse Attribue  ---------------
        For Each Attribue As XAttribute In ElementMain.Attributes
            Liste_Attribue.Add(Attribue.Name.LocalName)
        Next
        '----------  Analyse Attribue  --------------
        '----------  Analyse Element  ---------------
        If IsFirstLevel Then
            'For Each Element As XElement In ElementMain.Elements
            '    Dim NbOcurence As Integer = GetNBOcurenceElement(ElementMain, Element)
            '    Dim Signature_Element As New Class_XML_Signature(Element, False)
            '    Dim LocalName As String = Element.Name.LocalName
            '    If NbOcurence = 1 Then
            '        'In Liste_ElementIfFirstLevelTrue
            '        Liste_ElementIfFirstLevelTrue.Add(LocalName, Signature_Element.ExportToString)
            '    Else
            '        'In Liste_ElemenstIfFirstLevelTrue
            '        If Not Liste_ElementsIfFirstLevelTrue.ContainsKey(LocalName) Then
            '            Liste_ElementsIfFirstLevelTrue.Add(LocalName, New List(Of String))
            '        End If
            '        Liste_ElementsIfFirstLevelTrue.Item(LocalName).Add(Signature_Element.ExportToString)
            '        Liste_ElementsIfFirstLevelTrue.Item(LocalName) = Liste_ElementsIfFirstLevelTrue.Item(LocalName).Distinct.ToList
            '    End If
            'Next
        Else
            For Each Element As XElement In ElementMain.Elements
                Dim NbOcurence As Integer = GetNBOcurenceElement(ElementMain, Element)
                Dim LocalName As String = Element.Name.LocalName
                If NbOcurence = 1 Then
                    Liste_Element.Add(LocalName)
                Else
                    Liste_Elements.Add(LocalName)
                    Liste_Elements = Liste_Elements.Distinct.ToList
                End If
            Next
        End If

        '----------  Analyse Element  ---------------

    End Sub
    Public Function ExportSignature() As String
        Me.Liste_Element.Sort()
        Me.Liste_Elements.Sort()
        Return Join({Join(Me.Liste_Element.ToArray, ";"), Join(Me.Liste_Elements.ToArray, ";")}, "|")
    End Function
    Private Function ExportToStringWithoutAttibues() As String
        Dim txt As String = ""
        Dim Liste_Element As List(Of String) = Me.Liste_Element
        Liste_Element.Sort()
        txt += "Element = {"
        For Each Element As String In Liste_Element
            txt += Element + " , "
        Next
        txt += "}" + vbCrLf
        ' ------------   Liste_Element ---------------

        ' ------------   Liste_Elements ---------------
        Dim Liste_Elements As List(Of String) = Me.Liste_Elements
        Liste_Elements.Sort()
        txt += "Elements = {"
        For Each Element As String In Liste_Elements
            txt += Element + " , "
        Next
        txt += "}" + vbCrLf
        Return txt
    End Function

    Public Function ExportToString() As String
        Dim txt As String = ""
        ' ------------   Liste_Attribue ---------------
        Liste_Attribue.Sort()
        txt += "Attribue = {"
        For Each Attribue As String In Liste_Attribue
            txt += Attribue + ", "
        Next
        txt += "}" + vbCrLf
        ' ------------   Liste_Attribue ---------------

        ' ------------   IsFirstLevel= True ---------------
        'If IsFirstLevel = True Then
        '    ' ------------   Liste_Element ---------------
        '    Dim Liste_Element As List(Of String) = Liste_ElementIfFirstLevelTrue.Keys.ToList
        '    Liste_Element.Sort()
        '    txt += "Element = {"
        '    For Each Element As String In Liste_Element
        '        Dim Signature As String = Liste_ElementIfFirstLevelTrue.Item(Element)
        '        txt += "[" + Element + "](" + Signature + "), "
        '    Next
        '    txt += "}" + vbCrLf
        '    ' ------------   Liste_Element ---------------


        '    ' ------------   Liste_Elements ---------------
        '    Dim Liste_Elements As List(Of String) = Liste_ElementsIfFirstLevelTrue.Keys.ToList
        '    Liste_Elements.Sort()
        '    txt += "Elements = {"
        '    For Each Element As String In Liste_Elements
        '        txt += "[" + Element + "]("
        '        Dim Signatures As List(Of String) = Liste_ElementsIfFirstLevelTrue.Item(Element)
        '        Signatures.Sort()

        '        For Each Signature As String In Signatures
        '            txt += Signature + ", "
        '        Next
        '        txt += "), "
        '    Next
        '    txt += "}" + vbCrLf
        '    ' ------------   Liste_Elements ---------------
        'End If
        ' ------------   IsFirstLevel= True ---------------

        ' ------------   IsFirstLevel= False ---------------
        If IsFirstLevel = False Then
            ' ------------   Liste_Element ---------------
            txt += ExportToStringWithoutAttibues()
            ' ------------   Liste_Elements ---------------
        End If
        ' ------------   IsFirstLevel= False ---------------
        Return txt
    End Function
End Class
