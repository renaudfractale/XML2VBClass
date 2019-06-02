Public Class Class_XML2VB
    ' -------------------------- Dictionary(Of LocalName, Dictionary(Of SignatureSytring ,SigantureObject
    Shared Property Dico_Main As New Dictionary(Of String, Dictionary(Of String, Class_XML_Signature))
    Shared Property Dico_Simple As New Dictionary(Of String, Dictionary(Of String, Class_XML_Signature))

    Shared Property ElementMaster As New Dictionary(Of String, Class_XML_Signature)

    Shared Sub AnalyzeFile(Path_FileXml As String)
        Dim Document = XDocument.Load(Path_FileXml)
        Dim Elements = Document.Elements
        Dim Element_Master = Elements.First
        Dim Signature_Element_Master = Element_Master.GetSignature
        If Not ElementMaster.ContainsKey(Signature_Element_Master.ExportToString) Then
            ElementMaster.Add(Signature_Element_Master.ExportToString, Signature_Element_Master)
        End If
        Add_SigantureObject(Signature_Element_Master)
        AnalyzeElement(Element_Master)
    End Sub

    Private Shared Sub AnalyzeElement(ElementAnalyzed As XElement)
        For Each Element As XElement In ElementAnalyzed.Elements
            Dim Signature As Class_XML_Signature = Element.GetSignature
            Add_SigantureObject(Signature)
            AnalyzeElement(Element)
        Next

    End Sub


    Private Shared Sub Add_SigantureObject(SigantureObject As Class_XML_Signature)
        Dim LocalName = SigantureObject.LocalName

        If Not Dico_Main.ContainsKey(LocalName) Then
            Dico_Main.Add(LocalName, New Dictionary(Of String, Class_XML_Signature))
        End If
        Dim Dico_Of_Element = Dico_Main.Item(LocalName)

        Dim SignatureString = SigantureObject.ExportToString
        If Not Dico_Of_Element.ContainsKey(SignatureString) Then
            Dico_Of_Element.Add(SignatureString, SigantureObject)
        End If

    End Sub


    Public Shared Sub Simplification()
        Dim Dico = Class_Cloning.Clone(Of Dictionary(Of String, Dictionary(Of String, Class_XML_Signature)))(Dico_Main)
        Dim LocalNames As List(Of String) = Dico.Keys.ToList
        LocalNames.Sort()

        For Each LocalName As String In LocalNames
            Dim Dico_LocalName = Dico.Item(LocalName)
            Dim Signatures = Dico_LocalName.Keys.ToList
            Signatures.Sort()

            Dim Liste_Signatures_ToDeleted As New List(Of Integer)

            For i As Integer = 0 To Dico_LocalName.Count - 2
                Dim SignatureRefString = Signatures(i)
                Dim SignatureRefObject = Dico_LocalName.Item(SignatureRefString)

                For j As Integer = i + 1 To Dico_LocalName.Count - 1
                    Dim SignatureString = Signatures(j)
                    'Si à suprimer on passe au suivante
                    If Liste_Signatures_ToDeleted.Contains(j) Then Continue For
                    'Si Ref suprimé on Quitte la boucle
                    If Liste_Signatures_ToDeleted.Contains(i) Then Exit For

                    Dim SignaturObject = Dico_LocalName.Item(SignatureString)


                    If ToDeleted_Ref(SignatureRefObject, SignaturObject) Then
                        'Dico_LocalName.Item(SignatureString) = SignaturObject
                        Dico.Item(LocalName).Item(SignatureString) = SignaturObject
                        Liste_Signatures_ToDeleted.Add(i)
                    ElseIf ToDeleted_Sub(SignatureRefObject, SignaturObject) Then
                        Dico.Item(LocalName).Item(SignatureRefString) = SignatureRefObject
                        Liste_Signatures_ToDeleted.Add(j)
                    End If


                Next

            Next

            For Each No In Liste_Signatures_ToDeleted
                Dico_LocalName.Remove(Signatures(No))
            Next
        Next


        Dico_Simple = New Dictionary(Of String, Dictionary(Of String, Class_XML_Signature))

        For Each LocalName As String In LocalNames
            Dim Dico_LocalName = Dico.Item(LocalName)

            Dico_Simple.Add(LocalName, New Dictionary(Of String, Class_XML_Signature))

            For Each SignatureObject In Dico_LocalName.Values
                Dico_Simple.Item(LocalName).Add(SignatureObject.ExportToString, SignatureObject)
            Next
        Next


    End Sub


    Private Shared Function ToDeleted_Ref(ByRef SignatureRefObject As Class_XML_Signature, ByRef SignaturObject As Class_XML_Signature) As Boolean
        Dim ListAtribueRef As List(Of String) = SignatureRefObject.Liste_Attribue
        ListAtribueRef.Sort()
        Dim List_ElementRef As List(Of String) = SignatureRefObject.Liste_Element
        List_ElementRef.Sort()
        Dim List_ElementsRef As List(Of String) = SignatureRefObject.Liste_Elements
        List_ElementRef.Sort()

        Dim ListAtribue As List(Of String) = SignaturObject.Liste_Attribue
        ListAtribue.Sort()
        Dim List_Element As List(Of String) = SignaturObject.Liste_Element
        ListAtribue.Sort()
        Dim List_Elements As List(Of String) = SignaturObject.Liste_Elements
        ListAtribue.Sort()


        Dim First_ElementRef As String = List_ElementRef.FirstOrDefault
        Dim First_ElementsRef As String = List_ElementsRef.FirstOrDefault
        Dim First_Element As String = List_Element.FirstOrDefault
        Dim First_Elements As String = List_Elements.FirstOrDefault

        'Si liste avec un seule élément
        If (List_ElementRef.Count = 1 Xor List_ElementsRef.Count = 1) AndAlso (List_Element.Count = 1 Xor List_Elements.Count = 1) Then
            If First_ElementRef <> Nothing AndAlso First_Elements <> Nothing AndAlso First_ElementRef = First_Elements Then
                Return True
            End If
        End If

        'Si Nb Atribue Variable pour les même signature Element et Elemnts 
        If Join(List_ElementRef.ToArray, ";") = Join(List_Element.ToArray, ";") AndAlso
            Join(List_ElementsRef.ToArray, ";") = Join(List_Elements.ToArray, ";") Then
            If Join(ListAtribueRef.ToArray, ";") <> Join(ListAtribue.ToArray, ";") Then
                SignaturObject.Liste_Attribue.AddRange(ListAtribueRef)
                SignaturObject.Liste_Attribue = SignaturObject.Liste_Attribue.Distinct.ToList
                Return True
            End If

        End If
        Return False
    End Function

    Private Shared Function ToDeleted_Sub(ByRef SignatureRefObject As Class_XML_Signature, ByRef SignaturObject As Class_XML_Signature) As Boolean
        Dim ListAtribueRef As List(Of String) = SignatureRefObject.Liste_Attribue
        ListAtribueRef.Sort()
        Dim List_ElementRef As List(Of String) = SignatureRefObject.Liste_Element
        List_ElementRef.Sort()
        Dim List_ElementsRef As List(Of String) = SignatureRefObject.Liste_Elements
        List_ElementRef.Sort()

        Dim ListAtribue As List(Of String) = SignaturObject.Liste_Attribue
        ListAtribue.Sort()
        Dim List_Element As List(Of String) = SignaturObject.Liste_Element
        ListAtribue.Sort()
        Dim List_Elements As List(Of String) = SignaturObject.Liste_Elements
        ListAtribue.Sort()


        Dim First_ElementRef As String = List_ElementRef.FirstOrDefault
        Dim First_ElementsRef As String = List_ElementsRef.FirstOrDefault
        Dim First_Element As String = List_Element.FirstOrDefault
        Dim First_Elements As String = List_Elements.FirstOrDefault


        'Si liste avec un seule élément
        If (List_ElementRef.Count = 1 Xor List_ElementsRef.Count = 1) AndAlso (List_Element.Count = 1 Xor List_Elements.Count = 1) Then
            If First_ElementsRef <> Nothing AndAlso First_Element <> Nothing AndAlso First_ElementsRef = First_Element Then
                Return True
            End If
        End If



        Return False
    End Function

    Public Shared Sub GenerateVBFiles(PathDir As String)
        Dim txt As String = ""
        txt += "Module Module_Function" + vbCrLf
        For Each LocalName In Dico_Simple.Keys


            Dim NBClassAnalyzed = Dico_Simple.Item(LocalName).Count
            If NBClassAnalyzed = 1 Then
                Dim ClassAnalyzed = Dico_Simple.Item(LocalName).First.Value
                Dim NameClass = "ClassSub_" + ClassAnalyzed.LocalName.Normalisation
                Dim NameFile = PathDir + NameClass + ".vb"
                My.Computer.FileSystem.WriteAllText(NameFile, ExportToVB(ClassAnalyzed), False)
            Else
                Dim FunctionGet_Analyzed = Dico_Simple.Item(LocalName).First.Value
                Dim NameFunction = "FunctionGet_" + FunctionGet_Analyzed.LocalName.Normalisation


                txt += "Public Function " + NameFunction + "(Element As XElement) As Object" + vbCrLf
                txt += "   Dim Signature = GetSignature(Element)" + vbCrLf

                Dim Compteur As Integer = 0
                For Each ClassAnalyzed In Dico_Simple.Item(LocalName).Values
                    Dim Suffix = "_" + Intenger2String(Compteur)
                    Dim NameClass = "ClassSub_" + ClassAnalyzed.LocalName.Normalisation + Suffix
                    Dim NameFile = PathDir + NameClass + Suffix + ".vb"


                    My.Computer.FileSystem.WriteAllText(NameFile, ExportToVB(ClassAnalyzed, Suffix), False)


                    Dim Signature = ClassAnalyzed.ExportSignature
                    txt += "   If Signature = " + Signature.GString + " Then" + vbCrLf
                    txt += "      Return GetClassByName(" + NameClass.GString() + ",Element)" + vbCrLf
                    txt += "   End If" + vbCrLf

                    Compteur += 1
                Next

                txt += "Return New Object" + vbCrLf
                txt += "End Function" + vbCrLf
                Dim NameFileFunction = PathDir + NameFunction + ".vb"
            End If
        Next
        txt += "End Module" + vbCrLf
        My.Computer.FileSystem.WriteAllText(PathDir + "Module_Function.vb", txt, False)
    End Sub

    Private Shared Function Intenger2String(i As Integer) As String
        If i <= 9 Then
            Return "00" + i.ToString
        ElseIf i <= 99 Then
            Return "0" + i.ToString
        Else
            Return i.ToString
        End If
    End Function

    Private Shared Function ExportToVB(Signature As Class_XML_Signature, Optional Suffix As String = "") As String
        Dim Tab As Integer = 0
        Dim txt As String = ""
        txt += TabString(Tab) + "Public Class ClassSub_" + Signature.LocalName.Normalisation + Suffix + vbCrLf
        Tab += 1
        txt += TabString(Tab) + "Property RawValue As String" + vbCrLf
        txt += TabString(Tab) + "Property value As String" + vbCrLf
        If Signature.ElementMain.IsElementSimpleXML Then


            txt += TabString(Tab) + "Public Sub New(Element as XElement)" + vbCrLf
            Tab += 1
            txt += TabString(Tab) + "Me.RawValue=Element.Tostring" + vbCrLf
            txt += TabString(Tab) + "Me.Value=Element.Value" + vbCrLf
            Tab -= 1
            txt += TabString(Tab) + "End Sub" + vbCrLf
        Else
            For Each Attribue In Signature.Liste_Attribue
                Dim AttribueName_Norm = Attribue.Normalisation
                txt += TabString(Tab) + "Property Attribue_" + AttribueName_Norm + " As String = " + "".GString + vbCrLf
            Next

            For Each ElementNonListe In Signature.Liste_Element
                Dim ElementName_Norm = ElementNonListe.Normalisation
                Dim IsClassMultiple As Boolean = Dico_Simple.Item(ElementNonListe).Count <> 1
                'Si Class Multiple --> Utilisation du type Object
                If IsClassMultiple Then
                    txt += TabString(Tab) + "Property Element_" + ElementName_Norm + " As Object" + vbCrLf
                Else 'Sinon utilisation de la class spesifique
                    txt += TabString(Tab) + "Property Element_" + ElementName_Norm + " As ClassSub_" + ElementName_Norm + vbCrLf
                End If
            Next


            For Each ElementsLinte In Signature.Liste_Elements
                Dim ElementsName_Norm = ElementsLinte.Normalisation
                Dim IsClassMultiple As Boolean = Dico_Simple.Item(ElementsLinte).Count <> 1
                'Si Class Multiple --> Utilisation du type Object
                If IsClassMultiple Then
                    txt += TabString(Tab) + "Property Elements_" + ElementsName_Norm + " As New List( Of Object)" + vbCrLf
                Else 'Sinon utilisation de la class spesifique
                    txt += TabString(Tab) + "Property Elements_" + ElementsName_Norm + " As New List( Of ClassSub_" + ElementsName_Norm + ")" + vbCrLf
                End If
            Next


            txt += TabString(Tab) + "Public Sub New(Element as XElement)" + vbCrLf
            Tab += 1

            txt += TabString(Tab) + "Me.RawValue=Element.Tostring" + vbCrLf
            txt += TabString(Tab) + "Me.Value=Element.Value" + vbCrLf

            If Signature.Liste_Attribue.Count > 0 Then
                txt += TabString(Tab) + "For Each Attribue in Element.Attributes" + vbCrLf
                Tab += 1
                For Each Attribue In Signature.Liste_Attribue
                    Dim AttribueName_Norm = Attribue.Normalisation
                    txt += TabString(Tab) + "If  Attribue.Name.LocalName = " + Attribue.GString + " Then" + vbCrLf
                    Tab += 1
                    txt += TabString(Tab) + "Me.Attribue_" + AttribueName_Norm + "  = Attribue.value" + vbCrLf
                    Tab -= 1
                    txt += TabString(Tab) + "End If" + vbCrLf
                Next
                Tab -= 1
                txt += TabString(Tab) + "Next" + vbCrLf
            End If


            If Signature.Liste_Element.Count > 0 Then
                txt += TabString(Tab) + "For Each ElementSub in Element.Elements" + vbCrLf
                Tab += 1
                For Each ElementNonListe In Signature.Liste_Element
                    Dim ElementNonListe_Norm = ElementNonListe.Normalisation
                    txt += TabString(Tab) + "If  ElementSub.Name.LocalName = " + ElementNonListe.GString + " Then" + vbCrLf
                    Tab += 1
                    Dim IsClassMultiple As Boolean = Dico_Simple.Item(ElementNonListe).Count <> 1
                    If IsClassMultiple Then
                        txt += TabString(Tab) + "Element_" + ElementNonListe_Norm + " = FunctionGet_" + ElementNonListe_Norm + "(ElementSub)" + vbCrLf
                    Else
                        txt += TabString(Tab) + "Element_" + ElementNonListe_Norm + " = New ClassSub_" + ElementNonListe_Norm + "(ElementSub)" + vbCrLf
                    End If

                    Tab -= 1
                    txt += TabString(Tab) + "End If" + vbCrLf
                Next
                Tab -= 1
                txt += TabString(Tab) + "Next" + vbCrLf
            End If
            If Signature.Liste_Elements.Count > 0 Then
                txt += TabString(Tab) + "For Each ElementSub in Element.Elements" + vbCrLf
                Tab += 1
                For Each ElementListe In Signature.Liste_Elements
                    Dim ElementListe_Norm = ElementListe.Normalisation
                    txt += TabString(Tab) + "If  ElementSub.Name.LocalName = " + ElementListe.GString + " Then" + vbCrLf
                    Tab += 1
                    Dim IsClassMultiple As Boolean = Dico_Simple.Item(ElementListe).Count <> 1
                    If IsClassMultiple Then
                        txt += TabString(Tab) + "Elements_" + ElementListe_Norm + ".Add(FunctionGet_" + ElementListe_Norm + "(ElementSub))" + vbCrLf
                    Else
                        txt += TabString(Tab) + "Elements_" + ElementListe_Norm + ".Add(New ClassSub_" + ElementListe_Norm + "(ElementSub))" + vbCrLf
                    End If

                    Tab -= 1
                    txt += TabString(Tab) + "End If" + vbCrLf
                Next
                Tab -= 1
                txt += TabString(Tab) + "Next" + vbCrLf
            End If
            Tab -= 1
            txt += TabString(Tab) + "End Sub" + vbCrLf
        End If
        Tab -= 1
        txt += TabString(Tab) + "End Class" + vbCrLf
        Return txt
    End Function


    Private Shared Function TabString(Nb As Integer) As String
        Dim txt As String = ""
        For i As Integer = 0 To Nb - 1
            txt += "    "
        Next
        Return txt
    End Function
End Class
