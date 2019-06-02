Imports System.Runtime.CompilerServices

Public Module Module_Extension_XML
    <Extension>
    Public Function GetSignature(Element As XElement) As Class_XML_Signature
        Return New Class_XML_Signature(Element, False)
    End Function
    <Extension>
    Public Function IsElementSimple(Element As XElement) As Boolean
        Return Element.Attributes.Count = 0 And Element.Elements.Count = 0
    End Function
    <Extension>
    Public Function IsElementSimpleXML(ElementString As String) As Boolean
        Dim Element = XElement.Parse(ElementString)
        Return Element.Attributes.Count = 0 And Element.Elements.Count = 0
    End Function
End Module
