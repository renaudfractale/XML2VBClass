Public Module Module_XML_LIB
    Public Function GetNBOcurenceElement(ElementMain As XElement, ElementAnalized As XElement) As Integer
        Dim NbOcurence As Integer = 0
        For Each Element As XElement In ElementMain.Elements
            If Element.Name.LocalName = ElementAnalized.Name.LocalName Then
                NbOcurence += 1
            End If
        Next
        Return NbOcurence
    End Function


End Module
