Imports System.IO
Imports System.Runtime.Serialization
Imports System.Runtime.Serialization.Formatters.Binary
Public Class Class_Cloning
    Public Shared Function Clone(Of T)(ByVal inputObj As T) As T
        'creating a Memorystream which works like a temporary storeage '
        Using memStrm As New MemoryStream()
            'Binary Formatter for serializing the object into memory stream '
            Dim binFormatter As New BinaryFormatter(Nothing, New StreamingContext(StreamingContextStates.Clone))

            'talks for itself '
            binFormatter.Serialize(memStrm, inputObj)

            'setting the memorystream to the start of it '
            memStrm.Seek(0, SeekOrigin.Begin)

            'try to cast the serialized item into our Item '
            Try
                Return DirectCast(binFormatter.Deserialize(memStrm), T)
            Catch ex As Exception
                Trace.TraceError(ex.Message)
                Return Nothing
            End Try
        End Using
    End Function
End Class
