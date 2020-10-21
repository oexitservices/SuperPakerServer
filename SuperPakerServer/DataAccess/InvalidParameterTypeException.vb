#Region "Imports"
Imports System
Imports System.Collections.Generic
Imports System.Text
#End Region

Namespace DataAccess
    <Serializable()> _
    Public Class InvalidParameterTypeException
        Inherits Exception
        Public Sub New()

            MyBase.New()
        End Sub

        Public Sub New(ByVal message As String)

            MyBase.New(message)
        End Sub

        Public Sub New(ByVal message As String, ByVal innerException As Exception)

            MyBase.New(message, innerException)
        End Sub

        Protected Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)

            MyBase.New(info, context)
        End Sub

    End Class
End Namespace
