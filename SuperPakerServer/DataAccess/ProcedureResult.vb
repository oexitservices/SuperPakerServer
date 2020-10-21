Namespace DataAccess
    ''' <summary>
    ''' Status wywołania procedury
    ''' </summary>
    Public Class ProcedureResult

        Public Property Status() As Status
            Get
                Return m_Status
            End Get
            Set(ByVal value As Status)
                m_Status = value
            End Set
        End Property

        Private m_Status As Status

        Public Property Message() As String
            Get
                Return m_Message
            End Get
            Set(ByVal value As String)
                m_Message = value
            End Set
        End Property

        Private m_Message As String
    End Class

    ''' <summary>
    ''' Enumeracja statusu procedury
    ''' </summary>
    Public Enum Status
        Ok = 0
        [Error] = -1
        SessionExpired = -2
        Message = 1
    End Enum

End Namespace
