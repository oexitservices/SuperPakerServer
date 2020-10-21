#Region "Imports"
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Collections
#End Region

Namespace DataAccess

    Public Module Config

        Private _host As String
        Private _database As String
        Private _user As String
        Private _password As String
        Private _version As String
        Private _assemblyVersion As String
        Private _connStrMSSQL As String
        Private _connStrMSSQLSP As String
        Private _connStrQGUAR As String
        'Private _sesja As Byte()


        Public Property ConnectionString() As String
            Get
                If String.IsNullOrEmpty(_connStrMSSQL) Then
                    Throw New ApplicationException("Przed wywołaniem procedury uzupełnij connectionstring dla SuperPakera")
                End If
                Return _connStrMSSQLSP
            End Get
            Set(ByVal value As String)
                _connStrMSSQLSP = value
            End Set
        End Property

        Public Property ConnectionStringSuperPaker() As String
            Get
                If String.IsNullOrEmpty(_connStrMSSQL) Then
                    Throw New ApplicationException("Przed wywołaniem procedury uzupełnij connectionstring dla SuperPakera")
                End If
                Return _connStrMSSQL
            End Get
            Set(ByVal value As String)
                _connStrMSSQL = value
            End Set
        End Property

        Public Property ConnectionStringQGUAR() As String
            Get
                If String.IsNullOrEmpty(_connStrQGUAR) Then
                    Throw New ApplicationException("Przed wywołaniem procedury uzupełnij connectionstring do QGUARA")
                End If
                Return _connStrQGUAR
            End Get
            Set(ByVal value As String)
                _connStrQGUAR = value
            End Set
        End Property

        'Public Property Sesja() As Byte()
        '    Get
        '        If _sesja Is Nothing OrElse _sesja.Length = 0 Then
        '            Throw New ApplicationException("Przed wywołaniem procedury uzupełnij parametr sesji.")
        '        End If
        '        Return _sesja
        '    End Get
        '    Set(ByVal value As Byte())
        '        _sesja = value
        '    End Set
        'End Property

        Public Property Host() As String
            Get
                Return _host
            End Get
            Set(ByVal value As String)
                _host = value
            End Set
        End Property

        Public Property Database() As String
            Get
                Return _database
            End Get
            Set(ByVal value As String)
                _database = value
            End Set
        End Property

        Public Property User() As String
            Get
                Return _user
            End Get
            Set(ByVal value As String)
                _user = value
            End Set
        End Property

        Public Property Password() As String

            Get
                Return _password
            End Get
            Set(ByVal value As String)
                _password = value
            End Set
        End Property

        Public Property Version() As String
            Get
                Return _version
            End Get
            Set(ByVal value As String)
                _version = value
            End Set
        End Property

        Public Property AssemblyVersion() As String
            Get
                Return _assemblyVersion
            End Get
            Set(ByVal value As String)
                _assemblyVersion = value
            End Set
        End Property

    End Module
End Namespace
