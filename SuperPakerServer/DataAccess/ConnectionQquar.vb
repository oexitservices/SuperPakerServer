#Region "Imports"
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Data.OracleClient
'Imports Oracle.DataAccess.Client
Imports System.Configuration
#End Region

Namespace DataAccess
    Public Class ConnectionQGUAR
        Implements IDisposable

        'Dim cmd As OracleCommand
        Private _oraConn As OracleConnection
        Private _oraTran As OracleTransaction
        Private disposed As Boolean

        Public Sub New()
            OraConn = New OracleConnection()
        End Sub

        Public Sub Open()

            If (OraConn.State = System.Data.ConnectionState.Open) Then
                Return
            End If

            'dim connStr as string = ConfigurationManager.ConnectionStrings["default"].ConnectionString;
            'Dim connStr As String = "Server=DB_SERVER;Database=DB_NAME;User Id=DB_USER;Pwd=DB_PASS;Trusted_Connection=False;timeout=30;" '600;"
            'dim connStr as string = "Server=DB_SERVER;Database=DB_NAME;User Id=DB_USER;Pwd=DB_PASS;Trusted_Connection=False;Encrypt=yes;TrustServerCertificate=yes;";
            Dim connStr = Config.ConnectionStringQGUAR
            'connStr = connStr.Replace("DB_SERVER", Config.Host)
            'connStr = connStr.Replace("DB_NAME", Config.Database)
            'connStr = connStr.Replace("DB_USER", Config.User)
            'connStr = connStr.Replace("DB_PASS", Config.Password)

            OraConn.ConnectionString = connStr
            OraConn.Open()

        End Sub

        Public Sub BeginTransaction()

            Me.Open()
            Me._oraTran = Me.OraConn.BeginTransaction()

        End Sub

        Public Sub CommitTransaction()

            Me._oraTran.Commit()
            Me._oraTran = Nothing

        End Sub

        Public Sub RollbackTransaction()

            Me._oraTran.Rollback()
            Me._oraTran = Nothing

        End Sub

        Public Sub Close()

            If Not (_oraConn Is Nothing) Then

                _oraConn.Close()

            End If

        End Sub

        Property OraConn As OracleConnection

            Get
                Return _oraConn
            End Get
            Set(ByVal value As OracleConnection)
                _oraConn = value
            End Set

        End Property

        Property OraTran As OracleTransaction

            Get
                Return _oraTran
            End Get
            Set(ByVal value As OracleTransaction)
                _oraTran = value
            End Set
        End Property

        Public Overloads Sub Dispose() Implements IDisposable.Dispose


            If Not (Me.OraConn Is Nothing) And Not (Me.OraTran Is Nothing) Then
                Me.RollbackTransaction()
                Dispose(True)
                GC.SuppressFinalize(Me)
            End If

        End Sub

        Private Overloads Sub Dispose(ByVal disposing As Boolean)


            If (Not Me.disposed) Then

                If (disposing) Then

                    If Not (OraConn Is Nothing) Then

                        OraConn.Close()
                        OraConn.Dispose()
                    End If

                End If

                disposed = True

            End If

        End Sub

    End Class
End Namespace
