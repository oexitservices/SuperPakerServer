#Region "Imports"
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Data.SqlClient
Imports System.Configuration
#End Region

Namespace DataAccess
    Public Class ConnectionSuperPaker
        Implements IDisposable

        Private _sqlConn As SqlConnection
        Private _sqlTran As SqlTransaction
        Private disposed As Boolean

        Public Sub New()
            SqlConn = New SqlConnection()
        End Sub

        Public Sub Open()

            If (SqlConn.State = System.Data.ConnectionState.Open) Then
                Return
            End If

            'dim connStr as string = ConfigurationManager.ConnectionStrings["default"].ConnectionString;
            'Dim connStr As String = "Server=DB_SERVER;Database=DB_NAME;User Id=DB_USER;Pwd=DB_PASS;Trusted_Connection=False;timeout=30;" '600;"
            'dim connStr as string = "Server=DB_SERVER;Database=DB_NAME;User Id=DB_USER;Pwd=DB_PASS;Trusted_Connection=False;Encrypt=yes;TrustServerCertificate=yes;";
            Dim connStr = Config.ConnectionStringSuperPaker
            'connStr = connStr.Replace("DB_SERVER", Config.Host)
            'connStr = connStr.Replace("DB_NAME", Config.Database)
            'connStr = connStr.Replace("DB_USER", Config.User)
            'connStr = connStr.Replace("DB_PASS", Config.Password)

            SqlConn.ConnectionString = connStr
            SqlConn.Open()

        End Sub

        Public Sub BeginTransaction()

            Me.Open()
            Me._sqlTran = Me.SqlConn.BeginTransaction()

        End Sub

        Public Sub CommitTransaction()

            Me._sqlTran.Commit()
            Me._sqlTran = Nothing

        End Sub

        Public Sub RollbackTransaction()

            Me._sqlTran.Rollback()
            Me._sqlTran = Nothing

        End Sub

        Public Sub Close()

            If Not (_sqlConn Is Nothing) Then

                _sqlConn.Close()

            End If

        End Sub

        Property SqlConn As SqlConnection

            Get
                Return _sqlConn
            End Get
            Set(ByVal value As SqlConnection)
                _sqlConn = value
            End Set

        End Property

        Property SqlTran As SqlTransaction

            Get
                Return _sqlTran
            End Get
            Set(ByVal value As SqlTransaction)
                _sqlTran = value
            End Set
        End Property

        Public Overloads Sub Dispose() Implements IDisposable.Dispose


            If Not (Me.SqlConn Is Nothing) And Not (Me.SqlTran Is Nothing) Then
                Me.RollbackTransaction()
                Dispose(True)
                GC.SuppressFinalize(Me)
            End If

        End Sub

        Private Overloads Sub Dispose(ByVal disposing As Boolean)


            If (Not Me.disposed) Then

                If (disposing) Then

                    If Not (SqlConn Is Nothing) Then

                        SqlConn.Close()
                        SqlConn.Dispose()
                    End If

                End If

                disposed = True

            End If

        End Sub

    End Class
End Namespace
