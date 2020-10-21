#Region "Imports"
Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.Text
Imports System.Data
Imports System.Data.OracleClient
'Imports Oracle.DataAccess.Client
Imports System.IO
#End Region

Namespace DataAccess
    Public NotInheritable Class OracleHelpers

        Private Sub New()
        End Sub

        Public Shared Function ExecuteSQLDT(ByVal connQGUAR As ConnectionQGUAR, ByVal sqlString As String, ByVal args As Hashtable, ByRef tableData As DataTable) As ProcedureResult

            Dim result As ProcedureResult = New ProcedureResult
            tableData.Locale = System.Globalization.CultureInfo.CurrentCulture

            Try
                Dim cmd As OracleCommand = Nothing

                If connQGUAR.OraTran IsNot Nothing Then
                    'TODO:FIX
                    cmd = New OracleCommand(sqlString, connQGUAR.OraConn) ', connQGUAR.OraTran)
                Else
                    cmd = New OracleCommand(sqlString, connQGUAR.OraConn)
                End If
                'PrepareParams(cmd, parameters)
                For Each key As Object In args.Keys
                    cmd.Parameters.AddWithValue(key.ToString(), args(key))
                Next
                cmd.Parameters.Add("kursor", OracleType.Cursor).Direction = ParameterDirection.Output
                cmd.Parameters.Add("status", OracleType.Number).Direction = ParameterDirection.Output
                cmd.Parameters.Add("status_opis", OracleType.VarChar, 4000).Direction = ParameterDirection.Output
                'cmd.Parameters.Add("Kursor", OracleDbType.RefCursor).Direction = ParameterDirection.Output
                'cmd.Parameters.Add("status", OracleDbType.Int16).Direction = ParameterDirection.Output
                'cmd.Parameters.Add("status_opis", OracleDbType.NVarchar2, 4000).Direction = ParameterDirection.Output
                Using da As New OracleDataAdapter(cmd)
                    da.Fill(tableData)
                End Using
                result.Status = CType(cmd.Parameters("status").Value, Status)
                result.Message = IIf(IsDBNull(cmd.Parameters("status_opis").Value), "", cmd.Parameters("status_opis").Value)
                cmd.Dispose()
            Catch ex As Exception
                result.Status = Status.Error
                result.Message = String.Format("Database Communications Error: {0}", ex.Message)
            Finally

            End Try
            Return result
        End Function

        Public Shared Function ExecuteProcDT(ByVal connQGUAR As ConnectionQGUAR, ByVal procedure As String, ByVal parameters As Hashtable, ByRef tableData As DataTable) As ProcedureResult

            Dim result As ProcedureResult = New ProcedureResult
            tableData.Locale = System.Globalization.CultureInfo.CurrentCulture

            Try
                Dim cmd As OracleCommand = Nothing

                If connQGUAR.OraTran IsNot Nothing Then
                    cmd = New OracleCommand(procedure, connQGUAR.OraConn) ', connQGUAR.OraTran)
                Else
                    cmd = New OracleCommand(procedure, connQGUAR.OraConn)
                End If

                cmd.CommandTimeout = 3600
                cmd.CommandType = System.Data.CommandType.StoredProcedure
                'cmd.BindByName = True

                'PrepareParams(cmd, parameters)
                For Each key As Object In parameters.Keys
                    cmd.Parameters.AddWithValue(key.ToString(), parameters(key))
                Next
                cmd.Parameters.Add("kursor", OracleType.Cursor).Direction = ParameterDirection.Output
                cmd.Parameters.Add("status", OracleType.Number).Direction = ParameterDirection.Output
                cmd.Parameters.Add("status_opis", OracleType.VarChar, 4000).Direction = ParameterDirection.Output
                Using da As New OracleDataAdapter(cmd)
                    da.Fill(tableData)
                End Using

                result.Status = CType(cmd.Parameters("status").Value, Status)
                result.Message = IIf(IsDBNull(cmd.Parameters("status_opis").Value), "", cmd.Parameters("status_opis").Value)

            Catch ex As Exception
                result.Status = Status.Error
                result.Message = String.Format("Database Communications Error: {0}", ex.Message)
            Finally

            End Try
            Return result
        End Function

    End Class
End Namespace
