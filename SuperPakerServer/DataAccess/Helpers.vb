#Region "Imports"
Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
#End Region

Namespace DataAccess
    Public NotInheritable Class Helpers

        Private Sub New()
        End Sub

        Public Shared Function ExecuteSql(ByVal conn As ConnectionSuperPaker, ByVal sesjaId As Byte(), ByVal sql As String, ByVal args As Hashtable) As ProcedureResult

            Dim result As ProcedureResult = New ProcedureResult
            Dim cmd As SqlCommand = Nothing

            Try
                If conn.SqlTran IsNot Nothing Then
                    cmd = New SqlCommand(sql, conn.SqlConn, conn.SqlTran)
                Else
                    cmd = New SqlCommand(sql, conn.SqlConn)
                End If

                PrepareParams(cmd, args)
                cmd.Parameters.AddWithValue("@SESJA", sesjaId)
                cmd.Parameters.Add("@STATUS", SqlDbType.Int).Direction = ParameterDirection.Output
                cmd.Parameters.Add("@STATUS_OPIS", SqlDbType.NVarChar, 4000).Direction = ParameterDirection.Output
                cmd.ExecuteNonQuery()
            Catch extrans As SqlException
                If result.Message = "" Then
                    result.Status = Status.Error
                    result.Message = extrans.Message
                End If
            Catch ex As Exception
                result.Status = Status.Error
                result.Message = String.Format("Database Communications Error: {0}", ex.Message)
            Finally

            End Try
            Return result


        End Function

        Public Shared Function ExecuteProcDT(ByVal conn As ConnectionSuperPaker, ByVal sesjaId As Byte(), ByRef cmd As SqlCommand, ByVal procedure As String, ByVal parameters As Hashtable, ByRef tableData As DataTable) As ProcedureResult

            Dim result As ProcedureResult = New ProcedureResult
            tableData.Locale = System.Globalization.CultureInfo.CurrentCulture

            Try
                If conn.SqlTran IsNot Nothing Then
                    cmd = New SqlCommand(procedure, conn.SqlConn, conn.SqlTran)
                Else
                    cmd = New SqlCommand(procedure, conn.SqlConn)
                End If
                cmd.CommandTimeout = 3600
                cmd.CommandType = System.Data.CommandType.StoredProcedure

                PrepareParams(cmd, parameters)
                cmd.Parameters.AddWithValue("@SESJA", sesjaId)
                cmd.Parameters.Add("@STATUS", SqlDbType.Int).Direction = ParameterDirection.Output
                cmd.Parameters.Add("@STATUS_OPIS", SqlDbType.NVarChar, 4000).Direction = ParameterDirection.Output
                Using da As New SqlDataAdapter(cmd)
                    da.Fill(tableData)
                End Using

                result.Status = CType(cmd.Parameters("@STATUS").Value, Status)
                result.Message = DirectCast(cmd.Parameters("@STATUS_OPIS").Value, String)

            Catch extrans As SqlException
                If cmd Is Nothing OrElse cmd.Parameters Is Nothing Then
                    result.Status = Status.Error
                    result.Message = String.Format("Database Communications Error: (Błąd procedury {0})", procedure) 'ex.Message)
                Else
                    result.Status = CType(cmd.Parameters("@STATUS").Value, Status)
                    result.Message = DirectCast(cmd.Parameters("@STATUS_OPIS").Value, String)
                End If
                If result.Message = "" Then
                    result.Status = Status.Error
                    result.Message = extrans.Message
                End If
            Catch ex As Exception
                result.Status = Status.Error
                result.Message = String.Format("Database Communications Error: {0}", ex.Message)
            Finally

            End Try
            Return result
        End Function

        Public Shared Function ExecuteProcDS(ByVal conn As ConnectionSuperPaker, ByVal sesjaId As Byte(), ByVal procedure As String, ByVal parameters As Hashtable, ByRef sData As DataSet) As ProcedureResult

            Dim result As ProcedureResult = New ProcedureResult
            'tableData.Locale = System.Globalization.CultureInfo.CurrentCulture
            Dim cmd As SqlCommand = Nothing

            Try
                If conn.SqlTran IsNot Nothing Then
                    cmd = New SqlCommand(procedure, conn.SqlConn, conn.SqlTran)
                Else
                    cmd = New SqlCommand(procedure, conn.SqlConn)
                End If
                cmd.CommandTimeout = 3600
                cmd.CommandType = System.Data.CommandType.StoredProcedure

                PrepareParams(cmd, parameters)
                cmd.Parameters.AddWithValue("@SESJA", sesjaId)
                cmd.Parameters.Add("@STATUS", SqlDbType.Int).Direction = ParameterDirection.Output
                cmd.Parameters.Add("@STATUS_OPIS", SqlDbType.NVarChar, 4000).Direction = ParameterDirection.Output
                Using da As New SqlDataAdapter(cmd)
                    da.Fill(sData)
                End Using
                result.Status = CType(cmd.Parameters("@STATUS").Value, Status)
                result.Message = DirectCast(cmd.Parameters("@STATUS_OPIS").Value, String)
            Catch extrans As SqlException
                If cmd Is Nothing OrElse cmd.Parameters Is Nothing Then
                    result.Status = Status.Error
                    result.Message = String.Format("Database Communications Error: (Błąd procedury {0})", procedure) 'ex.Message)
                Else
                    result.Status = CType(cmd.Parameters("@STATUS").Value, Status)
                    result.Message = DirectCast(cmd.Parameters("@STATUS_OPIS").Value, String)
                End If
                If result.Message = "" Then
                    result.Status = Status.Error
                    result.Message = extrans.Message
                End If
            Catch ex As Exception
                result.Status = Status.Error
                result.Message = String.Format("Database Communications Error: {0}", ex.Message)
            Finally

            End Try
            Return result
        End Function

        Public Shared Function ExecuteProcDSOutput(ByVal conn As ConnectionSuperPaker, ByVal sesjaId As Byte(), ByVal procedure As String, ByVal parameters As Hashtable, ByRef sData As DataSet) As ProcedureResult

            Dim result As ProcedureResult = New ProcedureResult
            'tableData.Locale = System.Globalization.CultureInfo.CurrentCulture
            Dim cmd As SqlCommand = Nothing

            Try
                If conn.SqlTran IsNot Nothing Then
                    cmd = New SqlCommand(procedure, conn.SqlConn, conn.SqlTran)
                Else
                    cmd = New SqlCommand(procedure, conn.SqlConn)
                End If
                cmd.CommandTimeout = 3600
                cmd.CommandType = System.Data.CommandType.StoredProcedure

                PrepareParams(cmd, parameters)
                cmd.Parameters.AddWithValue("@SESJA", sesjaId)
                cmd.Parameters.Add("@STATUS", SqlDbType.Int).Direction = ParameterDirection.Output
                cmd.Parameters.Add("@STATUS_OPIS", SqlDbType.NVarChar, 4000).Direction = ParameterDirection.Output
                Using da As New SqlDataAdapter(cmd)
                    da.Fill(sData)
                End Using
                result.Status = CType(cmd.Parameters("@STATUS").Value, Status)
                result.Message = DirectCast(cmd.Parameters("@STATUS_OPIS").Value, String)
                GetParamOutValues(cmd, parameters)

            Catch extrans As SqlException
                If cmd Is Nothing OrElse cmd.Parameters Is Nothing Then
                    result.Status = Status.Error
                    result.Message = String.Format("Database Communications Error: (Błąd procedury {0})", procedure) 'ex.Message)
                Else
                    result.Status = CType(cmd.Parameters("@STATUS").Value, Status)
                    result.Message = DirectCast(cmd.Parameters("@STATUS_OPIS").Value, String)
                End If
                If result.Message = "" Then
                    result.Status = Status.Error
                    result.Message = extrans.Message
                End If
            Catch ex As Exception
                result.Status = Status.Error
                result.Message = String.Format("Database Communications Error: {0}", ex.Message)
            Finally

            End Try
            Return result
        End Function


        Public Shared Function ExecuteProcDT(ByVal conn As ConnectionSuperPaker, ByVal sesjaId As Byte(), ByVal procedure As String, ByVal parameters As Hashtable, ByRef tableData As DataTable) As ProcedureResult

            Dim result As ProcedureResult = New ProcedureResult
            tableData.Locale = System.Globalization.CultureInfo.CurrentCulture
            Dim cmd As SqlCommand = Nothing
            Try


                If conn.SqlTran IsNot Nothing Then
                    cmd = New SqlCommand(procedure, conn.SqlConn, conn.SqlTran)
                Else
                    cmd = New SqlCommand(procedure, conn.SqlConn)
                End If
                cmd.CommandTimeout = 3600
                cmd.CommandType = System.Data.CommandType.StoredProcedure

                PrepareParams(cmd, parameters)
                cmd.Parameters.AddWithValue("@SESJA", sesjaId)
                cmd.Parameters.Add("@STATUS", SqlDbType.Int).Direction = ParameterDirection.Output
                cmd.Parameters.Add("@STATUS_OPIS", SqlDbType.NVarChar, 4000).Direction = ParameterDirection.Output
                Using da As New SqlDataAdapter(cmd)
                    da.Fill(tableData)
                End Using
                result.Status = CType(cmd.Parameters("@STATUS").Value, Status)
                result.Message = DirectCast(cmd.Parameters("@STATUS_OPIS").Value, String)
            Catch extrans As SqlException
                If cmd Is Nothing OrElse cmd.Parameters Is Nothing Then
                    result.Status = Status.Error
                    result.Message = String.Format("Database Communications Error: (Błąd procedury {0})", procedure) 'ex.Message)
                Else
                    result.Status = CType(cmd.Parameters("@STATUS").Value, Status)
                    result.Message = DirectCast(cmd.Parameters("@STATUS_OPIS").Value, String)
                End If
                If result.Message = "" Then
                    result.Status = Status.Error
                    result.Message = extrans.Message
                End If
            Catch ex As Exception
                result.Status = Status.Error
                result.Message = String.Format("Database Communications Error: {0}", ex.Message)
            Finally

            End Try
            Return result
        End Function

        Public Shared Function ExecuteProcDTOutput(ByVal conn As ConnectionSuperPaker, ByVal sesjaId As Byte(), ByVal procedure As String, ByVal parameters As Hashtable, ByRef tableData As DataTable) As ProcedureResult

            Dim result As ProcedureResult = New ProcedureResult
            tableData.Locale = System.Globalization.CultureInfo.CurrentCulture
            Dim cmd As SqlCommand = Nothing
            Try


                If conn.SqlTran IsNot Nothing Then
                    cmd = New SqlCommand(procedure, conn.SqlConn, conn.SqlTran)
                Else
                    cmd = New SqlCommand(procedure, conn.SqlConn)
                End If
                cmd.CommandTimeout = 3600
                cmd.CommandType = System.Data.CommandType.StoredProcedure

                PrepareParams(cmd, parameters)
                cmd.Parameters.AddWithValue("@SESJA", sesjaId)
                cmd.Parameters.Add("@STATUS", SqlDbType.Int).Direction = ParameterDirection.Output
                cmd.Parameters.Add("@STATUS_OPIS", SqlDbType.NVarChar, 4000).Direction = ParameterDirection.Output
                Using da As New SqlDataAdapter(cmd)
                    da.Fill(tableData)
                End Using
                result.Status = CType(cmd.Parameters("@STATUS").Value, Status)
                result.Message = DirectCast(cmd.Parameters("@STATUS_OPIS").Value, String)
                GetParamOutValues(cmd, parameters)
            Catch extrans As SqlException
                If cmd Is Nothing OrElse cmd.Parameters Is Nothing Then
                    result.Status = Status.Error
                    result.Message = String.Format("Database Communications Error: (Błąd procedury {0})", procedure) 'ex.Message)
                Else
                    result.Status = CType(cmd.Parameters("@STATUS").Value, Status)
                    result.Message = DirectCast(cmd.Parameters("@STATUS_OPIS").Value, String)
                End If
                If result.Message = "" Then
                    result.Status = Status.Error
                    result.Message = extrans.Message
                End If
            Catch ex As Exception
                result.Status = Status.Error
                result.Message = String.Format("Database Communications Error: {0}", ex.Message)
            Finally

            End Try
            Return result
        End Function

        Public Shared Function ExecuteProc(ByVal conn As ConnectionSuperPaker, ByVal sesjaId As Byte(), ByVal procedure As String, ByVal parameters As Hashtable) As ProcedureResult
            Dim result As ProcedureResult = New ProcedureResult
            Dim cmd As SqlCommand = Nothing
            Try

                If conn.SqlTran IsNot Nothing Then
                    cmd = New SqlCommand(procedure, conn.SqlConn, conn.SqlTran)
                Else
                    cmd = New SqlCommand(procedure, conn.SqlConn)
                End If
                cmd.CommandTimeout = 600
                cmd.CommandType = System.Data.CommandType.StoredProcedure

                PrepareParams(cmd, parameters)
                cmd.Parameters.AddWithValue("@SESJA", sesjaId)
                cmd.Parameters.Add("@STATUS", SqlDbType.Int).Direction = ParameterDirection.Output
                cmd.Parameters.Add("@STATUS_OPIS", SqlDbType.NVarChar, 4000).Direction = ParameterDirection.Output
                cmd.ExecuteNonQuery()
                GetParamOutValues(cmd, parameters)
                result.Status = CType(cmd.Parameters("@STATUS").Value, Status)
                result.Message = DirectCast(cmd.Parameters("@STATUS_OPIS").Value, String)
            Catch extrans As SqlException
                If cmd Is Nothing OrElse cmd.Parameters Is Nothing Then
                    result.Status = Status.Error
                    result.Message = String.Format("Database Communications Error: (Błąd procedury {0})", procedure) 'ex.Message)
                Else
                    result.Status = CType(cmd.Parameters("@STATUS").Value, Status)
                    result.Message = DirectCast(cmd.Parameters("@STATUS_OPIS").Value, String)
                End If
                If result.Message = "" Then
                    result.Status = Status.Error
                    result.Message = extrans.Message
                End If
            Catch ex As Exception
                result.Status = Status.Error
                result.Message = String.Format("Database Communications Error: {0}", ex.Message)
            Finally

            End Try
            Return result
        End Function

        Public Shared Sub PrepareParams(ByVal cmd As SqlCommand, ByVal parameters As Hashtable)
            For Each key As Object In parameters.Keys
                Dim param As SqlParameter = Nothing
                Dim typeName As String = "DBNull"
                If parameters(key) IsNot Nothing AndAlso parameters(key) IsNot DBNull.Value Then
                    typeName = parameters(key).[GetType]().Name
                End If
                Dim pName As String = key.ToString()
                Dim pSize As Integer = 0
                If pName.Contains(":") Then
                    Dim arr As String() = pName.Split(":"c)
                    pName = arr(0)
                    pSize = Convert.ToInt32(arr(1), System.Globalization.CultureInfo.CurrentCulture)
                End If
                Select Case typeName
                    Case "String"
                        param = New SqlParameter(pName, SqlDbType.NVarChar)
                        Exit Select
                    Case "Int32"
                        param = New SqlParameter(pName, SqlDbType.Int)
                        Exit Select
                    Case "Int16"
                        param = New SqlParameter(pName, SqlDbType.SmallInt)
                        Exit Select
                    Case "Long"
                        param = New SqlParameter(pName, SqlDbType.BigInt)
                        Exit Select
                    Case "Int64"
                        param = New SqlParameter(pName, SqlDbType.BigInt)
                        Exit Select
                    Case "Decimal"
                        param = New SqlParameter(pName, SqlDbType.[Decimal])
                        Exit Select
                    Case "Double"
                        param = New SqlParameter(pName, SqlDbType.Float)
                        Exit Select
                    Case "DateTime"
                        param = New SqlParameter(pName, SqlDbType.DateTime)
                        Exit Select
                    Case "Boolean"
                        param = New SqlParameter(pName, SqlDbType.Bit)
                        Exit Select
                    Case "Byte[]"
                        param = New SqlParameter(pName, SqlDbType.VarBinary)
                        Exit Select
                    Case "DBNull"
                        param = New SqlParameter(pName, SqlDbType.NVarChar)
                        Exit Select
                    Case "SqlBytes"
                        param = New SqlParameter(pName, SqlDbType.Binary)
                        Exit Select
                    Case "Guid"
                        param = New SqlParameter(pName, SqlDbType.UniqueIdentifier)
                        Exit Select
                    Case "DataTable"
                        param = New SqlParameter(pName, SqlDbType.Structured)
                End Select
                If param IsNot Nothing Then
                    If typeName = "DBNull" Then
                        param.Value = DBNull.Value

                    ElseIf pName.StartsWith("@OUT_", StringComparison.CurrentCulture) Then
                        param.Direction = ParameterDirection.InputOutput
                        param.Value = parameters(key)
                        If pSize <> 0 Then
                            param.Size = pSize
                        End If
                    Else
                        param.Value = parameters(key)
                    End If
                    cmd.Parameters.Add(param)
                Else

                    Throw New InvalidParameterTypeException(String.Format(System.Globalization.CultureInfo.CurrentCulture, "Unsupported type '{0}' for parameter '{1}'", typeName, pName))
                End If
            Next
        End Sub

        Public Shared Sub GetParamOutValues(ByVal cmd As SqlCommand, ByVal parameters As Hashtable)
            For Each param As SqlParameter In cmd.Parameters
                If param.Direction = ParameterDirection.InputOutput Then
                    If param.SqlDbType = SqlDbType.NVarChar Then
                        parameters(param.ParameterName + ":" & param.Size.ToString(System.Globalization.CultureInfo.CurrentCulture)) = param.Value
                    Else
                        parameters(param.ParameterName) = param.Value
                    End If
                End If
            Next
        End Sub

        Public Shared Function getClassInstance(ByVal fileName As [String], ByVal className As String) As Type
            ' Load in the assembly. 

            Dim moduleAssembly As System.Reflection.Assembly = System.Reflection.Assembly.LoadFile(fileName)

            ' Get the types of classes that are in this assembly. 

            Dim types As Type() = moduleAssembly.GetTypes()

            ' Loop through the types in the assembly until we find
            '             * a class that implements a Module.
            '             

            For Each type As Type In types
                If type.Name = className Then
                    Return type
                End If
            Next

            Return Nothing
        End Function

    End Class
End Namespace

