Imports System.Data.SqlClient

Namespace EmailManager

    Public Class DalEmail
        'Private Property SqlConnString As String
        Public Property sqlConn As SqlConnection

        Public Sub New(ByVal odbcName As String, ByVal Optional dbUser As String = "", ByVal Optional dbPassword As String = "")
            Try
                Dim odbcConvert As OdbcConverter.OdbcConverter = New OdbcConverter.OdbcConverter()
                sqlConn = New SqlConnection(odbcConvert.GetSqlConnectionString(odbcName, dbUser, dbPassword))
                sqlConn.Open()
            Catch ex As Exception
                Throw New Exception(TraceException(ex.Message))
            End Try

        End Sub

        Friend Function GetEmailParams(ByRef dsEmail As DsEmail, ByVal operationCode As Integer, ByVal Optional userName As String = "") As Boolean

            Using sqlCmd As SqlCommand = New SqlCommand("GetEmailParams", sqlConn)
                Dim da As SqlDataAdapter

                Try
                    sqlCmd.CommandType = CommandType.StoredProcedure
                    sqlCmd.Parameters.Add("@OperationCode", SqlDbType.Int)
                    sqlCmd.Parameters("@OperationCode").Value = operationCode
                    sqlCmd.Parameters.Add("@UserName", SqlDbType.NVarChar, 100)
                    sqlCmd.Parameters("@UserName").Value = userName

                    da = New SqlDataAdapter(sqlCmd)
                    da.Fill(dsEmail, "SendEmailsParams")

                    Return dsEmail.SendEmailsParams.Count = 1
                    sqlConn.Close()
                    sqlConn.Dispose()
                Catch ex As Exception
                    sqlConn.Close()
                    sqlConn.Dispose()
                    Throw New Exception(TraceException(ex.Message))
                End Try
            End Using
        End Function

        Friend Function CreateKblOperation(ByVal Operation As Integer, ByVal ProcUserID As Integer, Optional ByVal mnt As Integer = 0, Optional ByVal thum As Integer = 0, Optional ByVal kabala As Integer = 0) As Integer

            Using sqlCmd As SqlCommand = New SqlCommand("InsertIntoKblEmailControlAndData", sqlConn)
                Try
                    sqlCmd.CommandType = CommandType.StoredProcedure
                    sqlCmd.Parameters.Add("@OperationCode", SqlDbType.Int)
                    sqlCmd.Parameters("@OperationCode").Value = Operation
                    sqlCmd.Parameters.Add("@ProcUserName", SqlDbType.Int)
                    sqlCmd.Parameters("@ProcUserName").Value = ProcUserID
                    If mnt > 0 Then
                        sqlCmd.Parameters.Add("@mnt", SqlDbType.Int)
                        sqlCmd.Parameters("@mnt").Value = mnt
                    End If
                    If thum > 0 Then
                        sqlCmd.Parameters.Add("@thum", SqlDbType.Int)
                        sqlCmd.Parameters("@thum").Value = thum
                    End If
                    If kabala > 0 Then
                        sqlCmd.Parameters.Add("@kabala", SqlDbType.Int)
                        sqlCmd.Parameters("@kabala").Value = kabala
                    End If
                    Return CInt(sqlCmd.ExecuteScalar())
                    sqlConn.Close()
                    sqlConn.Dispose()
                Catch ex As Exception
                    sqlConn.Close()
                    sqlConn.Dispose()
                    Throw New Exception(TraceException(ex.Message))
                End Try
            End Using
        End Function

    End Class

End Namespace
