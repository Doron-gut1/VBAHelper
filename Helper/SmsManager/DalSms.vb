Imports System.Data.SqlClient
Imports System.Reflection

Namespace SmsManager
    Public Class DalSms
        Private Property SqlConnString As String
        Public Property SqlConnTrans As SqlConnection

        Public Sub New(ByVal odbcName As String, ByVal Optional dbUser As String = "", ByVal Optional dbPassword As String = "")
            Try
                Dim odbcConvert As OdbcConverter.OdbcConverter = New OdbcConverter.OdbcConverter()
                SqlConnString = odbcConvert.GetSqlConnectionString(odbcName, dbUser, dbPassword)

            Catch ex As Exception
                Throw New Exception(TraceException(ex.Message))
            End Try

        End Sub

        Public Function TestConnection(ByRef errDesc As String) As Boolean
            Using sqlConn As SqlConnection = New SqlConnection(SqlConnString)

                Try
                    sqlConn.Open()
                    Dim cmd As SqlCommand = New SqlCommand("IF EXISTS(
                          SELECT 1 FROM INFORMATION_SCHEMA.TABLES 
                          WHERE TABLE_NAME = @table) 
                          SELECT 1 ELSE SELECT 0", sqlConn)
                    cmd.Parameters.Add("@table", SqlDbType.NVarChar).Value = "ParamSmsType"

                    If CInt(cmd.ExecuteScalar()) = 1 Then
                        Return True
                    Else
                        errDesc = $"Table ParamSmsType does NOT exists in DB"
                        Return False
                    End If

                Catch ex As Exception
                    errDesc = TraceException(ex.Message)
                    Return False
                End Try
            End Using
        End Function

        Public Function GetUrlShortenerApiParam() As String
            Dim urlApi As String
            Using sqlConn As SqlConnection = New SqlConnection(SqlConnString)
                Try
                    sqlConn.Open()
                    Using cmd As SqlCommand = New SqlCommand("SELECT ApiShortenerUrl FROM Param3", sqlConn)
                        urlApi = cmd.ExecuteScalar()
                        If urlApi = "" Then
                            Throw New Exception("Param3.ApiShortenerUrl is empty")
                        End If

                        Return urlApi
                    End Using
                Catch ex As Exception
                    Throw New Exception(TraceException(ex.Message))
                End Try
            End Using
        End Function

        Public Function GetSqlConnection(ByRef sqlTrans As SqlTransaction) As SqlConnection
            If sqlTrans Is Nothing Then
                Dim sqlConn As SqlConnection = New SqlConnection(SqlConnString)
                sqlConn.Open()
                Return sqlConn
            Else
                Return sqlTrans.Connection
            End If
        End Function

        Public Function GetInvoiceLink(invoiceNum As Integer, ByRef sqlTrans As SqlTransaction, ByRef ds As DsSms, ByRef errDesc As String)
            Dim sqlConn As SqlConnection
            Try
                sqlConn = GetSqlConnection(sqlTrans)
                Using sqlCmd As New SqlCommand("GetInvoiceLink")
                    sqlCmd.Connection = sqlConn
                    sqlCmd.Transaction = sqlTrans
                    sqlCmd.CommandType = CommandType.StoredProcedure
                    sqlCmd.Parameters.AddWithValue("@InvoiceNum", invoiceNum)

                    Using sqlDa As New SqlDataAdapter(sqlCmd)
                        sqlDa.Fill(ds, "InvoiceLink")
                    End Using
                End Using

                Return True
            Catch ex As Exception
                errDesc = TraceException(ex.Message)
                Return False
            End Try
        End Function

        Public Function GetInvoiceLinkById(invoiceId As Integer, ByRef sqlTrans As SqlTransaction, ByRef ds As DsSms, ByRef errDesc As String) As Boolean
            Dim sqlConn As SqlConnection
            Try
                sqlConn = GetSqlConnection(sqlTrans)
                Using sqlCmd As New SqlCommand("GetInvoiceLinkById", sqlConn)
                    sqlCmd.Connection = sqlConn
                    sqlCmd.Transaction = sqlTrans
                    sqlCmd.CommandType = CommandType.StoredProcedure
                    sqlCmd.Parameters.AddWithValue("@InvoiceId", invoiceId)

                    Using sqlDa As New SqlDataAdapter(sqlCmd)
                        sqlDa.Fill(ds, "InvoiceLink")
                    End Using
                End Using


                Return True
            Catch ex As Exception
                errDesc = TraceException(ex.Message)
                Return False
            End Try
        End Function

        Public Function AddInvoiceLink(ByVal invoiceRashut As Integer, ByVal invoiceNum As Integer, ByVal invoiceFileSize As Integer,
                                       ByVal baseUrl As String, ByRef sqlTrans As SqlTransaction, ByRef ds As DsSms, ByRef errDesc As String) As Boolean
            Dim invoiceId As Integer
            Dim sqlConn As SqlConnection
            Try
                sqlConn = GetSqlConnection(sqlTrans)
                Using sqlCmd As New SqlCommand("AddInvoiceLink")
                    sqlCmd.Connection = sqlConn
                    sqlCmd.Transaction = sqlTrans
                    sqlCmd.CommandType = CommandType.StoredProcedure
                    sqlCmd.Parameters.AddWithValue("@InvoiceRashut", invoiceRashut)
                    sqlCmd.Parameters.AddWithValue("@InvoiceNum", invoiceNum)
                    sqlCmd.Parameters.AddWithValue("@InvoiceFileSize", invoiceFileSize)
                    sqlCmd.Parameters.AddWithValue("@BaseUrl", baseUrl)

                    sqlCmd.Parameters.Add("@InvoiceId", SqlDbType.Int).Direction = ParameterDirection.Output

                    sqlCmd.ExecuteNonQuery()

                    invoiceId = sqlCmd.Parameters("@InvoiceId").Value
                End Using

                Return GetInvoiceLinkById(invoiceId, sqlTrans, ds, errDesc)
            Catch ex As Exception
                errDesc = TraceException(ex.Message)
                Return False
            End Try
        End Function

        Public Function AddShovarLink(ByVal shovarNum As Integer, ByVal mainTz As Long, ByVal fullUrl As String, ByVal shortUrl As String,
                                      ByRef shovarId As Integer, ByRef sqlTrans As SqlTransaction, ByRef errDesc As String) As Boolean
            Dim sqlConn As SqlConnection
            Try
                sqlConn = GetSqlConnection(sqlTrans)
                Using sqlCmd As New SqlCommand("AddShovarLink")
                    sqlCmd.Connection = sqlConn
                    sqlCmd.Transaction = sqlTrans
                    sqlCmd.CommandType = CommandType.StoredProcedure
                    sqlCmd.Parameters.AddWithValue("@ShovarNum", shovarNum)
                    sqlCmd.Parameters.AddWithValue("@Maintz", mainTz)
                    sqlCmd.Parameters.AddWithValue("@Url", fullUrl)
                    sqlCmd.Parameters.AddWithValue("@ShortUrl", shortUrl)

                    sqlCmd.Parameters.Add("@ShovarId", SqlDbType.Int).Direction = ParameterDirection.Output

                    sqlCmd.ExecuteNonQuery()

                    shovarId = sqlCmd.Parameters("@ShovarId").Value
                End Using

                Return shovarId > 0
            Catch ex As Exception
                errDesc = TraceException(ex.Message)
                Return False
            End Try
        End Function


        Friend Function AddSmsLog(ByVal objId As Integer, ByVal smsType As Integer, ByVal smsSenderText As String, ByVal smsText As String,
                                  ByVal smsLength As Integer, ByVal smsTel As String, ByVal smsStatus As Integer,
                                  ByRef errDesc As String, ByRef sqlTrans As SqlTransaction) As Boolean
            Dim sqlConn As SqlConnection

            If sqlTrans Is Nothing Then
                sqlConn = New SqlConnection(SqlConnString)
                sqlConn.Open()
            Else
                sqlConn = sqlTrans.Connection
            End If

            Try

                Using sqlCmd As SqlCommand = New SqlCommand("AddSmsLog", sqlConn)
                    sqlCmd.Transaction = sqlTrans
                    sqlCmd.CommandType = CommandType.StoredProcedure
                    sqlCmd.Parameters.AddWithValue("@SmsTypeObjId", objId)
                    sqlCmd.Parameters.AddWithValue("@SmsType", smsType)
                    sqlCmd.Parameters.AddWithValue("@SmsSender", smsSenderText)
                    sqlCmd.Parameters.AddWithValue("@SmsText", smsText)
                    sqlCmd.Parameters.AddWithValue("@SmsLength", smsLength)
                    sqlCmd.Parameters.AddWithValue("@SmsTel", smsTel)
                    sqlCmd.Parameters.AddWithValue("@smsStatus", smsStatus)

                    Return sqlCmd.ExecuteNonQuery() = 1
                End Using

            Catch ex As Exception
                errDesc = TraceException(ex.Message)
                Return False
            Finally
                If sqlTrans Is Nothing Then
                    If sqlConn.State <> ConnectionState.Closed Then sqlConn.Close()
                    If sqlConn IsNot Nothing Then sqlConn.Dispose()
                End If
            End Try
        End Function

        Public Function UpdateInvoiceLink(ByRef ds As DsSms, ByRef errDesc As String, ByRef sqlTrans As SqlTransaction) As Boolean
            Dim sqlConn As SqlConnection
            Try
                sqlConn = GetSqlConnection(sqlTrans)
                Using sqlCmd As New SqlCommand("UpdateInvoiceLink", sqlConn)
                    sqlCmd.Connection = sqlConn
                    sqlCmd.Transaction = sqlTrans
                    sqlCmd.CommandType = CommandType.StoredProcedure
                    sqlCmd.Parameters.Add("@InvoiceId", SqlDbType.Int, 4, "InvoiceId")
                    'sqlCmd.Parameters.Add("@FullUrl", SqlDbType.NVarChar, 100, "Url")
                    sqlCmd.Parameters.Add("@ShortUrl", SqlDbType.NVarChar, 60, "ShortUrl")
                    sqlCmd.Parameters.Add("@InvoiceFileSize", SqlDbType.Int, 4, "InvoiceFileSize")

                    Using sqlDa As New SqlDataAdapter(sqlCmd)
                        sqlDa.UpdateCommand = sqlCmd
                        If sqlDa.Update(ds, "InvoiceLink") = 0 Then Throw New Exception("failed updating InvoiceLink table")
                    End Using
                End Using

                Return True
            Catch ex As Exception
                errDesc = TraceException(ex.Message)
                Return False
            End Try
        End Function


        Friend Function GetParamSmsType(ByVal Optional smsType As SmsTypeEnum? = Nothing) As List(Of SmsParams)
            Dim model As List(Of SmsParams)

            Using sqlConn As SqlConnection = New SqlConnection(SqlConnString)

                Using sqlCmd As SqlCommand = New SqlCommand("GetParamSmsType", sqlConn)
                    Dim ds As DsSms
                    Dim da As SqlDataAdapter

                    Try
                        sqlCmd.CommandType = CommandType.StoredProcedure
                        sqlCmd.Parameters.Add("@SmsType", SqlDbType.Int)
                        sqlCmd.Parameters("@SmsType").Value = CInt(smsType)
                        ds = New DsSms()
                        da = New SqlDataAdapter(sqlCmd)
                        da.Fill(ds, "ParamSmsType")
                        model = New List(Of SmsParams)()

                        If ds IsNot Nothing Then
                            model = CreateListFromTable(Of SmsParams)(ds.ParamSmsType)
                        End If

                        Return model
                    Catch ex As Exception
                        Throw New Exception(TraceException(ex.Message))
                    End Try
                End Using
            End Using
        End Function
        Public Function GetShvaHkUnacceptableSmsData(ByVal thum As Integer, ByVal sendnum As Integer, ByRef ds As DsSms) As Boolean


            Using sqlConn As SqlConnection = New SqlConnection(SqlConnString)

                Using sqlCmd As SqlCommand = New SqlCommand("getShvaHkUnacceptableSmsData", sqlConn)

                    Dim da As SqlDataAdapter

                    Try
                        sqlCmd.CommandType = CommandType.StoredProcedure
                        sqlCmd.Parameters.Add("@thum", SqlDbType.Int)
                        sqlCmd.Parameters("@thum").Value = thum
                        sqlCmd.Parameters.Add("@sendnum", SqlDbType.Int)
                        sqlCmd.Parameters("@sendnum").Value = sendnum
                        ds = New DsSms()
                        da = New SqlDataAdapter(sqlCmd)
                        da.Fill(ds, "ShvaHkUnacceptableSmsData")


                        If ds.ShvaHkUnacceptableSmsData.Rows.Count > 0 Then
                            Return True
                        Else
                            Return False
                        End If


                    Catch ex As Exception
                        Throw New Exception(TraceException(ex.Message))
                    End Try
                End Using
            End Using
        End Function
        Public Function FillShvaHkUnacceptableData(ByVal thum As Integer, ByVal sendnum As Integer, ByVal SendUniqueId As String, ByRef ErrDesc As String, ByRef sqlTrans As SqlTransaction) As Boolean

            Dim sqlConn As SqlConnection, affectredRows As Integer

            Try
                Using sqlCmd As New SqlCommand("FillShvaHkUnacceptableData")
                    sqlConn = GetSqlConnection(sqlTrans)
                    sqlCmd.Connection = sqlConn
                    sqlCmd.Transaction = sqlTrans
                    sqlCmd.CommandType = CommandType.StoredProcedure
                    sqlCmd.Parameters.Add("@thum", SqlDbType.Int)
                    sqlCmd.Parameters("@thum").Value = thum
                    sqlCmd.Parameters.Add("@sendnum", SqlDbType.Int)
                    sqlCmd.Parameters("@sendnum").Value = sendnum
                    sqlCmd.Parameters.Add("@SendUniqueId", SqlDbType.NVarChar, 50)
                    sqlCmd.Parameters("@SendUniqueId").Value = SendUniqueId

                    affectredRows = sqlCmd.ExecuteNonQuery()
                    sqlTrans.Commit()
                    Return affectredRows > 0
                End Using
            Catch ex As Exception
                ErrDesc = TraceException(ex.Message)
                Return False
            Finally
                If sqlTrans Is Nothing Then
                    If sqlConn.State <> ConnectionState.Closed Then sqlConn.Close()
                    If sqlConn IsNot Nothing Then sqlConn.Dispose()
                End If
            End Try

        End Function
        Public Function InsertPrtnfs(ByVal mspkod As Integer, ByVal teur As String,
              ByVal pratim As String, ByVal Username As String, ByRef ErrDesc As String, ByRef sqlTrans As SqlTransaction) As Boolean

            Dim sqlConn As SqlConnection

            Try
                Using sqlCmd As New SqlCommand("InsertPrtnfs")
                    sqlConn = GetSqlConnection(sqlTrans)
                    sqlCmd.Connection = sqlConn
                    sqlCmd.Transaction = sqlTrans
                    sqlCmd.CommandType = CommandType.StoredProcedure
                    sqlCmd.Parameters.Add("@mspkod", SqlDbType.Int)
                    sqlCmd.Parameters("@mspkod").Value = mspkod
                    sqlCmd.Parameters.Add("@teur", SqlDbType.NVarChar, 255)
                    sqlCmd.Parameters("@teur").Value = teur
                    sqlCmd.Parameters.Add("@pratim", SqlDbType.NVarChar, 3000)
                    sqlCmd.Parameters("@pratim").Value = pratim
                    sqlCmd.Parameters.Add("@writer", SqlDbType.NVarChar, 50)
                    sqlCmd.Parameters("@writer").Value = Username

                    Return sqlCmd.ExecuteNonQuery() = 1


                End Using
            Catch ex As Exception
                ErrDesc = TraceException(ex.Message)
                Return False
            Finally
                If sqlTrans Is Nothing Then
                    If sqlConn.State <> ConnectionState.Closed Then sqlConn.Close()
                    If sqlConn IsNot Nothing Then sqlConn.Dispose()
                End If
            End Try

        End Function
        Public Function UpdateShvaHkUnacceptableSmsStatus(ds As DsSms, tran As SqlTransaction) As Boolean

            Using sqlConn As SqlConnection = New SqlConnection(SqlConnString)
                Try

                    sqlConn.Open()
                    tran = sqlConn.BeginTransaction()


                    Using sqlCmd As New SqlCommand("UpdateShvaHkUnacceptableSmsStatus", sqlConn, tran)

                        sqlCmd.CommandType = CommandType.StoredProcedure
                        sqlCmd.Parameters.Add("@thum", SqlDbType.Int, 4, "thum")
                        sqlCmd.Parameters.Add("@sendnum", SqlDbType.Int, 4, "sendnum")
                        sqlCmd.Parameters.Add("@mspkod", SqlDbType.Int, 4, "mspkod")
                        sqlCmd.Parameters.Add("@hkhs", SqlDbType.NVarChar, 50, "hkhs")
                        sqlCmd.Parameters.Add("@sent", SqlDbType.Int, 4, "sent")


                        Using da As New SqlDataAdapter(sqlCmd)

                            da.UpdateCommand = sqlCmd
                            If da.Update(ds.ShvaHkUnacceptableSmsData) = ds.ShvaHkUnacceptableSmsData.Count Then
                                tran.Commit()
                                Return True
                            Else
                                Throw New Exception("בעיה בעדכון סטטוס שליחה")
                            End If
                        End Using
                    End Using
                Catch

                End Try
            End Using

        End Function
        Public Function CreateTransaction() As SqlTransaction
            If SqlConnTrans Is Nothing Then
                SqlConnTrans = New SqlConnection(SqlConnString)
                SqlConnTrans.Open()
            Else
                If Not SqlConnTrans.State = ConnectionState.Open Then
                    SqlConnTrans.Open()
                End If
            End If

            Return SqlConnTrans.BeginTransaction()
        End Function

        Public Function CommitTransaction(ByRef sqlTrans As SqlTransaction, ByRef errDesc As String) As Boolean
            Try
                sqlTrans.Commit()
                Return True
            Catch ex As Exception
                errDesc = TraceException(ex.Message)
                Return False
            End Try
        End Function

        ' function that creates a list of an object from the given data table
        Private Shared Function CreateListFromTable(Of T As New)(ByVal tbl As DataTable) As List(Of T)
            ' define return list
            Dim lst As List(Of T) = New List(Of T)()


            ' go through each row
            For Each r As DataRow In tbl.Rows
                ' add to the list
                lst.Add(CreateItemFromRow(Of T)(r))
            Next


            ' return the list
            Return lst
        End Function

        ' function that creates an object from the given data row
        Private Shared Function CreateItemFromRow(Of T As New)(ByVal row As DataRow) As T
            ' create a new object
            Dim item As T = New T()

            ' set the item
            SetItemFromRow(item, row)

            ' return 
            Return item
        End Function

        Private Shared Sub SetItemFromRow(Of T As New)(ByVal item As T, ByVal row As DataRow)
            ' go through each column
            For Each c As DataColumn In row.Table.Columns
                ' find the property for the column
                Dim p As PropertyInfo = item.GetType().GetProperty(c.ColumnName)


                ' if exists, set the value
                If p IsNot Nothing AndAlso row(c) IsNot DBNull.Value Then
                    p.SetValue(item, row(c), Nothing)
                End If
            Next
        End Sub



#Region "SmsSeqSender Functions"

        ''' <summary>
        ''' הכנסת רשומות חדשות למנה מתוך טבלאות shovar ו-msp
        ''' </summary>
        Public Function InsertSmsSeqSenderFromSelect(ByVal manaHovNum As Integer, ByRef errDesc As String) As Integer
            Using sqlConn As SqlConnection = New SqlConnection(SqlConnString)
                Try
                    sqlConn.Open()
                    Using sqlCmd As SqlCommand = New SqlCommand("InsertSmsSeqSenderFromSelect", sqlConn)
                        sqlCmd.CommandType = CommandType.StoredProcedure
                        sqlCmd.Parameters.AddWithValue("@ManaHovNum", manaHovNum)

                        Using reader As SqlDataReader = sqlCmd.ExecuteReader()
                            If reader.Read() Then
                                Return reader.GetInt32(0) ' RowsInserted
                            End If
                        End Using

                        Return 0
                    End Using
                Catch ex As Exception
                    errDesc = TraceException(ex.Message)
                    Return -1
                End Try
            End Using
        End Function

        ''' <summary>
        ''' עדכון סטטוס לרשומה בודדת
        ''' </summary>
        Public Function UpdateSmsSeqSenderStatus(ByVal shovar As Integer, ByVal manaHovNum As Integer,
                                                ByVal newStatus As Integer, ByRef errDesc As String) As Boolean
            Using sqlConn As SqlConnection = New SqlConnection(SqlConnString)
                Try
                    sqlConn.Open()
                    Using sqlCmd As SqlCommand = New SqlCommand("UpdateSmsSeqSenderStatus", sqlConn)
                        sqlCmd.CommandType = CommandType.StoredProcedure
                        sqlCmd.Parameters.AddWithValue("@Shovar", shovar)
                        sqlCmd.Parameters.AddWithValue("@ManaHovNum", manaHovNum)
                        sqlCmd.Parameters.AddWithValue("@NewStatus", newStatus)

                        Return sqlCmd.ExecuteNonQuery() > 0
                    End Using
                Catch ex As Exception
                    errDesc = TraceException(ex.Message)
                    Return False
                End Try
            End Using
        End Function

        ''' <summary>
        ''' שליפת רשומות לפי מנה וסטטוס
        ''' </summary>
        Public Function GetSmsSeqSenderByBatch(ByVal manaHovNum As Integer, ByVal status As Integer?,
                                               ByRef ds As DsSms, ByRef errDesc As String) As Boolean
            Using sqlConn As SqlConnection = New SqlConnection(SqlConnString)
                Try
                    sqlConn.Open()
                    Using sqlCmd As SqlCommand = New SqlCommand("GetSmsSeqSenderByBatch", sqlConn)
                        sqlCmd.CommandType = CommandType.StoredProcedure
                        sqlCmd.Parameters.AddWithValue("@ManaHovNum", manaHovNum)
                        If status.HasValue Then
                            sqlCmd.Parameters.AddWithValue("@Status", status.Value)
                        Else
                            sqlCmd.Parameters.AddWithValue("@Status", DBNull.Value)
                        End If

                        Using sqlDa As New SqlDataAdapter(sqlCmd)
                            ' יצירת DataTable חדש אם לא קיים
                            If Not ds.Tables.Contains("SmsSeqSender") Then
                                ds.Tables.Add("SmsSeqSender")
                            Else
                                ds.Tables("SmsSeqSender").Clear()
                            End If

                            sqlDa.Fill(ds, "SmsSeqSender")
                        End Using
                    End Using

                    Return True
                Catch ex As Exception
                    errDesc = TraceException(ex.Message)
                    Return False
                End Try
            End Using
        End Function

        ''' <summary>
        ''' קבלת מספר רשות מהפרמטרים
        ''' </summary>
        Public Function GetRashutNumber(ByRef errDesc As String) As Integer
            Using sqlConn As SqlConnection = New SqlConnection(SqlConnString)
                Try
                    sqlConn.Open()
                    Dim query As String = "select atareprnum from paramnext"
                    Using cmd As SqlCommand = New SqlCommand(query, sqlConn)
                        Dim result = cmd.ExecuteScalar()
                        If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                            Return CInt(result)
                        End If
                    End Using
                    Return 0 ' ערך ברירת מחדל
                Catch ex As Exception
                    errDesc = TraceException(ex.Message)
                    Return 1
                End Try
            End Using
        End Function

        ''' <summary>
        ''' קבלת Short URL עבור שובר
        ''' </summary>
        Public Function GetShortUrlForShovar(ByVal shovarNum As Integer, ByVal maintz As Long, ByRef errDesc As String) As String
            Using sqlConn As SqlConnection = New SqlConnection(SqlConnString)
                Try
                    sqlConn.Open()
                    Dim query As String = "SELECT TOP 1 ShortUrl FROM ShovarLink WHERE ShovarNum = @ShovarNum AND Maintz = @Maintz AND IsActiveRow = 1 ORDER BY ShovarId DESC"
                    Using cmd As SqlCommand = New SqlCommand(query, sqlConn)
                        cmd.Parameters.AddWithValue("@ShovarNum", shovarNum)
                        cmd.Parameters.AddWithValue("@Maintz", maintz)

                        Dim result = cmd.ExecuteScalar()
                        If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                            Return result.ToString()
                        End If
                    End Using
                    Return ""
                Catch ex As Exception
                    errDesc = TraceException(ex.Message)
                    Return ""
                End Try
            End Using
        End Function

        ''' <summary>
        ''' קבלת ShovarId עבור שובר
        ''' </summary>
        Public Function GetShovarIdForShovar(ByVal shovarNum As Integer, ByVal maintz As Long, ByRef errDesc As String) As Integer
            Using sqlConn As SqlConnection = New SqlConnection(SqlConnString)
                Try
                    sqlConn.Open()
                    Dim query As String = "SELECT TOP 1 ShovarId FROM ShovarLink WHERE ShovarNum = @ShovarNum AND Maintz = @Maintz AND IsActiveRow = 1 ORDER BY ShovarId DESC"
                    Using cmd As SqlCommand = New SqlCommand(query, sqlConn)
                        cmd.Parameters.AddWithValue("@ShovarNum", shovarNum)
                        cmd.Parameters.AddWithValue("@Maintz", maintz)

                        Dim result = cmd.ExecuteScalar()
                        If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                            Return CInt(result)
                        End If
                    End Using
                    Return 0
                Catch ex As Exception
                    errDesc = TraceException(ex.Message)
                    Return 0
                End Try
            End Using
        End Function

#End Region

    End Class
End Namespace
