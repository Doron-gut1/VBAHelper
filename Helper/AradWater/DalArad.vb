Imports System.Data.SqlClient
Imports System.Security.Cryptography
Imports System.Text
Imports FxResources
Imports System.Xml

Namespace AradManager
	Public Class DalArad

		Public Property SqlConnString As String
		Public Property SqlConnTrans As SqlConnection
		Public Sub New(ByVal odbcName As String, ByVal Optional dbUser As String = "", ByVal Optional dbPassword As String = "")
			Try
				Dim odbcConvert As OdbcConverter.OdbcConverter = New OdbcConverter.OdbcConverter()
				SqlConnString = odbcConvert.GetSqlConnectionString(odbcName, dbUser, dbPassword)

			Catch ex As Exception
				Throw New Exception(TraceException(ex.Message))
			End Try

		End Sub

		Public Function GetKaramParam(ByRef ds As DsArad, ByRef errDesc As String) As Boolean
			'get data from karamparam into ds, if rows count <> 1 set appropriate error description in errDesc, and return false
			ds.KaramParam.Clear()
			Try
				Using sqlConn As SqlConnection = New SqlConnection(SqlConnString)
					sqlConn.Open()
					Dim cmd As SqlCommand = New SqlCommand("GetKaramParam", sqlConn)
					cmd.CommandType = CommandType.StoredProcedure
					Dim adapter As New SqlDataAdapter(cmd)
					adapter.SelectCommand.CommandTimeout = 300

					adapter.Fill(ds, "KaramParam")
				End Using

				If ds.KaramParam.Count = 1 Then
					Return True
				Else
					Throw New Exception("KaramParam table should have 1 row")
				End If

			Catch ex As Exception
				errDesc = TraceException(ex.Message)
				Return False
			End Try
		End Function

		Public Function GetSettlements(ByRef ds As DsArad, ByRef errDesc As String, ByVal Optional isuvkod As Integer = Nothing) As Boolean
			'get data from karamparam into ds, if rows count <> 1 set appropriate error description in errDesc, and return false
			ds.AradSettlements.Clear()
			Try
				Using sqlConn As SqlConnection = New SqlConnection(SqlConnString)
					sqlConn.Open()
					Dim query As String = $"SELECT [SettlementID] ,[DiameterGroup] ,[ModuleGroup] From [AradSettlements]{IIf(isuvkod <> Nothing, $" Where [isuvkod] = @isuvkod", String.Empty)}"
					Dim cmd As SqlCommand = New SqlCommand(query, sqlConn)
					cmd.CommandType = CommandType.StoredProcedure
					If isuvkod <> Nothing Then
						cmd.Parameters.AddWithValue("@isuvkod", isuvkod)
					End If
					Dim adapter As New SqlDataAdapter(cmd)
					adapter.SelectCommand.CommandTimeout = 300

					adapter.Fill(ds, "AradSettlements")
				End Using

				If ds.AradSettlements.Count > 0 Then
					Return True
				Else
					Return False
				End If

			Catch ex As Exception
				errDesc = TraceException(ex.Message)
				Return False
			End Try
		End Function

		Public Function GetSettlement(ByRef ds As DsArad, ByRef errDesc As String, ByVal isuvkod As Integer) As DsArad.AradSettlementsRow
			'get data from karamparam into ds, if rows count <> 1 set appropriate error description in errDesc, and return false
			ds.AradSettlements.Clear()
			Try
				Using sqlConn As SqlConnection = New SqlConnection(SqlConnString)
					sqlConn.Open()
					Dim query As String = "SELECT [SettlementID] ,[DiameterGroup] ,[ModuleGroup] From [AradSettlements] Where [isuvkod] = @isuvkod"
					Dim cmd As SqlCommand = New SqlCommand(query, sqlConn)
					cmd.CommandType = CommandType.StoredProcedure
					If isuvkod <> Nothing Then
						cmd.Parameters.AddWithValue("@isuvkod", isuvkod)
					End If
					Dim adapter As New SqlDataAdapter(cmd)
					adapter.SelectCommand.CommandTimeout = 300

					adapter.Fill(ds, "AradSettlements")
				End Using

				If ds.AradSettlements.Count = 1 Then
					Return ds.AradSettlements.FirstOrDefault()
				Else
					Return Nothing
				End If

			Catch ex As Exception
				errDesc = TraceException(ex.Message)
				Return Nothing
			End Try
		End Function

		Public Function GetAradWaterAvData(ByRef ds As DsArad, ByRef errDesc As String) As Boolean
			Try
				Using sqlConn As SqlConnection = New SqlConnection(SqlConnString)
					sqlConn.Open()
					Dim cmd As SqlCommand = New SqlCommand("GetAradAvData", sqlConn)
					cmd.CommandType = CommandType.StoredProcedure
					Using sqlDA As New SqlDataAdapter(cmd)
						sqlDA.Fill(ds, "AvData")
					End Using
				End Using

			Catch ex As Exception
				AddDbErrors(SqlConnString, TraceException(ex.Message))
				errDesc = TraceException(ex.Message)
				Return False
			End Try

			Return True
		End Function

		Public Function GetNetworkDir() As String
			Dim netDir As String
			Try
				Using sqlConn As SqlConnection = New SqlConnection(SqlConnString)
					sqlConn.Open()
					Using cmd As SqlCommand = New SqlCommand("SELECT dbo.GetPicPath()", sqlConn)
						netDir = cmd.ExecuteScalar
					End Using

				End Using

				Return netDir

			Catch ex As Exception
				Throw ex
			End Try
		End Function

		Public Function GetThisMoz() As Integer
			Dim thisMoz As Integer
			Try
				Using sqlConn As SqlConnection = New SqlConnection(SqlConnString)
					sqlConn.Open()
					Using cmd As SqlCommand = New SqlCommand("SELECT dbo.thismoz()", sqlConn)
						thisMoz = cmd.ExecuteScalar
					End Using

				End Using

				Return thisMoz

			Catch ex As Exception
				Throw ex
			End Try
		End Function

		Public Function GetApiGwParams(ByRef ds As DsArad, ByRef errDesc As String) As Boolean
			ds.ParamApiGateway.Clear()

			Try
				Using sqlConn As SqlConnection = New SqlConnection(SqlConnString)
					'sqlConn.Open()
					Dim cmd As SqlCommand = New SqlCommand("GetParamApiGw", sqlConn)
					cmd.CommandType = CommandType.StoredProcedure
					Dim adapter As New SqlDataAdapter(cmd)
					adapter.SelectCommand.CommandTimeout = 300

					adapter.Fill(ds, "ParamApiGateway")
				End Using

				If ds.ParamApiGateway.Count = 1 Then
					Return True
				Else
					Throw New Exception("ParamApiGateway table should have 1 row")
				End If

			Catch ex As Exception
				errDesc = TraceException(ex.Message)
				Return False
			End Try
		End Function

		Public Function AddWtrFileControl(ByVal processDate As Integer, ByVal wtrCompany As Short, ByVal wtrRashut As Integer,
				ByVal wtrFileDir As String, ByVal wtrFileName As String, ByVal wtrFileSize As Long, ByVal wtrFileRowsCount As Integer,
				ByVal wtrFileProcessId As Short, ByVal wtrFileProcessStatus As Short, ByVal processStarted As DateTime, ByVal userName As String,
				ByRef ds As DsArad, ByRef errDesc As String, ByRef wtrFileIndex As Integer) As Boolean

			'Dim wtrFileIndex As Integer

			Try
				Using sqlConn As SqlConnection = New SqlConnection(SqlConnString)
					sqlConn.Open()

					Using sqlCmd As SqlCommand = New SqlCommand("AddWtrFileControl", sqlConn)
						sqlCmd.CommandType = CommandType.StoredProcedure
						sqlCmd.Parameters.AddWithValue("@ProcessDate", processDate)
						sqlCmd.Parameters.AddWithValue("@WtrCompany", wtrCompany)
						sqlCmd.Parameters.AddWithValue("@WtrRashut", wtrRashut)
						sqlCmd.Parameters.AddWithValue("@WtrFileDir", wtrFileDir)
						sqlCmd.Parameters.AddWithValue("@WtrFileName", wtrFileName)
						sqlCmd.Parameters.AddWithValue("@WtrFileSize", wtrFileSize)
						sqlCmd.Parameters.AddWithValue("@WtrFileRowsCount", wtrFileRowsCount)
						sqlCmd.Parameters.AddWithValue("@WtrFileProcessId", wtrFileProcessId)
						sqlCmd.Parameters.AddWithValue("@WtrFileProcessStatus", wtrFileProcessStatus)
						sqlCmd.Parameters.AddWithValue("@ProcessStarted", processStarted)
						sqlCmd.Parameters.AddWithValue("@ProcUserName", userName)
						sqlCmd.Parameters.Add("@WtrFileIndex", SqlDbType.Int).Direction = ParameterDirection.Output

						sqlCmd.ExecuteNonQuery()

						wtrFileIndex = sqlCmd.Parameters("@WtrFileIndex").Value
					End Using
				End Using

				Return GetWtrFileControlByIndex(wtrFileIndex, ds, errDesc)

			Catch ex As Exception
				errDesc = TraceException(ex.Message)
				Return False
			End Try
		End Function

		Public Function GetWtrFileControlByIndex(ByVal wtrFileIndex As Integer, ByRef ds As DsArad, ByRef errDesc As String) As Boolean
			Try
				Using sqlConn As SqlConnection = New SqlConnection(SqlConnString)
					Using sqlCmd As New SqlCommand("GetWtrFileControlByIndex", sqlConn)
						sqlCmd.Connection = sqlConn
						sqlCmd.CommandType = CommandType.StoredProcedure
						sqlCmd.Parameters.AddWithValue("@wtrFileIndex", wtrFileIndex)

						Using sqlDa As New SqlDataAdapter(sqlCmd)
							sqlDa.Fill(ds, "WtrFileControl")
						End Using
					End Using
				End Using

				Return True
			Catch ex As Exception
				errDesc = TraceException(ex.Message)
				Return False
			End Try
		End Function
		Public Function UpdateWtrFileControl(ByRef ds As DsArad, ByRef errDesc As String) As Boolean
			Try
				Using sqlConn As SqlConnection = New SqlConnection(SqlConnString)
					Using sqlCmd As New SqlCommand("UpdateWtrFileControl", sqlConn)
						sqlCmd.Connection = sqlConn
						sqlCmd.CommandType = CommandType.StoredProcedure
						sqlCmd.Parameters.Add("@WtrFileIndex", SqlDbType.Int, 4, "WtrFileIndex")
						sqlCmd.Parameters.Add("@WtrFileTaskRef", SqlDbType.Int, 4, "WtrFileTaskRef")
						sqlCmd.Parameters.Add("@WtrFileProcessId", SqlDbType.SmallInt, 2, "WtrFileProcessId")
						sqlCmd.Parameters.Add("@WtrFileProcessStatus", SqlDbType.SmallInt, 2, "WtrFileProcessStatus")
						sqlCmd.Parameters.Add("@WtrFileDir", SqlDbType.NVarChar, 255, "WtrFileDir")
						sqlCmd.Parameters.Add("@WtrFileName", SqlDbType.NVarChar, 255, "WtrFileName")
						sqlCmd.Parameters.Add("@WtrFileSize", SqlDbType.BigInt, 8, "WtrFileSize")
						sqlCmd.Parameters.Add("@WtrFileRowsCount", SqlDbType.Int, 4, "WtrFileRowsCount")
						sqlCmd.Parameters.Add("@ProcessStarted", SqlDbType.DateTime, 8, "ProcessStarted")
						sqlCmd.Parameters.Add("@ProcessFinished", SqlDbType.DateTime, 8, "ProcessFinished")
						sqlCmd.Parameters.Add("@ProcErrorDesc", SqlDbType.NVarChar, 1000, "ProcErrorDesc")
						sqlCmd.Parameters.Add("@ProcUserName", SqlDbType.NVarChar, 50, "ProcUserName")

						Using sqlDa As New SqlDataAdapter(sqlCmd)
							sqlDa.UpdateCommand = sqlCmd
							sqlDa.Update(ds, "WtrFileControl")
						End Using
					End Using
				End Using

				Return True
			Catch ex As Exception
                Throw New Exception(TraceException(ex.Message))
            End Try
		End Function

		Public Function GetWtrFileControlToRequestProcess(ByRef ds As DsArad, ByRef errDesc As String) As Boolean
			Try
				Using sqlConn As SqlConnection = New SqlConnection(SqlConnString)
					Using sqlCmd As New SqlCommand("GetWtrFileControlToRequestProcess", sqlConn)
						sqlCmd.Connection = sqlConn
						sqlCmd.CommandType = CommandType.StoredProcedure

						Using sqlDa As New SqlDataAdapter(sqlCmd)
							sqlDa.Fill(ds, "WtrFileControl")
						End Using
					End Using
				End Using

				Return True
			Catch ex As Exception
				errDesc = TraceException(ex.Message)
				Return False
			End Try
		End Function

		Public Function GetLastOpenMnt() As Integer
			Dim mnt As Integer
			Try
				Using sqlConn As SqlConnection = New SqlConnection(SqlConnString)
					sqlConn.Open()
					Using cmd As SqlCommand = New SqlCommand("SELECT dbo.LastOpen1010()", sqlConn)
						mnt = cmd.ExecuteScalar
					End Using

				End Using

				Return mnt
			Catch ex As Exception
				Throw ex
			End Try
		End Function

		Public Function UpdateReads(ByRef ds As DsArad, ByRef errDesc As String) As Boolean
			Dim mnt As Integer
			Dim totalRowsAffected As Integer
			Try
				mnt = GetLastOpenMnt()

				Using sqlConn As SqlConnection = New SqlConnection(SqlConnString)
					Using sqlCmd As New SqlCommand("UpdateKriotWater", sqlConn)
						sqlCmd.Connection = sqlConn
						sqlCmd.CommandType = CommandType.StoredProcedure
						sqlCmd.Parameters.Add("@zihuyWater", SqlDbType.NVarChar, 255, "ZihuyWater")
						sqlCmd.Parameters.Add("@kria", SqlDbType.Int, 4, "CurrRead")
						sqlCmd.Parameters.Add("@kriadt", SqlDbType.DateTime, 4, "ReadDate")
						sqlCmd.Parameters.Add("@lastkria", SqlDbType.Int, 4, "LastRead")
						sqlCmd.Parameters.AddWithValue("@ZihuyIsHskod", 0)
						sqlCmd.Parameters.AddWithValue("@mnt", mnt)

						Using sqlDa As New SqlDataAdapter(sqlCmd)
							sqlDa.UpdateCommand = sqlCmd
							sqlDa.InsertCommand = sqlCmd
							totalRowsAffected = sqlDa.Update(ds, "ReadsData")
						End Using
					End Using
				End Using

				Return True
			Catch ex As Exception
				errDesc = TraceException(ex.Message)
				Return False
			End Try
		End Function

	End Class

End Namespace
