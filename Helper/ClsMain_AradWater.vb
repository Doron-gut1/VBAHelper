Imports System.Collections.Specialized
Imports System.IO
Imports System.Runtime.InteropServices
'Imports System.Xml.Serialization
Imports Epr.AradWaterStructures.Response
'Imports System.Text.Json
'Imports System.Text.Json.Serialization
Imports Epr.AradWaterStructures.Request
Imports System.Reflection
Imports System.Text
Imports System.Xml
Imports System.Runtime.InteropServices.ComTypes
Imports Newtonsoft.Json
Imports System.Security.Policy

Partial Class ClsMain
	Public Enum WtrFileProcIdEnum
		AvFileUploaded = 1
		AvFileRequestForProcess = 2
		AvFileProcessFinished = 3
		ReadsFileRequest = 11
		ReadsFileStatus = 12
	End Enum

	Public Enum WtrFileProcStatusEnum
		None = 0
		Failed = -1
		Canceled = -2
		Working = 9
		Finished = 1
	End Enum

	Function BuildAvFile(ByVal odbcName As String, ByVal Optional isuvkod As Integer = Nothing, ByVal Optional user As String = Nothing,
					  ByVal Optional password As String = Nothing) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

		Dim avFileStatus, rowsCount As Integer
		Dim errDesc As String = ""
		Dim StatusCode As Integer = 0, StatusText As String = ""
		Dim errObj As ClsError = Nothing
		Dim dal As AradManager.DalArad = Nothing
		Dim fileData As String = ""
		Dim aradApi As AradManager.ApiArad
		Dim queryString As NameValueCollection = System.Web.HttpUtility.ParseQueryString(String.Empty)
		Dim outputFileName As String
		Dim outputFileDir As String
		Dim startTime As DateTime
		Dim fi As FileInfo
		Dim ds As New DsArad
		Dim res As Object
		Dim aradRes As AradResponse
		Dim control As DsArad.WtrFileControlRow = Nothing
		Try
			avFileStatus = 0
			errDesc = ""
			dal = New AradManager.DalArad(odbcName)
			aradApi = New AradManager.ApiArad(odbcName, user, password)
			startTime = DateTime.Now
			rowsCount = 0
			outputFileDir = Path.Combine(dal.GetNetworkDir(), aradApi.KovetzAvFileBackupDir, DateTime.Now.ToString("yyyyMM"))

			If Not Directory.Exists(outputFileDir) Then
				Directory.CreateDirectory(outputFileDir)
			End If

			If dal.GetSettlements(ds, errDesc, isuvkod) Then
				For Each row As DsArad.AradSettlementsRow In ds.AradSettlements.Rows
					If dal.GetAradWaterAvData(ds, errDesc) Then
						outputFileName = $"Import_{row.SettlementID}_MASTER_FILE_{DateTime.Now.ToString("ddMMyyyyHHmmss")}.txt"
						'File.WriteAllText(Path.Combine(outputFileDir, outputFileName), fileData, Encoding.UTF8)

						If Not SaveAvDataToFile(Path.Combine(outputFileDir, outputFileName), ds, errDesc) Then Throw New Exception(errDesc)

						fi = New FileInfo(Path.Combine(outputFileDir, outputFileName))
						Dim rowID As Integer
						If Not dal.AddWtrFileControl(DateTime.Now.ToString("yyyyMMdd"), 4, row.SettlementID, outputFileDir, outputFileName, fi.Length,
										   rowsCount, WtrFileProcIdEnum.AvFileUploaded, WtrFileProcStatusEnum.Working, startTime, Environment.UserName, ds, errDesc, rowID) Then
							errObj = New ClsError(0, $"Message: {errDesc}", errDesc)
						Else
							res = aradApi.UploadAvFile(outputFileDir, outputFileName, row.SettlementID)
							aradRes = JsonConvert.DeserializeObject(Of AradResponse)(res)

							control = ds.WtrFileControl.First(Function(x) x.WtrFileIndex = rowID)
							If aradRes.Succeeded Then
								control.WtrFileProcessStatus = WtrFileProcStatusEnum.Finished
							Else
								'ds.WtrFileControl(0).ProcErrorDesc = aradRes.ErrorDescription
								'ds.WtrFileControl(0).WtrFileProcessStatus = WtrFileProcStatusEnum.Failed
								Throw New Exception(aradRes.ErrorDescription)
							End If

							control.ProcessFinished = DateTime.Now
							dal.UpdateWtrFileControl(ds, errDesc)

							If aradRes.Succeeded Then
								If Not AvFileProcess(aradApi, dal, ds, errDesc, rowID) Then
									Throw New Exception(errDesc)
								End If
							End If
						End If

					Else
						errObj = New ClsError(0, $"Message: {errDesc}", errDesc)
					End If
				Next
				StatusCode = 1
			Else
				errObj = New ClsError(0, $"Message: {errDesc}", errDesc)
			End If


		Catch ex As Exception
			AddDbErrors(dal.SqlConnString, ex.Message)
			If dal IsNot Nothing And control IsNot Nothing Then
				control.ProcessFinished = DateTime.Now
				control.WtrFileProcessStatus = WtrFileProcStatusEnum.Failed
				control.ProcErrorDesc = If(ex.Message.Length > 1000, ex.Message.Substring(0, 1000), ex.Message)
				dal.UpdateWtrFileControl(ds, errDesc)
			End If

			Dim InnerEx As Exception = ExtractMustInnerException(ex)
			StatusCode = 0
			errObj = New ClsError(0, $"Message: {InnerEx.Message}{vbCrLf}{vbCrLf}Stack-Trace: {InnerEx.StackTrace}", InnerEx.Message)
		End Try

		Return New ClsReturn(StatusCode, StatusText, Nothing, errObj)
	End Function

	'Function StartProcessAvFile(ByVal odbcName As String, ByVal Optional user As String = Nothing,
	'				  ByVal Optional password As String = Nothing) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn
	'	Dim dal As AradManager.DalArad = Nothing
	'	Dim aradApi As AradManager.ApiArad
	'	Dim ds As New DsArad
	'	Dim errDesc As String = ""
	'	Dim errObj As ClsError = Nothing
	'	Dim dsToUpdate As New DsArad
	'	Dim aradRes As AradAvFileProcessResponse
	'	Try
	'		dal = New AradManager.DalArad(odbcName)
	'		aradApi = New AradManager.ApiArad(odbcName, user, password)

	'		If Not dal.GetWtrFileControlToRequestProcess(ds, errDesc) Then
	'			Throw New Exception(errDesc)
	'		End If

	'		For i As Integer = 0 To ds.WtrFileControl.Count - 1
	'			'Try
	'			dsToUpdate.WtrFileControl.Clear()
	'			dsToUpdate.WtrFileControl.ImportRow(ds.WtrFileControl(i))

	'			AvFileProcess(aradApi, dal, dsToUpdate, errDesc)

	'			'dsToUpdate.WtrFileControl(0).WtrFileProcessId = WtrFileProcIdEnum.RequestForProcess
	'			'dsToUpdate.WtrFileControl(0).WtrFileProcessStatus = WtrFileProcStatusEnum.Working
	'			'dsToUpdate.WtrFileControl(0).ProcessStarted = DateTime.Now
	'			'dal.UpdateWtrFileControl(dsToUpdate, errDesc)

	'			'queryString.Clear()
	'			'queryString.Add("FileName", ds.WtrFileControl(i).WtrFileName)
	'			'res = aradApi.ApiAradMethodCall("AvFileProcess", queryString.ToString())

	'			'aradRes = JsonConvert.DeserializeObject(Of AradAvFileProcessResponse)(res)

	'			'If Not aradRes.Succeeded Then
	'			'	Throw New Exception(aradRes.ErrorDescription)
	'			'End If

	'			'dsToUpdate.WtrFileControl(0).WtrFileTaskRef = aradRes.FileTaskRef
	'			'dsToUpdate.WtrFileControl(0).WtrFileProcessStatus = WtrFileProcStatusEnum.Finished
	'			'dsToUpdate.WtrFileControl(0).ProcessFinished = DateTime.Now
	'			'dal.UpdateWtrFileControl(dsToUpdate, errDesc)
	'			'Catch ex As Exception
	'			'If dal IsNot Nothing And dsToUpdate.WtrFileControl.Count = 1 Then
	'			'	dsToUpdate.WtrFileControl(i).ProcessFinished = DateTime.Now
	'			'	dsToUpdate.WtrFileControl(i).WtrFileProcessStatus = WtrFileProcStatusEnum.Failed
	'			'	dsToUpdate.WtrFileControl(i).ProcErrorDesc = ex.Message
	'			'	dal.UpdateWtrFileControl(dsToUpdate, errDesc)
	'			'End If
	'			'End Try
	'		Next
	'	Catch ex As Exception
	'		Dim InnerEx As Exception = ExtractMustInnerException(ex)
	'		errObj = New ClsError(0, $"Message: {InnerEx.Message}{vbCrLf}{vbCrLf}Stack-Trace: {InnerEx.StackTrace}", InnerEx.Message)
	'	End Try
	'End Function

	Function SaveAvDataToFile(ByVal fileName As String, ByVal ds As DsArad, ByRef errDesc As String) As Boolean
		Dim fileData As New StringBuilder 'the most efficient way to write to file is appending to stringbuilder and then writing at once
		Try
			For currRow As Integer = 0 To ds.AvData.Count - 1
				If currRow = 0 Then
					fileData.Append(ds.AvData.Columns("RowType").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("MeterID").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("MeterSN").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("MunicipalCode").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("PropertyID").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("City").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("Street").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("Number").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("Floor").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("Apartment").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("Zip").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("Remarks").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("ConsumerID").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("First_Name").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("Last_Name").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("ConsumerHomePhone").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("ConsumerWorkPhone").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("ConsumerMobilePhone").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("Telephone").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("ConsumerEmail").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("Diameter").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("InstallDate").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("Routing").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("Longitude").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("Latitude").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("Zone").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("SubZone").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("ConsuptionType").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("MeterCategory").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("ParticipationPCT").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("ConsumerCity").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("ConsumerStreet").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("ConsumerNumber").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("ConsumerApartment").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("WaterTypeID").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("WaterTypeDesc").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("ParentMeterID").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("MainMeterID").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("Residents").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("ParallelID").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("TransponderID").Caption & vbTab)
					fileData.Append(ds.AvData.Columns("Model").Caption & vbCrLf)
				End If

				fileData.Append(ds.AvData(currRow).RowType & vbTab)
				fileData.Append(ds.AvData(currRow).MeterID.Replace(" ", "").PadLeft(12, "0") & vbTab)
				fileData.Append(TruncateString(ds.AvData(currRow).MeterSN, 16) & vbTab)
				fileData.Append(ds.AvData(currRow).MunicipalCode & vbTab)
				fileData.Append(TruncateString(ds.AvData(currRow).PropertyID, 15) & vbTab)
				fileData.Append(TruncateString(ds.AvData(currRow).City, 50) & vbTab)
				fileData.Append(TruncateString(ds.AvData(currRow).Street, 50) & vbTab)
				fileData.Append(TruncateString(ds.AvData(currRow).Number, 7) & vbTab)
				fileData.Append(TruncateString(If(ds.AvData(currRow).IsFloorNull, "", ds.AvData(currRow).Floor), 5) & vbTab)
				fileData.Append(TruncateString(If(ds.AvData(currRow).IsApartmentNull, "", ds.AvData(currRow).Apartment), 7) & vbTab)
				fileData.Append(If(ds.AvData(currRow).IsZipNull, "", ds.AvData(currRow).Zip) & vbTab)
				fileData.Append(TruncateString(ds.AvData(currRow).Remarks, 100) & vbTab)
				fileData.Append(If(ds.AvData(currRow).IsConsumerIDNull, "", ds.AvData(currRow).ConsumerID) & vbTab)
				fileData.Append(TruncateString(ds.AvData(currRow).First_Name, 20) & vbTab)
				fileData.Append(TruncateString(ds.AvData(currRow).Last_Name, 20) & vbTab)
				fileData.Append(TruncateString(ds.AvData(currRow).ConsumerHomePhone, 20) & vbTab)
				fileData.Append(TruncateString(ds.AvData(currRow).ConsumerWorkPhone, 20) & vbTab)
				fileData.Append(TruncateString(ds.AvData(currRow).ConsumerMobilePhone, 20) & vbTab)
				fileData.Append(TruncateString(ds.AvData(currRow).Telephone, 20) & vbTab)
				fileData.Append(TruncateString(ds.AvData(currRow).ConsumerEmail, 40) & vbTab)
				fileData.Append(TruncateString(ds.AvData(currRow).Diameter, 50) & vbTab)
				fileData.Append(If(ds.AvData(currRow).IsInstallDateNull, "", ds.AvData(currRow).InstallDate.ToString("dd/MM/yyyy")) & vbTab)
				fileData.Append(If(ds.AvData(currRow).IsRoutingNull, "", ds.AvData(currRow).Routing) & vbTab)
				fileData.Append(If(ds.AvData(currRow).IsLongitudeNull, "", ds.AvData(currRow).Longitude) & vbTab)
				fileData.Append(If(ds.AvData(currRow).IsLatitudeNull, "", ds.AvData(currRow).Latitude) & vbTab)
				fileData.Append(If(ds.AvData(currRow).IsZoneNull, "", ds.AvData(currRow).Zone) & vbTab)
				fileData.Append(If(ds.AvData(currRow).IsSubZoneNull, "", ds.AvData(currRow).SubZone) & vbTab)
				fileData.Append(If(ds.AvData(currRow).IsConsuptionTypeNull, "", ds.AvData(currRow).ConsuptionType) & vbTab)
				fileData.Append(If(ds.AvData(currRow).IsMeterCategoryNull, "", ds.AvData(currRow).MeterCategory) & vbTab)
				fileData.Append(If(ds.AvData(currRow).IsParticipationPCTNull, "", ds.AvData(currRow).ParticipationPCT) & vbTab)
				fileData.Append(TruncateString(ds.AvData(currRow).ConsumerCity, 50) & vbTab)
				fileData.Append(TruncateString(ds.AvData(currRow).ConsumerStreet, 50) & vbTab)
				fileData.Append(If(ds.AvData(currRow).IsConsumerNumberNull, "", ds.AvData(currRow).ConsumerNumber) & vbTab)
				fileData.Append(If(ds.AvData(currRow).IsConsumerApartmentNull, "", ds.AvData(currRow).ConsumerApartment) & vbTab)
				fileData.Append(If(ds.AvData(currRow).IsWaterTypeIDNull, "", ds.AvData(currRow).WaterTypeID) & vbTab)
				fileData.Append(TruncateString(ds.AvData(currRow).WaterTypeDesc, 50) & vbTab)
				fileData.Append(TruncateString(ds.AvData(currRow).ParentMeterID, 15) & vbTab)
				fileData.Append(TruncateString(ds.AvData(currRow).MainMeterID, 15) & vbTab)
				fileData.Append(If(ds.AvData(currRow).IsResidentsNull, "", ds.AvData(currRow).Residents) & vbTab)
				fileData.Append(If(ds.AvData(currRow).IsParallelIDNull, "", ds.AvData(currRow).ParallelID) & vbTab)
				fileData.Append(TruncateString(ds.AvData(currRow).TransponderID, 16) & vbTab)
				fileData.Append(If(ds.AvData(currRow).IsModelNull, "", ds.AvData(currRow).Model) & vbCrLf)
			Next
			File.WriteAllText(fileName, fileData.ToString(), Encoding.UTF8)
			Return True
		Catch ex As Exception
			errDesc = ex.Message
			Return False
		End Try
	End Function



	Private Function AvFileProcess(ByVal aradApi As AradManager.ApiArad, ByVal dal As AradManager.DalArad,
								   ByRef dsToUpdate As DsArad, ByRef errDesc As String, ByVal rowID As Integer) As Boolean
		Dim queryString As NameValueCollection = System.Web.HttpUtility.ParseQueryString(String.Empty)
		Dim res As Object
		Dim aradRes As AradAvFileProcessResponse
		Dim control As DsArad.WtrFileControlRow = dsToUpdate.WtrFileControl.First(Function(x) x.WtrFileIndex = rowID)
		Try
			control.WtrFileProcessId = WtrFileProcIdEnum.AvFileRequestForProcess
			control.WtrFileProcessStatus = WtrFileProcStatusEnum.Working
			control.ProcessStarted = DateTime.Now
			dal.UpdateWtrFileControl(dsToUpdate, errDesc)

			queryString.Clear()
			queryString.Add("FileName", dsToUpdate.WtrFileControl(0).WtrFileName)
			res = aradApi.ApiAradMethodCall("AvFileProcess", queryString.ToString(), control.WtrRashut)

			aradRes = JsonConvert.DeserializeObject(Of AradAvFileProcessResponse)(res)

			If Not aradRes.Succeeded Then
				Throw New Exception(aradRes.ErrorDescription)
			End If

			control.WtrFileTaskRef = aradRes.FileTaskRef
			control.WtrFileProcessStatus = WtrFileProcStatusEnum.Finished
			control.ProcessFinished = DateTime.Now
			dal.UpdateWtrFileControl(dsToUpdate, errDesc)

			Return True
		Catch ex As Exception
			AddDbErrors(dal.SqlConnString, ex.Message)
			If dal IsNot Nothing And dsToUpdate.WtrFileControl.Count = 1 Then
				control.ProcessFinished = DateTime.Now
				control.WtrFileProcessStatus = WtrFileProcStatusEnum.Failed
				control.ProcErrorDesc = If(ex.Message.Length > 1000, ex.Message.Substring(0, 1000), ex.Message)
				dal.UpdateWtrFileControl(dsToUpdate, errDesc)
			End If
			errDesc = ex.Message
			Return False
		End Try
	End Function

	Function GetOnlineRead(ByVal moneId As String, ByVal readDate As String, ByVal odbcName As String, ByVal isuvkod As Integer, ByVal Optional user As String = Nothing,
					  ByVal Optional password As String = Nothing) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn
		Dim errObj As ClsError = Nothing
		Dim StatusCode As Integer = 0, StatusText As String = ""
		Dim queryString As NameValueCollection = System.Web.HttpUtility.ParseQueryString(String.Empty)
		Dim aradApi As AradManager.ApiArad
		Dim res As Object
		Dim aradRes As AradReadingsResponse
		Dim ds As New DsArad
		Dim errDesc As String = ""
		Dim dal As AradManager.DalArad = Nothing
		Dim parsedDate As DateTime
		Try
			aradApi = New AradManager.ApiArad(odbcName, user, password)
			dal = New AradManager.DalArad(odbcName)

			queryString.Add("moneId", moneId)
			If readDate IsNot Nothing AndAlso DateTime.TryParseExact(readDate, "yyyy-MM-dd", New Globalization.CultureInfo("en-US"),
							  Globalization.DateTimeStyles.AdjustToUniversal, parsedDate) Then
				queryString.Add("readDate", readDate)
			End If
			Dim settlement As DsArad.AradSettlementsRow = dal.GetSettlement(ds, errDesc, isuvkod)
			res = aradApi.ApiAradMethodCall("OnlineRead", queryString.ToString(), settlement.SettlementID)
			aradRes = JsonConvert.DeserializeObject(Of AradReadingsResponse)(res)

			If aradRes.Results.ReadingInfo IsNot Nothing AndAlso
				aradRes.Results.ReadingInfo.Count = 1 AndAlso aradRes.Results.ReadingInfo(0).errorid = 0 Then

				Dim newReadRow As DsArad.ReadsDataRow
				newReadRow = ds.ReadsData.NewReadsDataRow

				newReadRow.ZihuyWater = moneId
				newReadRow.LastRead = -1
				newReadRow.CurrRead = aradRes.Results.ReadingInfo(0).read
				newReadRow.ReadDate = DateTime.ParseExact(aradRes.Results.ReadingInfo(0).readtime, "yyyy-MM-dd", Nothing)
				ds.ReadsData.AddReadsDataRow(newReadRow)

				If Not dal.UpdateReads(ds, errDesc) Then
					Throw New Exception($"GetOnlineRead exception - {errDesc}")
				End If
			Else
				Throw New Exception($"{aradRes.Results.ReadingInfo(0).errorid} - {aradRes.Results.ReadingInfo(0).errordescription}")
			End If
			StatusCode = 1
		Catch ex As Exception
			Try
				AddDbErrors(dal.SqlConnString, TraceException(ex.Message))
			Catch

			End Try
			Dim InnerEx As Exception = ExtractMustInnerException(ex)
			errObj = New ClsError(0, $"Message: {InnerEx.Message}{vbCrLf}{vbCrLf}Stack-Trace: {InnerEx.StackTrace}", InnerEx.Message)
		End Try

		Return New ClsReturn(StatusCode, StatusText, aradRes, errObj)
	End Function

	Function CrossMoneWtr(ByVal moneId1 As String, ByVal moneId2 As String, ByVal odbcName As String, ByVal isuvkod As Integer, ByVal Optional user As String = Nothing,
				  ByVal Optional password As String = Nothing) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn
		Dim errObj As ClsError = Nothing
		Dim StatusCode As Integer = 0, StatusText As String = ""
		Dim queryString As NameValueCollection = System.Web.HttpUtility.ParseQueryString(String.Empty)
		Dim aradApi As AradManager.ApiArad
		Dim aradRes As AradResponse
		Dim dal As AradManager.DalArad = Nothing
		Dim res As Object
		Dim ds As New DsArad
		Dim errDesc As String = String.Empty
		Try

			aradApi = New AradManager.ApiArad(odbcName, user, password)
			dal = New AradManager.DalArad(odbcName)

			queryString.Add("moneId1", moneId1)
			queryString.Add("moneId2", moneId2)
			Dim settlement As DsArad.AradSettlementsRow = dal.GetSettlement(ds, errDesc, isuvkod)
			res = aradApi.ApiAradMethodCall("CrossMoneWtr", queryString.ToString(), settlement.SettlementID)
			aradRes = JsonConvert.DeserializeObject(Of AradResponse)(res)
			If Not aradRes.Succeeded Then
				Throw New Exception($"{aradRes.ErrorCode} - {aradRes.ErrorDescription}")
			End If
			StatusCode = 1
		Catch ex As Exception
			Dim InnerEx As Exception = ExtractMustInnerException(ex)
			errObj = New ClsError(0, $"Message: {IIf(errDesc <> String.Empty, errDesc, InnerEx.Message)}{vbCrLf}{vbCrLf}Stack-Trace: {InnerEx.StackTrace}", InnerEx.Message)
		End Try

		Return New ClsReturn(StatusCode, StatusText, Nothing, errObj)
	End Function

	Function RemoveMoneWtr(ByVal moneId As String, ByVal oldRead As String, ByVal replaceTime As DateTime,
						   ByVal replaceReason As Integer, ByVal odbcName As String, ByVal isuvkod As Integer, ByVal Optional user As String = Nothing,
				  ByVal Optional password As String = Nothing) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn
		Dim dal As AradManager.DalArad = New AradManager.DalArad(odbcName)
		Dim ds As New DsArad
		Dim errDesc As String = String.Empty
		Dim settlement As DsArad.AradSettlementsRow = dal.GetSettlement(ds, errDesc, isuvkod)
		Return HandleMoneWtr(moneId, oldRead, -1, replaceTime, -1, -1, -1, -1, -1, -1, replaceReason, 3, odbcName, settlement.SettlementID, user, password)
	End Function

	Function ReplaceMoneWtr(ByVal moneId As String, ByVal oldRead As String, ByVal newRead As String, ByVal replaceTime As DateTime,
						   ByVal newModelId As Integer, ByVal newDiameterId As Integer, ByVal oldMoneSN As String,
						   ByVal newMoneSN As String, ByVal oldTransponderID As String, ByVal newTransponderID As String,
						   ByVal replaceReason As Integer, ByVal odbcName As String, ByVal isuvkod As Integer, ByVal Optional user As String = Nothing,
				  ByVal Optional password As String = Nothing) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn
		Dim dal As AradManager.DalArad = New AradManager.DalArad(odbcName)
		Dim ds As New DsArad
		Dim errDesc As String = String.Empty
		Dim settlement As DsArad.AradSettlementsRow = dal.GetSettlement(ds, errDesc, isuvkod)
		Return HandleMoneWtr(moneId, oldRead, newRead, replaceTime, newModelId, newDiameterId, oldMoneSN, newMoneSN, oldTransponderID,
							 newTransponderID, replaceReason, 1, odbcName, settlement.SettlementID, user, password)
	End Function

	Private Function HandleMoneWtr(ByVal moneId As String, ByVal oldRead As String, ByVal newRead As String, ByVal replaceTime As DateTime,
						   ByVal newModelId As Integer, ByVal newDiameterId As Integer, ByVal oldMoneSN As String,
						   ByVal newMoneSN As String, ByVal oldTransponderID As String, ByVal newTransponderID As String,
						   ByVal replaceReason As Integer, ByVal operationType As Short, ByVal odbcName As String, ByVal settlementID As Integer, ByVal Optional user As String = Nothing,
				  ByVal Optional password As String = Nothing) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn
		Dim errObj As ClsError = Nothing
		Dim StatusCode As Integer = 0, StatusText As String = ""
		Dim aradApi As AradManager.ApiArad
		Dim res As Object
		Dim aradRes As AradResponse
		Dim handleReq As MoneHandleRequest = New MoneHandleRequest()
		Try
			aradApi = New AradManager.ApiArad(odbcName, user, password)
#Region "query string"
			'Dim properties As PropertyInfo() = GetType(MoneHandleRequest).GetProperties()

			'For Each currProp As PropertyInfo In properties
			'	If currProp.GetValue(req) IsNot Nothing Then
			'		If currProp.GetValue(req).GetType() Is GetType(System.String) AndAlso currProp.GetValue(req) <> "" Then
			'			queryString.Add(currProp.Name, currProp.GetValue(req))
			'		ElseIf currProp.GetValue(req).GetType() IsNot GetType(System.String) AndAlso currProp.GetValue(req) IsNot Nothing Then
			'			queryString.Add(currProp.Name, currProp.GetValue(req))
			'		End If
			'	End If
			'Next

			'queryString.Add("MunicipalID", req.MunicipalID)
			'queryString.Add("MoneID", req.MoneID)
			'queryString.Add("OldMoneSN", req.OldMoneSN)
			'queryString.Add("OldRead", req.OldRead)
			'queryString.Add("NewRead", req.NewRead)
			'queryString.Add("NewMoneSN", req.NewMoneSN)
			'queryString.Add("NewDiameterID", req.NewDiameterID)
			'queryString.Add("NewModelID", req.NewModelID)
			'queryString.Add("NewTransponderID", req.NewTransponderID)
			'queryString.Add("OldTransponderID", req.OldTransponderID)
			'queryString.Add("ReplaceReason", req.ReplaceReason)
			'queryString.Add("ReplaceTime", req.ReplaceTime)
			'queryString.Add("OperationType", req.OperationType)
			'queryString.Add("PropertyID", req.PropertyID)
			'queryString.Add("City", req.City)
			'queryString.Add("Street", req.Street)
			'queryString.Add("Number", req.Number)
			'queryString.Add("Apartment", req.Apartment)
			'queryString.Add("ConsumerCity", req.ConsumerCity)
			'queryString.Add("ConsumerStreet", req.ConsumerStreet)
			'queryString.Add("ConsumerNumber", req.ConsumerNumber)
			'queryString.Add("ConsumerApartment", req.ConsumerApartment)
			'queryString.Add("MoneLocation", req.MoneLocation)
			'queryString.Add("ContractAccount", req.ContractAccount)
			'queryString.Add("ConsumerID", req.ConsumerID)
			'queryString.Add("FirstName", req.FirstName)
			'queryString.Add("LastName", req.LastName)
			'queryString.Add("ZoneNo", req.ZoneNo)
			'queryString.Add("TelNumber", req.TelNumber)
			'queryString.Add("TelNumber2", req.TelNumber2)
			'queryString.Add("PostCode2", req.PostCode2)
#End Region


			handleReq.MunicipalID = settlementID
			handleReq.MoneID = moneId
			handleReq.OldRead = oldRead
			If newRead <> "-1" Then handleReq.NewRead = newRead
			handleReq.ReplaceTime = replaceTime.ToString("yyyy-MM-dd HH:mm")
			If newModelId <> -1 Then handleReq.NewModelID = newModelId '25
			If newDiameterId <> -1 Then handleReq.NewDiameterID = newDiameterId '25
			If oldMoneSN <> "-1" Then handleReq.OldMoneSN = oldMoneSN
			If newMoneSN <> "-1" Then handleReq.NewMoneSN = newMoneSN
			If oldTransponderID <> "-1" Then handleReq.OldTransponderID = oldTransponderID
			If newTransponderID <> "-1" Then handleReq.NewTransponderID = newTransponderID
			handleReq.ReplaceReason = replaceReason
			handleReq.OperationType = operationType

			res = aradApi.HandleMoneArad(handleReq)

			'aradRes = JsonSerializer.Deserialize(Of AradResponse)(res)
			aradRes = JsonConvert.DeserializeObject(Of AradResponse)(res)

			If Not aradRes.Succeeded Then
				Throw New Exception($"{aradRes.ErrorCode} - {aradRes.ErrorDescription}")
			End If

			StatusCode = 1
		Catch ex As Exception
			Dim InnerEx As Exception = ExtractMustInnerException(ex)
			errObj = New ClsError(0, $"Message: {InnerEx.Message}{vbCrLf}{vbCrLf}Stack-Trace: {InnerEx.StackTrace}", InnerEx.Message)
		End Try

		Return New ClsReturn(StatusCode, StatusText, Nothing, errObj)
	End Function

	Function GetReadsFile(ByVal readDate As String, ByVal zoneNo As String, ByVal odbcName As String, ByVal Optional isuvkod As Integer = Nothing, ByVal Optional user As String = Nothing,
					  ByVal Optional password As String = Nothing) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

		Dim errDesc As String = ""
		Dim StatusCode As Integer = 0, StatusText As String = ""
		Dim errObj As ClsError = Nothing
		Dim dal As AradManager.DalArad = Nothing
		'Dim fileData As String = ""
		Dim aradApi As AradManager.ApiArad
		Dim queryString As NameValueCollection = System.Web.HttpUtility.ParseQueryString(String.Empty)
		'Dim outputFileName As String
		Dim outputFileDir As String
		Dim startTime As DateTime
		'Dim fi As FileInfo
		Dim ds As New DsArad
		Dim res As Object
		Dim aradRes As AradReadsFileResponse
		Dim aradReadsRes As AradReadsFileStatusResponse
		Dim fileTaskRef, rowsCount As Integer
		Dim readsFileName As String
		Dim readsFI As FileInfo
		Dim control As DsArad.WtrFileControlRow = Nothing
		Try
			errDesc = ""
			dal = New AradManager.DalArad(odbcName)
			aradApi = New AradManager.ApiArad(odbcName, user, password)
			startTime = DateTime.Now
			rowsCount = 0
			outputFileDir = Path.Combine(dal.GetNetworkDir(), aradApi.ReadingsFileBackupDir, DateTime.Now.ToString("yyyyMM"))

			If Not Directory.Exists(outputFileDir) Then
				Directory.CreateDirectory(outputFileDir)
			End If

			queryString.Add("readDate", readDate)
			If zoneNo IsNot Nothing Then
				queryString.Add("zoneNo", zoneNo)
			End If
			If dal.GetSettlements(ds, errDesc, isuvkod) Then
				For Each row As DsArad.AradSettlementsRow In ds.AradSettlements.Rows
					res = aradApi.ApiAradMethodCall("GetReadsFile", queryString.ToString(), row.SettlementID)

					aradRes = JsonConvert.DeserializeObject(Of AradReadsFileResponse)(res)

					fileTaskRef = aradRes.FileTaskRef
					readsFileName = aradRes.ReadsFileName
					Dim rowID As Integer
					If Not dal.AddWtrFileControl(DateTime.Now.ToString("yyyyMMdd"), 4, row.SettlementID, outputFileDir, readsFileName, 0,
											   0, WtrFileProcIdEnum.ReadsFileRequest, WtrFileProcStatusEnum.Working, startTime,
											   Environment.UserName, ds, errDesc, rowID) Then
						'errObj = New ClsError(0, $"Message: {errDesc}", errDesc)
						Throw New Exception($"AddWtrFileControl exception - {errDesc}")
					End If
					control = ds.WtrFileControl.First(Function(x) x.WtrFileIndex = rowID)
					control.WtrFileTaskRef = fileTaskRef
					control.WtrFileName = readsFileName
					dal.UpdateWtrFileControl(ds, errDesc)

					Threading.Thread.Sleep(2000)

					queryString.Clear()
					queryString.Add("taskRef", fileTaskRef)
					queryString.Add("networkDir", outputFileDir)
					queryString.Add("fileName", readsFileName)

					'GetReadsFileStatus - gets the reads file status, and downloads it if its ready
					res = aradApi.ApiAradMethodCall("ReadsFileProcess", queryString.ToString(), row.SettlementID)

					aradReadsRes = JsonConvert.DeserializeObject(Of AradReadsFileStatusResponse)(res)

					If aradReadsRes.ErrorCode = -1 Then
						Throw New Exception($"AradApi.ReadsFileProcess error - {aradReadsRes.ErrorDescription}")
					End If

					Dim readsFileFullName = Path.Combine(outputFileDir, readsFileName)

					If File.Exists(readsFileFullName) Then
						File.Delete(readsFileFullName)
					End If

					File.WriteAllText(readsFileFullName, aradReadsRes.FileData)

					If aradReadsRes.ResultsFileRowsCount = 0 Then
						Throw New Exception($"No results - file was not created")
					End If

					If Not File.Exists(readsFileFullName) Then
						Throw New Exception($"readsFile {readsFileFullName} Not exists")
					End If

					readsFI = New FileInfo(readsFileFullName)

					If Not ReadReadsFile(row.SettlementID, readsFileFullName, ds, rowsCount, errDesc) Then
						Throw New Exception($"ReadReadsFile exception - {errDesc}")
					End If

					If Not dal.UpdateReads(ds, errDesc) Then
						Throw New Exception($"UpdateReads exception - {errDesc}")
					End If

					control.WtrFileDir = outputFileDir
					control.WtrFileSize = readsFI.Length
					control.WtrFileRowsCount = rowsCount
					control.ProcessFinished = DateTime.Now
					control.WtrFileProcessStatus = WtrFileProcStatusEnum.Finished
					dal.UpdateWtrFileControl(ds, errDesc)
				Next
			End If

			StatusCode = 1
		Catch ex As Exception
			Try
				AddDbErrors(dal.SqlConnString, TraceException(ex.Message))
			Catch

			End Try

			If dal IsNot Nothing And control IsNot Nothing Then
				control.ProcessFinished = DateTime.Now
				control.WtrFileProcessStatus = WtrFileProcStatusEnum.Failed
				control.ProcErrorDesc = ex.Message
				dal.UpdateWtrFileControl(ds, errDesc)
			End If

			Dim InnerEx As Exception = ExtractMustInnerException(ex)
			errObj = New ClsError(0, $"Message: {InnerEx.Message}{vbCrLf}{vbCrLf}Stack-Trace: {InnerEx.StackTrace}", InnerEx.Message)
		End Try

		Return New ClsReturn(StatusCode, StatusText, Nothing, errObj)
	End Function

	Private Function ReadReadsFile(ByVal muniCode As Integer, ByVal filename As String, ByRef ds As DsArad, ByRef rowsCount As Integer, ByRef errDesc As String) As Boolean
		Dim lines() As String
		Dim newRead As DsArad.ReadsDataRow
		Dim zihuyLong As Long
		Try
			rowsCount = 0
			ds.ReadsData.Clear()
			lines = IO.File.ReadAllLines(filename)

			For Each line In lines
				rowsCount = rowsCount + 1
				zihuyLong = 0
				If line.Substring(0, 6) = muniCode Then
					newRead = ds.ReadsData.NewReadsDataRow
					If Long.TryParse(line.Substring(10, 15), zihuyLong) Then
						newRead.ZihuyWater = zihuyLong.ToString()
					Else
						newRead.ZihuyWater = line.Substring(10, 15)
					End If
					newRead.LastRead = System.Math.Round(Convert.ToDecimal(line.Substring(25, 10)), 0)
					newRead.CurrRead = System.Math.Round(Convert.ToDecimal(line.Substring(35, 10)), 0)
					Dim readData As String = line.Substring(47, 6)
                    newRead.ReadDate = New DateTime("20" + readData.Substring(4, 2), readData.Substring(2, 2), readData.Substring(0, 2))
                    ds.ReadsData.AddReadsDataRow(newRead)
				End If
			Next

			Return True
		Catch ex As Exception
			errDesc = ex.Message
			Return False
		End Try
	End Function

End Class
