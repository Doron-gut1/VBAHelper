Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Net.WebRequestMethods
Imports System.Text
Imports System.Text.Json
Imports System.Text.Json.Serialization
Imports System.Threading.Tasks
Imports System.Web.Services.Description
Imports CreditGuard.My
Imports Epr.AradWaterStructures.Request
Imports Epr.AradWaterStructures.Response
Imports Microsoft.SqlServer
Imports Newtonsoft.Json

Namespace AradManager
	Public Class ApiArad

		Private GwURL As String

		Public GwUser As String
		Public GwPassword As String
		Public GwEntity As Integer

		'Public MuniCode As Integer
		Public KaramCompany As Integer
		Public KovetzAvFileBackupDir As String
		Public ReadingsFileBackupDir As String

		Private m_DalArad As DalArad
		Private m_DsArad As New DsArad

		'gwUser=orik-office-brener-test,gwPass=Epr123Epr,gwEntity=270
		Public Sub New(ByVal odbcName As String, ByVal Optional user As String = Nothing, ByVal Optional password As String = Nothing)
			Dim errDesc As String = ""
			Try
				m_DalArad = New DalArad(odbcName, user, password)
				GetGatewayAuth()

				If Not m_DalArad.GetKaramParam(m_DsArad, errDesc) Then
					Throw New Exception(errDesc)
				End If

				'MuniCode = m_DsArad.KaramParam(0).muniCode
				KovetzAvFileBackupDir = m_DsArad.KaramParam(0).BackupKaramKovetzAv
				ReadingsFileBackupDir = m_DsArad.KaramParam(0).BackupKaramReadingsFiles

			Catch ex As Exception
				Throw New Exception(TraceException(ex.Message))
			End Try
		End Sub

		Private Function GetFullUri(ByVal baseUri As String, ByVal relativeUri As String) As String
			'Dim myBaseUri As Uri = New Uri(baseUri)
			'Dim fullUri = New Uri(myBaseUri, relativeUri)
			'Return fullUri.ToString()
			baseUri = baseUri.TrimEnd("/")
			relativeUri = relativeUri.TrimStart("/")
			Return string.Format("{0}/{1}", baseUri, relativeUri)
		End Function


		Private Sub GetGatewayAuth()
			Dim errDesc As String = ""

			If Not m_DalArad.GetApiGwParams(m_DsArad, errDesc) Then
				Throw New Exception(errDesc)
			End If

			GwUser = m_DsArad.ParamApiGateway(0).ApiGwUser
			GwPassword = m_DsArad.ParamApiGateway(0).ApiGwPass
			GwEntity = m_DalArad.GetThisMoz.ToString
			GwURL = GetFullUri(m_DsArad.ParamApiGateway(0).ApiGwUrl, "/AradWater/")
		End Sub

		Public Function ApiAradMethodCall(ByVal methodName As String, ByVal paramQuery As String, ByVal settlementID As Integer, Optional ByVal fileName As String = "") As Object
			'Dim requestUri As String = $"{GwURL}/{methodName}?MunicipalID={m_DsArad.KaramParam(0).muniCode}&{paramQuery}"
			Dim requestUri As String = GetFullUri(GwURL, $"{methodName}?MunicipalID={settlementID}&{paramQuery}")

			Using message As HttpRequestMessage = New HttpRequestMessage(HttpMethod.[Get], requestUri)
				message.Headers.Authorization = New AuthenticationHeaderValue("Basic", Convert.ToBase64String(Encoding.UTF8.GetBytes($"{GwUser}:{GwPassword}")))
				message.Headers.Add("X-Epr-Entity", GwEntity)

				Using client As HttpClient = New HttpClient()
					ServicePointManager.ServerCertificateValidationCallback = Function(obj, cert, chain, err) True

					Using task1 As Task(Of HttpResponseMessage) = client.SendAsync(message)
						task1.Wait()

						If Not task1.Result.IsSuccessStatusCode Then
							Throw New Exception($"{requestUri} Error: {task1.Result.ReasonPhrase}({CInt(task1.Result.StatusCode)})")
						Else
							Using task2 As Task(Of String) = task1.Result.Content.ReadAsStringAsync()
								task2.Wait()
								Dim obj As Object = task2.Result
								Return obj
							End Using
						End If
					End Using
				End Using
			End Using
		End Function

		Public Function UploadAvFile(ByVal dirName As String, ByVal fileName As String, ByVal SettlementID As Integer) As Object
			Dim requestUri As String = GetFullUri(GwURL, "AvFile") '$"{GwURL}AvFile" 

			Dim fileReq As AvFileRequest

			Using message As HttpRequestMessage = New HttpRequestMessage(HttpMethod.[Post], requestUri)
				message.Headers.Authorization = New AuthenticationHeaderValue("Basic", Convert.ToBase64String(Encoding.UTF8.GetBytes($"{GwUser}:{GwPassword}")))
				message.Headers.Add("X-Epr-Entity", GwEntity)
				Dim bytesArray As Byte()
				Dim jsonData As String

				Using client As HttpClient = New HttpClient()
					ServicePointManager.ServerCertificateValidationCallback = Function(obj, cert, chain, err) True

					bytesArray = IO.File.ReadAllBytes(Path.Combine(dirName, fileName))

					fileReq = New AvFileRequest()
					fileReq.MunicipalID = SettlementID
					fileReq.FileName = fileName
					fileReq.FileData = Convert.ToBase64String(bytesArray)
					jsonData = JsonConvert.SerializeObject(fileReq)

					Dim contentData As StringContent = New StringContent(jsonData)
					contentData.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json")

					message.Content = contentData

					Using task1 As Task(Of HttpResponseMessage) = client.SendAsync(message)
						task1.Wait()

						If Not task1.Result.IsSuccessStatusCode Then
							Throw New Exception($"{requestUri} Error: {task1.Result.ReasonPhrase}({CInt(task1.Result.StatusCode)})")
						Else
							Using task2 As Task(Of String) = task1.Result.Content.ReadAsStringAsync()
								task2.Wait()
								Dim obj As Object = task2.Result
								Return obj
							End Using
						End If
					End Using
				End Using
			End Using
		End Function

		Public Function HandleMoneArad(ByVal handleRequest As MoneHandleRequest) As Object
			Dim requestUri As String = GetFullUri(GwURL, "HandleMoneWtr") '$"{GwURL}HandleMoneWtr"

			Using message As HttpRequestMessage = New HttpRequestMessage(HttpMethod.[Post], requestUri)
				message.Headers.Authorization = New AuthenticationHeaderValue("Basic", Convert.ToBase64String(Encoding.UTF8.GetBytes($"{GwUser}:{GwPassword}")))
				message.Headers.Add("X-Epr-Entity", GwEntity)
				Dim jsonData As String

				Using client As HttpClient = New HttpClient()
					ServicePointManager.ServerCertificateValidationCallback = Function(obj, cert, chain, err) True

					jsonData = JsonConvert.SerializeObject(handleRequest)

					Dim contentData As StringContent = New StringContent(jsonData)
					contentData.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json")

					message.Content = contentData

					Using task1 As Task(Of HttpResponseMessage) = client.SendAsync(message)
						task1.Wait()

						If Not task1.Result.IsSuccessStatusCode Then
							Throw New Exception($"{requestUri} Error: {task1.Result.ReasonPhrase}({CInt(task1.Result.StatusCode)})")
						Else
							Using task2 As Task(Of String) = task1.Result.Content.ReadAsStringAsync()
								task2.Wait()
								Dim obj As Object = task2.Result
								Return obj
							End Using
						End If
					End Using
				End Using
			End Using
		End Function
	End Class
End Namespace
