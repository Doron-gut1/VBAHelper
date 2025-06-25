Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Data.SqlClient
Imports System.Net
Imports System.Reflection
Imports System.Web
Imports System.Threading.Tasks
Imports System.Threading
Imports System.Data


Namespace SmsManager
	Public Class SmsTypeFunctions
		Private Const MAX_FILE_SIZE As Integer = 3 * 1024 * 1024
		Private ReadOnly m_DalSms As DalSms
		Private m_Params As List(Of SmsParams)

		Public Sub New(ByVal odbcName As String, ByVal Optional user As String = Nothing, ByVal Optional password As String = Nothing)
			Try
				m_DalSms = New DalSms(odbcName, user, password)
			Catch ex As Exception
				Throw New Exception(TraceException(ex.Message))
			End Try

		End Sub

		Public Function SendInvoiceLinkSms(ByVal invoiceRashut As Integer, ByVal invoiceNum As Integer, ByVal tel As String,
										   ByVal pdfLocation As String, ByVal shouldOverwrite As Boolean, ByRef smsStatus As Integer,
										   ByRef smsStatusDesc As String, ByRef errDesc As String) As Boolean

			Dim pdfFile As FileInfo
			Dim msg As String
			Dim shortUrl, fullUrl As String
			Dim ds As New DsSms
			Dim shouldAddNew As Boolean = False
			Try
				'Get ParamsSms
				m_Params = m_DalSms.GetParamSmsType(SmsTypeEnum.InvoiceLink)
				pdfFile = New FileInfo(pdfLocation)

				If m_Params Is Nothing OrElse m_Params.Count = 0 Then
					Throw New Exception($"Missing parameters for SmsType = {SmsTypeEnum.InvoiceLink} (InvoiceLink)")
				ElseIf Not pdfFile.Exists Then
					Throw New Exception($"File {pdfLocation} not exists")
				ElseIf pdfFile.Length = 0 Then
					Throw New Exception($"File {pdfLocation} is empty")
				ElseIf pdfFile.Length > Helper.SmsManager.SmsTypeFunctions.MAX_FILE_SIZE Then
					Throw New Exception($"File {pdfLocation} is too large")
				End If

				fullUrl = ""
				shortUrl = ""

				Using trans As SqlTransaction = m_DalSms.CreateTransaction()

					'Checking if invoice already exists
					If Not m_DalSms.GetInvoiceLink(invoiceNum, trans, ds, errDesc) Then Throw New Exception(errDesc)

					If ds.InvoiceLink.Count = 0 OrElse Not ds.InvoiceLink(0).IsActiveRow Then
						shouldAddNew = True
					End If

					If shouldAddNew Then
						'Adding new invoice - getting back invoicelink data
						If Not m_DalSms.AddInvoiceLink(invoiceRashut, invoiceNum, pdfFile.Length, m_Params(0).SmsUrl, trans, ds, errDesc) Then Throw New Exception(errDesc)

						If Not ds.InvoiceLink.Count = 1 Then Throw New Exception("error getting Invoice link data")

						'Copying New pdf to network dir
						CopyPdfToNetworkDir(pdfFile, ds.InvoiceLink(0).InvoiceGuid, invoiceRashut)

						fullUrl = ds.InvoiceLink(0).Url

						'call URL shortener service 
						If Not GetShortUrl(fullUrl, shortUrl, errDesc) Then Throw New Exception(errDesc)

						'updating shorturl
						ds.InvoiceLink(0).ShortUrl = shortUrl

						m_DalSms.UpdateInvoiceLink(ds, errDesc, trans)
					Else
						If shouldOverwrite Then
							'Copying New pdf to network dir
							CopyPdfToNetworkDir(pdfFile, ds.InvoiceLink(0).InvoiceGuid, invoiceRashut)

							'updating pdf file size
							ds.InvoiceLink(0).InvoiceFileSize = pdfFile.Length

							m_DalSms.UpdateInvoiceLink(ds, errDesc, trans)
						End If
					End If

					'build sms message
					msg = BuildMessage(ds.InvoiceLink(0).ShortUrl)

					'send link in sms
					Dim wsSMS As SmsSender.WsInforu = New SmsSender.WsInforu()

					If wsSMS.Send(msg, New List(Of String) From {tel}, m_Params(0).SmsSender, smsStatus, errDesc) Then
						smsStatusDesc = wsSMS.GetEnumDescription(DirectCast(smsStatus, SmsSender.WsInforu.SmsStatusEnum))
						m_DalSms.AddSmsLog(ds.InvoiceLink(0).InvoiceId, m_Params(0).SmsType, m_Params(0).SmsSender, msg, msg.Length, tel, smsStatus, errDesc, trans)
					Else
						Throw New Exception(errDesc)
					End If

					trans.Commit()
				End Using
				Return True
			Catch ex As Exception
				errDesc = TraceException(ex.Message)
				Return False
			End Try
		End Function

		Public Function SendAllShovarLinkSms(ByVal ShovarType As SmsTypeEnum, ByVal shovarRashut As Integer, ByVal shovarNum As Integer, ByVal tz As Long, ByVal tel As String,
										 ByRef smsStatus As Integer, ByRef smsStatusDesc As String, ByRef errDesc As String) As Boolean

			Dim msg As String
			Dim shortUrl, fullUrl As String
			Dim ds As New DsSms
			Dim shovarId As Integer = 0
			Try
				'Get ParamsSms
				m_Params = m_DalSms.GetParamSmsType(ShovarType)

				Using trans As SqlTransaction = m_DalSms.CreateTransaction()
					fullUrl = $"{m_Params(0).SmsUrl}?rshut={shovarRashut}&maintz={tz}&shovarnumber={shovarNum}"
					shortUrl = ""

					'call URL shortener service 
					If Not GetShortUrl(fullUrl, shortUrl, errDesc) Then Throw New Exception(errDesc)

					m_DalSms.AddShovarLink(shovarNum, tz, fullUrl, shortUrl, shovarId, trans, errDesc)

					'build sms message
					msg = BuildMessage(shortUrl)

					'send link in sms
					Dim wsSMS As SmsSender.WsInforu = New SmsSender.WsInforu()

					If wsSMS.Send(msg, New List(Of String) From {tel}, m_Params(0).SmsSender, smsStatus, errDesc) Then
						smsStatusDesc = wsSMS.GetEnumDescription(DirectCast(smsStatus, SmsSender.WsInforu.SmsStatusEnum))
						m_DalSms.AddSmsLog(shovarId, m_Params(0).SmsType, m_Params(0).SmsSender, msg, msg.Length, tel, smsStatus, errDesc, trans)
					Else
						Throw New Exception(errDesc)
					End If

					trans.Commit()
				End Using
				Return True
			Catch ex As Exception
				errDesc = TraceException(ex.Message)
				Return False
			End Try
		End Function

		Public Function SendImutPhonePassSms(ByVal rnd As String, ByVal tel As String, ByRef smsStatus As Integer,
										ByRef smsStatusDesc As String, ByRef errDesc As String) As Boolean

			Dim msg As String
			Dim ds As New DsSms
			Try
				'Get ParamsSms
				m_Params = m_DalSms.GetParamSmsType(SmsTypeEnum.ImutPhonePass)

				Using trans As SqlTransaction = m_DalSms.CreateTransaction()
					'build sms message
					msg = BuildMessage(rnd)

					'send link in sms
					Dim wsSMS As SmsSender.WsInforu = New SmsSender.WsInforu()

					If wsSMS.Send(msg, New List(Of String) From {tel}, m_Params(0).SmsSender, smsStatus, errDesc) Then
						smsStatusDesc = wsSMS.GetEnumDescription(DirectCast(smsStatus, SmsSender.WsInforu.SmsStatusEnum))
						m_DalSms.AddSmsLog(0, m_Params(0).SmsType, m_Params(0).SmsSender, msg, msg.Length, tel, smsStatus, errDesc, trans)
					Else
						Throw New Exception(errDesc)
					End If

					trans.Commit()
				End Using
				Return True
			Catch ex As Exception
				errDesc = TraceException(ex.Message)
				Return False
			End Try
		End Function

		Public Function SendUpdateInfoPassSms(ByVal rnd As String, ByVal tel As String, ByRef smsStatus As Integer,
										ByRef smsStatusDesc As String, ByRef errDesc As String) As Boolean

			Dim msg As String
			Dim ds As New DsSms
			Try
				'Get ParamsSms
				m_Params = m_DalSms.GetParamSmsType(SmsTypeEnum.UpdateInfoPass)

				Using trans As SqlTransaction = m_DalSms.CreateTransaction()
					'build sms message
					msg = BuildMessage(rnd)

					'send link in sms
					Dim wsSMS As SmsSender.WsInforu = New SmsSender.WsInforu()

					If wsSMS.Send(msg, New List(Of String) From {tel}, m_Params(0).SmsSender, smsStatus, errDesc) Then
						smsStatusDesc = wsSMS.GetEnumDescription(DirectCast(smsStatus, SmsSender.WsInforu.SmsStatusEnum))
						m_DalSms.AddSmsLog(0, m_Params(0).SmsType, m_Params(0).SmsSender, msg, msg.Length, tel, smsStatus, errDesc, trans)
					Else
						Throw New Exception(errDesc)
					End If

					trans.Commit()
				End Using
				Return True
			Catch ex As Exception
				errDesc = TraceException(ex.Message)
				Return False
			End Try
		End Function


		Private Sub CopyPdfToNetworkDir(ByVal pdfFile As FileInfo, ByVal invoiceGuid As Guid, ByVal invoiceRashut As Integer)
			Dim pdfNetworkName As String
			Dim pdfGuidFile As FileInfo

			'setting pdf netwrok name
			pdfNetworkName = $"{m_Params(0).NetworkFolder}\{invoiceRashut.ToString().PadLeft(6, "0"c)}\{invoiceGuid}.pdf"
			pdfGuidFile = New FileInfo(pdfNetworkName)

			'Creating network folder if not exists
			If Not Directory.Exists(pdfGuidFile.Directory.FullName) Then Directory.CreateDirectory(pdfGuidFile.Directory.FullName)

			'Save pdf in network folder with guid name, if not exists already
			'If Not File.Exists(pdfNetworkName) Then
			File.Copy(pdfFile.FullName, pdfNetworkName, True)
			If Not File.Exists(pdfNetworkName) Then Throw New Exception("Failed copying file to network folder")
			'End If
		End Sub

		Private Function BuildMessage(ByVal link As String) As String
			Return $"{m_Params(0).SmsHeaderText} {link} {m_Params(0).SmsFooterText}"
		End Function

		'Public Function PayLinkSmsSender(ByRef errDesc As String) As Boolean
		'	Return SmsSenderByType(errDesc, SmsTypeEnum.PayLink)
		'End Function

		'Public Function SmsSenderByType(ByRef errDesc As String, ByVal Optional smsType As SmsTypeEnum? = Nothing) As Boolean
		'	Return True
		'End Function

		Private Function GetShortUrl(longUrl As String, ByRef shortUrl As String, ByRef errDesc As String) As Boolean
			Dim auth As String = "5ef23b4d-3659-49fd-9bde-4f8aebc74b38-810179c7-5f9d-4adc-8cd0-cfbacb486aa8"
			Dim expireDate As DateTime = DateTime.Now.AddMonths(3)
			Dim reqUrl As String = $"{m_DalSms.GetUrlShortenerApiParam}?auth={auth}&url={Uri.EscapeDataString(longUrl)}&expiryDate={expireDate.ToString("MM/dd/yyyy")}"
			Dim request As HttpWebRequest
			Try
				request = WebRequest.Create(reqUrl)
				request.ServerCertificateValidationCallback = Function(sender, cert, chain, [error]) True
				request.Method = "GET"
				'request.KeepAlive = true;
				request.ContentType = "application/x-www-form-urlencoded"
				Dim response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)

				Using sr As New IO.StreamReader(response.GetResponseStream())
					shortUrl = sr.ReadToEnd()
					Return True
				End Using
			Catch ex As Exception
				errDesc = TraceException(ex.Message)
				Return False
			End Try
		End Function


		Public Function SendHinHarshama(ByVal yldName As String, ByVal tel As String, ByRef smsStatus As Integer, ByRef smsStatusDesc As String, ByRef errDesc As String) As Boolean


			Dim msg As String
			Dim ds As New DsSms
			Dim shovarId As Integer = 0
			Try
				'Get ParamsSms
				m_Params = m_DalSms.GetParamSmsType(SmsTypeEnum.HinHarshama)

				Using trans As SqlTransaction = m_DalSms.CreateTransaction()



					'build sms message
					msg = $"{BuildMessage(yldName)}{vbCrLf}{m_Params(0).SmsUrl} "

					'send link in sms
					Dim wsSMS As SmsSender.WsInforu = New SmsSender.WsInforu()

					If wsSMS.Send(msg, New List(Of String) From {tel}, m_Params(0).SmsSender, smsStatus, errDesc) Then
						smsStatusDesc = wsSMS.GetEnumDescription(DirectCast(smsStatus, SmsSender.WsInforu.SmsStatusEnum))
						m_DalSms.AddSmsLog(shovarId, m_Params(0).SmsType, m_Params(0).SmsSender, msg, msg.Length, tel, smsStatus, errDesc, trans)
					Else
						Throw New Exception(errDesc)
					End If

					trans.Commit()
				End Using
				Return True
			Catch ex As Exception
				errDesc = TraceException(ex.Message)
				Return False
			End Try
		End Function
		Public Function SendGenearlMessage(ByVal Content As String, ByVal tel As String, ByRef smsStatus As Integer, ByRef smsStatusDesc As String, ByRef errDesc As String) As Boolean


			Dim msg As String
			Dim ds As New DsSms
			Dim shovarId As Integer = 0
			Try
				'Get ParamsSms
				m_Params = m_DalSms.GetParamSmsType(SmsTypeEnum.GeneralSms)

				Using trans As SqlTransaction = m_DalSms.CreateTransaction()



					'build sms message
					msg = $"{BuildMessage(Content)}{vbCrLf}{m_Params(0).SmsUrl} "

					'send link in sms
					Dim wsSMS As SmsSender.WsInforu = New SmsSender.WsInforu()

					If wsSMS.Send(msg, New List(Of String) From {tel}, m_Params(0).SmsSender, smsStatus, errDesc) Then
						smsStatusDesc = wsSMS.GetEnumDescription(DirectCast(smsStatus, SmsSender.WsInforu.SmsStatusEnum))
						m_DalSms.AddSmsLog(shovarId, m_Params(0).SmsType, m_Params(0).SmsSender, msg, msg.Length, tel, smsStatus, errDesc, trans)
					Else
						Throw New Exception(errDesc)
					End If

					trans.Commit()
				End Using
				Return True
			Catch ex As Exception
				errDesc = TraceException(ex.Message)
				Return False
			End Try
		End Function
		Public Async Function SendShvaHkUnacceptableSms(ByVal thum As Integer, ByVal sendnum As Integer, SendUniqueId As String, UserName As String, smsStatus As Integer, smsStatusDesc As String, errDesc As String) As Task(Of Boolean)

			Dim shovarId As Integer = 0
			Try
				'Get ParamsSms
				m_Params = m_DalSms.GetParamSmsType(SmsTypeEnum.UnacceptableHk)
				Using trans As SqlTransaction = m_DalSms.CreateTransaction()
					If Not m_DalSms.FillShvaHkUnacceptableData(thum, sendnum, SendUniqueId, errDesc, trans) Then
						errDesc = " בעיה בהוספת רשימת התושבים שלא נשלחה להם הוראת הקבע - FillShvaHkUnacceptableData " & errDesc
						Throw New Exception(errDesc)
					End If
				End Using

				Using trans As SqlTransaction = m_DalSms.CreateTransaction()

					' Dim msg As String
					Dim ds As New DsSms
					If m_DalSms.GetShvaHkUnacceptableSmsData(thum, sendnum, ds) Then
						Await ProcessDataTableAsync(ds.ShvaHkUnacceptableSmsData, trans, UserName)
					Else
						errDesc = " בעיה בקבלת המשלמים להם יש לשלוח הודעת SMS - GetShvaHkUnacceptableSmsData " & errDesc
						Throw New Exception(errDesc)
					End If
					If Not m_DalSms.UpdateShvaHkUnacceptableSmsStatus(ds, trans) Then
						errDesc = " בעיה בעדכון סטטוס ההודעות שנשלחו SMS - UpdateShvaHkUnacceptableSmsStatus " & errDesc
						Throw New Exception(errDesc)
					End If
					' trans.Commit()
				End Using
				Return True
			Catch ex As Exception
				errDesc = TraceException(ex.Message)
				Throw New Exception(errDesc)
				Return False
			End Try
		End Function
		Private Function SendOneShvaHkUnacceptableSms(row As DataRow, trans As SqlTransaction, UserName As String) As Task
			' Your existing logic here
			Dim wsSMS As SmsSender.WsInforu = New SmsSender.WsInforu()
			Dim msg As String = "", errDesc As String = "", smsStatus As Integer, smsStatusDesc As String
			Try
				msg = m_Params(0).SmsUrl
				msg = Replace(msg, "[mnt]", row("mntname"))
				msg = Replace(msg, "[last4krts]", row("last4krts"))
				msg = Replace(msg, "[reason]", row("stat"))
				msg = Replace(msg, "[hksum]", row("hksum"))
				msg = $"{msg}{vbCrLf}{m_Params(0).SmsHeaderText}{vbCrLf}{m_Params(0).SmsFooterText}"
				'send link in sms

				If wsSMS.Send(msg, New List(Of String) From {row("ntel")}, m_Params(0).SmsSender, smsStatus, errDesc) Then
					smsStatusDesc = wsSMS.GetEnumDescription(DirectCast(smsStatus, SmsSender.WsInforu.SmsStatusEnum))
					row("Sent") = 2
					m_DalSms.AddSmsLog(0, m_Params(0).SmsType, m_Params(0).SmsSender, msg, msg.Length, row("ntel"), smsStatus, errDesc, trans)
					m_DalSms.InsertPrtnfs(row("mspkod"), "שליחת SMS  - ה.ק אשראי שלא כובדה", msg, UserName, errDesc, trans)
				Else
					Throw New Exception(errDesc)
				End If
			Catch ex As Exception
				row("Sent") = -1
				errDesc = TraceException(ex.Message)
				m_DalSms.AddSmsLog(0, m_Params(0).SmsType, m_Params(0).SmsSender, msg, msg.Length, row("ntel"), smsStatus, errDesc, trans)
			End Try
		End Function


		Public Async Function ProcessDataTableAsync(tb As DataTable, trans As SqlTransaction, UserName As String) As Task
			Dim tasks As New List(Of Task)(), errDesc As String = ""
			Try

				For Each row As DataRow In tb.Rows
					tasks.Add(SendOneShvaHkUnacceptableSms(row, trans, UserName))
				Next

				Await Task.WhenAll(tasks)
			Catch ex As Exception
				errDesc = TraceException(ex.Message)
			End Try
		End Function

#Region "Batch Processing Functions for SmsSeqSender"

		''' <summary>
		''' פונקציה ראשית לעיבוד מנת שוברים
		''' </summary>
		Public Function ProcessShovarBatch(ByVal manaHovNum As Integer,
										  ByVal smsType As SmsTypeEnum,
										  ByVal isRecovery As Boolean,
										  ByRef totalInBatch As Integer,
										  ByRef newSuccessCount As Integer,
										  ByRef totalSuccessCount As Integer,
										  ByRef failedCount As Integer,
										  ByRef errDesc As String) As Boolean
			Try
				' איפוס מונים
				totalInBatch = 0
				newSuccessCount = 0
				totalSuccessCount = 0
				failedCount = 0

				' קבלת פרמטרים לסוג ההודעה
				m_Params = m_DalSms.GetParamSmsType(smsType)

				If m_Params Is Nothing OrElse m_Params.Count = 0 Then
					Throw New Exception($"Missing parameters for SmsType = {smsType}")
				End If

				' שמירת מצב התחלתי במצב התאוששות
				Dim initialSuccessCount As Integer = 0
				If isRecovery Then
					' ספירת כמה כבר נשלחו בהצלחה לפני ההתאוששות
					Dim ds As New DsSms
					If m_DalSms.GetSmsSeqSenderByBatch(manaHovNum, 3, ds, errDesc) Then
						If ds.Tables.Contains("SmsSeqSender") Then
							initialSuccessCount = ds.Tables("SmsSeqSender").Rows.Count
						End If
					End If
				Else
					' הכנסת רשומות חדשות - רק אם לא במצב התאוששות
					Dim rowsInserted As Integer = m_DalSms.InsertSmsSeqSenderFromSelect(manaHovNum, errDesc)
					If rowsInserted < 0 Then
						Throw New Exception("Failed to insert records: " & errDesc)
					End If
				End If

				' שלב 1: בדיקת טלפונים
				If Not ValidatePhonesBatch(manaHovNum, isRecovery) Then
					Throw New Exception("Failed in phone validation step")
				End If

				' שלב 2: יצירת קישורים מקוצרים
				Dim step2Task = CreateShortUrlsBatchAsync(manaHovNum, smsType, isRecovery)
				Dim step2Result = step2Task.GetAwaiter().GetResult()
				If Not step2Result Then
					Throw New Exception("Failed in short URL creation step")
				End If

				' שלב 3: שליחת הודעות SMS
				Dim step3Task = SendSmsBatchAsync(manaHovNum, smsType, isRecovery)
				Dim step3Result = step3Task.GetAwaiter().GetResult()
				If Not step3Result Then
					Throw New Exception("Failed in SMS sending step")
				End If

				' קבלת סיכום תוצאות
				If Not GetBatchSummary(manaHovNum, isRecovery, initialSuccessCount,
					  totalInBatch, newSuccessCount, totalSuccessCount,
					  failedCount, errDesc) Then
					Throw New Exception("Failed to get batch summary: " & errDesc)
				End If

				Return True

			Catch ex As Exception
				errDesc = TraceException(ex.Message)
				Return False
			End Try
		End Function
		''' <summary>
		''' קבלת סיכום תוצאות של מנת שוברים
		''' </summary>
		Private Function GetBatchSummary(ByVal manaHovNum As Integer,
								ByVal isRecovery As Boolean,
								ByVal initialSuccessCount As Integer,
								ByRef totalInBatch As Integer,
								ByRef newSuccessCount As Integer,
								ByRef totalSuccessCount As Integer,
								ByRef failedCount As Integer,
								ByRef errDesc As String) As Boolean
			Try
				Dim dsFinal As New DsSms
				If Not m_DalSms.GetSmsSeqSenderByBatch(manaHovNum, Nothing, dsFinal, errDesc) Then
					Return False
				End If

				If Not dsFinal.Tables.Contains("SmsSeqSender") Then
					Return False
				End If

				Dim dt = dsFinal.Tables("SmsSeqSender")
				totalInBatch = dt.Rows.Count

				' איפוס מונים
				totalSuccessCount = 0
				failedCount = 0

				' ספירת תוצאות
				For Each row As DataRow In dt.Rows
					Dim status As Integer = CInt(row("Status"))
					If status = 3 Then
						totalSuccessCount += 1
					ElseIf status < 0 Then
						failedCount += 1
					End If
				Next

				' חישוב כמה חדשים נשלחו
				If isRecovery Then
					newSuccessCount = totalSuccessCount - initialSuccessCount
				Else
					newSuccessCount = totalSuccessCount
				End If

				Return True

			Catch ex As Exception
				errDesc = TraceException(ex.Message)
				Return False
			End Try
		End Function
		''' <summary>
		''' שלב 1: בדיקת תקינות טלפונים
		''' </summary>
		Private Function ValidatePhonesBatch(ByVal manaHovNum As Integer, ByVal isRecovery As Boolean) As Boolean
			Dim ds As New DsSms
			Dim errDesc As String = ""

			Try
				' קבלת רשומות עם סטטוס 0 או שלילי (במצב התאוששות)
				If isRecovery Then
					' במצב התאוששות - מביאים כל רשומה שלא בסטטוס חיובי > 0
					If Not m_DalSms.GetSmsSeqSenderByBatch(manaHovNum, Nothing, ds, errDesc) Then
						Throw New Exception(errDesc)
					End If

					If Not ds.Tables.Contains("SmsSeqSender") OrElse ds.Tables("SmsSeqSender").Rows.Count = 0 Then
						Return True
					End If
				Else
					' מצב רגיל - רק סטטוס 0
					If Not m_DalSms.GetSmsSeqSenderByBatch(manaHovNum, 0, ds, errDesc) Then
						Throw New Exception(errDesc)
					End If

					If Not ds.Tables.Contains("SmsSeqSender") OrElse ds.Tables("SmsSeqSender").Rows.Count = 0 Then
						Return True
					End If
				End If

				' בדיקת כל טלפון ועדכון סטטוס
				For Each row As DataRow In ds.Tables("SmsSeqSender").Rows
					Dim currentStatus As Integer = CInt(row("Status"))

					' במצב התאוששות, מעבדים רק את מי שעדיין לא נבדק או נכשל בבדיקה
					If isRecovery AndAlso (currentStatus = 0 OrElse currentStatus = -1) Then
						' בודקים רק את מי שצריך
					ElseIf isRecovery Then
						Continue For
					End If

					Dim phone As String = If(IsDBNull(row("cell")), "", row("cell").ToString())
					Dim isValid As Boolean = IsValidPhoneNumber(phone)
					Dim newStatus As Integer = If(isValid, 1, -1)

					' עדכון סטטוס
					m_DalSms.UpdateSmsSeqSenderStatus(CInt(row("shovar")), manaHovNum, newStatus, errDesc)
				Next

				Return True

			Catch ex As Exception
				Throw
			End Try
		End Function

		''' <summary>
		''' שלב 2: יצירת קישורים מקוצרים
		''' </summary>
		Private Async Function CreateShortUrlsBatchAsync(ByVal manaHovNum As Integer, ByVal smsType As SmsTypeEnum, ByVal isRecovery As Boolean) As Task(Of Boolean)
			Dim ds As New DsSms
			Dim errDesc As String = ""
			Dim shovarRashut As Integer = 0

			Try
				' קבלת מספר רשות מהפרמטרים או מה-DB
				shovarRashut = m_DalSms.GetRashutNumber(errDesc)

				' קבלת רשומות לעיבוד
				If isRecovery Then
					' במצב התאוששות - מביאים כל רשומה שלא בסטטוס 2 או 3
					If Not m_DalSms.GetSmsSeqSenderByBatch(manaHovNum, Nothing, ds, errDesc) Then
						Throw New Exception(errDesc)
					End If

					If Not ds.Tables.Contains("SmsSeqSender") OrElse ds.Tables("SmsSeqSender").Rows.Count = 0 Then
						Return True
					End If
				Else
					' מצב רגיל - רק סטטוס 1
					If Not m_DalSms.GetSmsSeqSenderByBatch(manaHovNum, 1, ds, errDesc) Then
						Throw New Exception(errDesc)
					End If

					If Not ds.Tables.Contains("SmsSeqSender") OrElse ds.Tables("SmsSeqSender").Rows.Count = 0 Then
						Return True
					End If
				End If

				' עיבוד בקבוצות של 10
				Dim tasks As New List(Of Task)
				Using semaphore As New SemaphoreSlim(10, 10)
					For Each row As DataRow In ds.Tables("SmsSeqSender").Rows
						Dim currentStatus As Integer = CInt(row("Status"))

						' במצב התאוששות, מעבדים רק את מי שבסטטוס 1 או -2
						If isRecovery AndAlso (currentStatus = 1 OrElse currentStatus = -2) Then
							tasks.Add(CreateShortUrlForRecordAsync(row, manaHovNum, shovarRashut, semaphore))
						ElseIf Not isRecovery AndAlso currentStatus = 1 Then
							tasks.Add(CreateShortUrlForRecordAsync(row, manaHovNum, shovarRashut, semaphore))
						End If
					Next

					Await Task.WhenAll(tasks)
				End Using

				Return True

			Catch ex As Exception
				Throw
			End Try
		End Function

		''' <summary>
		''' יצירת Short URL לרשומה בודדת
		''' </summary>
		Private Async Function CreateShortUrlForRecordAsync(ByVal row As DataRow, ByVal manaHovNum As Integer,
															   ByVal shovarRashut As Integer, ByVal semaphore As SemaphoreSlim) As Task
			Await semaphore.WaitAsync()

			Try
				Dim shovarNum As Integer = CInt(row("shovar"))
				Dim maintz As Long = CLng(row("mainttz"))
				Dim fullUrl As String = $"{m_Params(0).SmsUrl}?rshut={shovarRashut}&maintz={maintz}&shovarnumber={shovarNum}"
				Dim shortUrl As String = ""
				Dim errDesc As String = ""
				Dim shovarId As Integer = 0

				' בדיקה אם כבר קיים Short URL
				Dim existingShortUrl As String = m_DalSms.GetShortUrlForShovar(shovarNum, maintz, errDesc)

				If Not String.IsNullOrEmpty(existingShortUrl) Then
					' כבר קיים Short URL, רק צריך לעדכן סטטוס
					m_DalSms.UpdateSmsSeqSenderStatus(shovarNum, manaHovNum, 2, errDesc)
					Return
				End If

				' יצירת Short URL חדש
				If GetShortUrl(fullUrl, shortUrl, errDesc) Then
					' שמירה ב-DB
					Using trans As SqlTransaction = m_DalSms.CreateTransaction()
						m_DalSms.AddShovarLink(shovarNum, maintz, fullUrl, shortUrl, shovarId, trans, errDesc)
						trans.Commit()
					End Using

					' עדכון סטטוס ל-2
					m_DalSms.UpdateSmsSeqSenderStatus(shovarNum, manaHovNum, 2, errDesc)
				Else
					' עדכון סטטוס ל -2
					m_DalSms.UpdateSmsSeqSenderStatus(shovarNum, manaHovNum, -2, errDesc)
				End If

			Catch ex As Exception
				' במקרה של שגיאה, מעדכנים סטטוס ל -2
				Dim errDesc As String = ""
				m_DalSms.UpdateSmsSeqSenderStatus(CInt(row("shovar")), manaHovNum, -2, errDesc)
			Finally
				semaphore.Release()
			End Try
		End Function

		''' <summary>
		''' שלב 3: שליחת הודעות SMS
		''' </summary>
		Private Async Function SendSmsBatchAsync(ByVal manaHovNum As Integer, ByVal smsType As SmsTypeEnum, ByVal isRecovery As Boolean) As Task(Of Boolean)
			Dim ds As New DsSms
			Dim errDesc As String = ""

			Try
				' קבלת רשומות לעיבוד
				If isRecovery Then
					' במצב התאוששות - מביאים כל רשומה שלא בסטטוס 3
					If Not m_DalSms.GetSmsSeqSenderByBatch(manaHovNum, Nothing, ds, errDesc) Then
						Throw New Exception(errDesc)
					End If

					If Not ds.Tables.Contains("SmsSeqSender") OrElse ds.Tables("SmsSeqSender").Rows.Count = 0 Then
						Return True
					End If
				Else
					' מצב רגיל - רק סטטוס 2
					If Not m_DalSms.GetSmsSeqSenderByBatch(manaHovNum, 2, ds, errDesc) Then
						Throw New Exception(errDesc)
					End If

					If Not ds.Tables.Contains("SmsSeqSender") OrElse ds.Tables("SmsSeqSender").Rows.Count = 0 Then
						Return True
					End If
				End If

				' שליחה מקבילית עם הגבלה של 10 threads
				Dim tasks As New List(Of Task)
				Using semaphore As New SemaphoreSlim(10, 10)
					For Each row As DataRow In ds.Tables("SmsSeqSender").Rows
						Dim currentStatus As Integer = CInt(row("Status"))

						' במצב התאוששות, מעבדים רק את מי שבסטטוס 2 או -3
						If isRecovery AndAlso (currentStatus = 2 OrElse currentStatus = -3) Then
							tasks.Add(SendSingleSmsAsync(row, manaHovNum, smsType, semaphore))
						ElseIf Not isRecovery AndAlso currentStatus = 2 Then
							tasks.Add(SendSingleSmsAsync(row, manaHovNum, smsType, semaphore))
						End If
					Next

					Await Task.WhenAll(tasks)
				End Using

				Return True

			Catch ex As Exception
				Throw
			End Try
		End Function

		''' <summary>
		''' שליחת SMS בודד
		''' </summary>
		Private Async Function SendSingleSmsAsync(ByVal row As DataRow, ByVal manaHovNum As Integer,
													 ByVal smsType As SmsTypeEnum, ByVal semaphore As SemaphoreSlim) As Task
			Await semaphore.WaitAsync()

			Try
				Dim shovarNum As Integer = CInt(row("shovar"))
				Dim maintz As Long = CLng(row("mainttz"))
				Dim phone As String = row("cell").ToString()
				Dim errDesc As String = ""
				Dim smsStatus As Integer = 0
				Dim smsStatusDesc As String = ""

				' קבלת Short URL מטבלת ShovarLink
				Dim shortUrl As String = m_DalSms.GetShortUrlForShovar(shovarNum, maintz, errDesc)

				If String.IsNullOrEmpty(shortUrl) Then
					Throw New Exception("Short URL not found for shovar " & shovarNum)
				End If

				' בניית הודעה
				Dim msg As String = BuildMessage(shortUrl)

				' שליחת SMS
				Dim wsSMS As SmsSender.WsInforu = New SmsSender.WsInforu()

				Using trans As SqlTransaction = m_DalSms.CreateTransaction()
					If wsSMS.Send(msg, New List(Of String) From {phone}, m_Params(0).SmsSender, smsStatus, errDesc) Then
						smsStatusDesc = wsSMS.GetEnumDescription(DirectCast(smsStatus, SmsSender.WsInforu.SmsStatusEnum))

						' שמירת לוג
						Dim shovarId As Integer = m_DalSms.GetShovarIdForShovar(shovarNum, maintz, errDesc)
						m_DalSms.AddSmsLog(shovarId, smsType, m_Params(0).SmsSender, msg, msg.Length, phone, smsStatus, errDesc, trans)

						trans.Commit()

						' עדכון סטטוס ל-3
						m_DalSms.UpdateSmsSeqSenderStatus(shovarNum, manaHovNum, 3, errDesc)
					Else
						' עדכון סטטוס ל -3
						m_DalSms.UpdateSmsSeqSenderStatus(shovarNum, manaHovNum, -3, errDesc)
					End If
				End Using

			Catch ex As Exception
				' במקרה של שגיאה, מעדכנים סטטוס ל -3
				Dim errDesc As String = ""
				m_DalSms.UpdateSmsSeqSenderStatus(CInt(row("shovar")), manaHovNum, -3, errDesc)
			Finally
				semaphore.Release()
			End Try
		End Function

#Region "Helper Functions"

		''' <summary>
		''' בדיקת תקינות מספר טלפון
		''' </summary>
		Private Function IsValidPhoneNumber(ByVal phone As String) As Boolean
			If String.IsNullOrWhiteSpace(phone) Then Return False

			' הסרת תווים לא רלוונטיים
			phone = phone.Replace("-", "").Replace(" ", "").Replace("(", "").Replace(")", "")

			' בדיקת אורך
			If phone.Length < 9 OrElse phone.Length > 10 Then Return False

			' בדיקה שכל התווים הם ספרות
			For Each c As Char In phone
				If Not Char.IsDigit(c) Then Return False
			Next

			' בדיקת קידומת
			If phone.Length = 10 AndAlso Not phone.StartsWith("0") Then Return False
			If phone.Length = 9 AndAlso Not phone.StartsWith("5") Then Return False

			Return True
		End Function




#End Region

#End Region

	End Class

End Namespace