Imports System.Runtime.InteropServices
Imports CreditGuardClientExecuter
Imports CreditGuardClientExecuter.ClsExecuter
Imports CreditGuardClientExecuter.ResultData
Imports CreditGuardGlobals

<ComVisible(True), ClassInterface(ClassInterfaceType.AutoDual)>
Partial Public Class ClsMain

#Region "Globals"

	Private Enum EnmResultCategory
		CGSuccess = 1
		CGFailure
		WebPageUserCancelled
		WebPageTimeoutExpired
		InvalidRequestProperties
		GeneralRuntimeError
	End Enum

	Private Const MessagesMaxLength_ As UShort = 5000

	Private Const SuccessCode_ As Integer = 0

	Private Const WebPageCancelledCode_ As Integer = -1

	Private Const WebPageTimedOutCode_ As Integer = -2

	Private Const WebPageSuccessText_ As String = "עסקה תקינה"

	Private Const WebPageCancelledText_ As String = "הפעולה בוטלה על-ידי המשתמש"

	Private Const WebPageTimedOutText_ As String = "תם הזמן הקצוב לעסקה"

	Private Const EmptyRequestProperties_ As Integer = -10

	Private Const InvalidRequestProperties_ As Integer = -11

	Private Const InvalidCombinationRequestProperties_ As Integer = -12

	Private Const EmptyAndInvalidRequestProperties_ As Integer = -13

	Private Const EmptyAndInvalidCombinationRequestProperties_ As Integer = -14

	Private Const InvalidAndInvalidCombinationRequestProperties_ As Integer = -15

	Private Const EmptyAndInvalidAndInvalidCombinationRequestProperties_ As Integer = -16

	Private Const RuntimeError_ As Integer = -20


	Private Shared Function CreateCGClient(CgMerchantID As Integer, CgApiUrl As String, CgApiUser As String, CgApiPassword As String) As ClsExecuter
		Return New ClsExecuter(New Types.PInteger(CgMerchantID), CgApiUrl, CgApiUser, CgApiPassword)
	End Function

	Private Shared Function CreateCGClient(OdbcName As String, Optional DbUser As String = Nothing, Optional DbPassword As String = Nothing) As ClsExecuter
		Return New ClsExecuter(OdbcName, DbUser, DbPassword)
	End Function

	Private Shared Function MyLogLevelToExecuterLogLevel(LogLevel As CreditGuardClient.EnmBrowserListenerLogLevel) As ClsExecuter.EnmBrowserListenerLogLevel
		Return Microsoft.VisualBasic.Switch(LogLevel = CreditGuardClient.EnmBrowserListenerLogLevel.None, ClsExecuter.EnmBrowserListenerLogLevel.None,
											LogLevel = CreditGuardClient.EnmBrowserListenerLogLevel.Basic, ClsExecuter.EnmBrowserListenerLogLevel.Basic,
											LogLevel = CreditGuardClient.EnmBrowserListenerLogLevel.Debug, ClsExecuter.EnmBrowserListenerLogLevel.Debug)
	End Function

	Private Shared Function ExecuterResultToMyResult(ExecuterResult As ClsResult.EnmResult) As Integer
		Return Microsoft.VisualBasic.Switch(ExecuterResult = ClsResult.EnmResult.Success, SuccessCode_,
											ExecuterResult = ClsResult.EnmResult.EmptyRequestProperties, EmptyRequestProperties_,
											ExecuterResult = ClsResult.EnmResult.InvalidRequestProperties, InvalidRequestProperties_,
											ExecuterResult = ClsResult.EnmResult.InvalidCombinationRequestProperties, InvalidCombinationRequestProperties_,
											ExecuterResult = ClsResult.EnmResult.EmptyAndInvalidRequestProperties, EmptyAndInvalidRequestProperties_,
											ExecuterResult = ClsResult.EnmResult.EmptyAndInvalidCombinationRequestProperties, EmptyAndInvalidCombinationRequestProperties_,
											ExecuterResult = ClsResult.EnmResult.InvalidAndInvalidCombinationRequestProperties, InvalidAndInvalidCombinationRequestProperties_,
											ExecuterResult = ClsResult.EnmResult.EmptyAndInvalidAndInvalidCombinationRequestProperties, EmptyAndInvalidAndInvalidCombinationRequestProperties_)
	End Function

	Private Shared Sub WebPageResultToMyResult(ExecuterResult As ClsResult, ByRef ResultCategory As EnmResultCategory, ByRef MyCode As Integer, ByRef MyText As String, ByRef ActualData As Object)
		Dim WebPageResult As EnmWebPageTransactionResult

		MyCode = ExecuterResultToMyResult(ExecuterResult.ResultCode)

		If MyCode = SuccessCode_ Then
			If TypeOf ExecuterResult.ResultData Is ClsWebPageTokenization Then
				ActualData = DirectCast(ExecuterResult.ResultData, ClsWebPageTokenization)
			ElseIf TypeOf ExecuterResult.ResultData Is ClsWebPageVerification Then
				ActualData = DirectCast(ExecuterResult.ResultData, ClsWebPageVerification)
			ElseIf TypeOf ExecuterResult.ResultData Is ClsWebPageHKRegistration Then
				ActualData = DirectCast(ExecuterResult.ResultData, ClsWebPageHKRegistration)
			ElseIf TypeOf ExecuterResult.ResultData Is ClsWebPagePayment Then
				ActualData = DirectCast(ExecuterResult.ResultData, ClsWebPagePayment)
			End If

			WebPageResult = ActualData.WebPageTransactionResult

			Select Case WebPageResult
				Case EnmWebPageTransactionResult.Succeeded
					ResultCategory = EnmResultCategory.CGSuccess
					MyCode = SuccessCode_
					MyText = WebPageSuccessText_
				Case EnmWebPageTransactionResult.Cancelled
					ResultCategory = EnmResultCategory.WebPageUserCancelled
					MyCode = WebPageCancelledCode_
					MyText = WebPageCancelledText_
				Case EnmWebPageTransactionResult.TimeoutExpired
					ResultCategory = EnmResultCategory.WebPageTimeoutExpired
					MyCode = WebPageTimedOutCode_
					MyText = WebPageTimedOutText_
				Case EnmWebPageTransactionResult.Failed
					ResultCategory = EnmResultCategory.CGFailure
					MyCode = Integer.Parse(ActualData.ResultCode)
					MyText = ActualData.ResultInnerText
			End Select
		Else
			ResultCategory = EnmResultCategory.InvalidRequestProperties
			MyText = ExecuterResult.ResultText
		End If
	End Sub

	Private Shared Sub APIResultToMyResult(ExecuterResult As ClsResult, ByRef ResultCategory As EnmResultCategory, ByRef MyCode As Integer, ByRef MyText As String, ByRef ActualData As Object)
		MyCode = ExecuterResultToMyResult(ExecuterResult.ResultCode)

		If MyCode = SuccessCode_ Then
			If TypeOf ExecuterResult.ResultData Is ClsAPIWebPageUrl Then
				ActualData = DirectCast(ExecuterResult.ResultData, ClsAPIWebPageUrl)
			ElseIf TypeOf ExecuterResult.ResultData Is ClsAPITokenization Then
				ActualData = DirectCast(ExecuterResult.ResultData, ClsAPITokenization)
			ElseIf TypeOf ExecuterResult.ResultData Is ClsAPIVerification Then
				ActualData = DirectCast(ExecuterResult.ResultData, ClsAPIVerification)
			ElseIf TypeOf ExecuterResult.ResultData Is ClsAPIHKRegistration Then
				ActualData = DirectCast(ExecuterResult.ResultData, ClsAPIHKRegistration)
			ElseIf TypeOf ExecuterResult.ResultData Is ClsAPIPayment Then
				ActualData = DirectCast(ExecuterResult.ResultData, ClsAPIPayment)
			ElseIf TypeOf ExecuterResult.ResultData Is ClsAPIDealCancelation Then
				ActualData = DirectCast(ExecuterResult.ResultData, ClsAPIDealCancelation)
			ElseIf TypeOf ExecuterResult.ResultData Is ClsAPITransactionInquiry Then
				ActualData = DirectCast(ExecuterResult.ResultData, ClsAPITransactionInquiry)
			End If

			MyCode = Integer.Parse(ActualData.ResultCode)
			MyText = ActualData.ResultInnerText
			If MyCode = 0 Then
				ResultCategory = EnmResultCategory.CGSuccess
			Else
				ResultCategory = EnmResultCategory.CGFailure
			End If
		Else
			ResultCategory = EnmResultCategory.InvalidRequestProperties
			MyText = ExecuterResult.ResultText
		End If
	End Sub

	Private Shared Function CreateRuntimeErrorReturn(Ex As Exception) As ClsReturn
		With ExtractMustInnerException(Ex)
			Return New ClsReturn(RuntimeError_, "אירעה שגיאה", Nothing,
								 New ClsError(.HResult, $"Message: { .Message}{vbCrLf}{vbCrLf}Stack-Trace: { .StackTrace}", .Message))
		End With
	End Function

	Private Shared Function MyCreditTypeToExecuterCreditType(CreditType As CreditGuardClient.EnmCreditType) As EnmCreditType
		Return Microsoft.VisualBasic.Switch(CreditType = CreditGuardClient.EnmCreditType.RegularCredit, EnmCreditType.RegularCredit,
											CreditType = CreditGuardClient.EnmCreditType.IsraCredit, EnmCreditType.IsraCredit,
											CreditType = CreditGuardClient.EnmCreditType.AdHock, EnmCreditType.AdHock,
											CreditType = CreditGuardClient.EnmCreditType.ClubDeal, EnmCreditType.ClubDeal,
											CreditType = CreditGuardClient.EnmCreditType.SpecialAlpha, EnmCreditType.SpecialAlpha,
											CreditType = CreditGuardClient.EnmCreditType.SpecialCredit, EnmCreditType.SpecialCredit,
											CreditType = CreditGuardClient.EnmCreditType.Payments, EnmCreditType.Payments,
											CreditType = CreditGuardClient.EnmCreditType.PaymentsClub, EnmCreditType.PaymentsClub)
	End Function

	Private Shared Function MyWebPageTransactionTypeToExecuterType(TransactionType As CreditGuardClient.EnmWebPageTransactionType) As ClsExecuter.EnmTransactionType
		Return Microsoft.VisualBasic.Switch(TransactionType = CreditGuardClient.EnmWebPageTransactionType.Tokenization, ClsExecuter.EnmTransactionType.Tokenization,
											TransactionType = CreditGuardClient.EnmWebPageTransactionType.Verification, ClsExecuter.EnmTransactionType.Verification,
											TransactionType = CreditGuardClient.EnmWebPageTransactionType.HKRegistration, ClsExecuter.EnmTransactionType.HKRegistration,
											TransactionType = CreditGuardClient.EnmWebPageTransactionType.Payment, ClsExecuter.EnmTransactionType.Payment)
	End Function

	Private Shared Function WebPageTransactionResultToMyResult(TransactionResult As EnmWebPageTransactionResult) As CreditGuardClient.EnmWebPageTransactionResult
		Return Microsoft.VisualBasic.Switch(TransactionResult = EnmWebPageTransactionResult.Succeeded, CreditGuardClient.EnmWebPageTransactionResult.Succeeded,
											TransactionResult = EnmWebPageTransactionResult.Cancelled, CreditGuardClient.EnmWebPageTransactionResult.UserCancelled,
											TransactionResult = EnmWebPageTransactionResult.TimeoutExpired, CreditGuardClient.EnmWebPageTransactionResult.TimeOutExpired,
											TransactionResult = EnmWebPageTransactionResult.Failed, CreditGuardClient.EnmWebPageTransactionResult.Failed)
	End Function

#End Region

#Region "InitAndShutdown"

	Function InitializeCreditGuardBrowser(<MarshalAs(UnmanagedType.BStr)> BrowserListenerPath As String,
										 <MarshalAs(UnmanagedType.I4)> Optional CallerHandle As Integer = 0,
										 <MarshalAs(UnmanagedType.U1)> Optional LogLevel As CreditGuardClient.EnmBrowserListenerLogLevel = CreditGuardClient.EnmBrowserListenerLogLevel.None,
										 <MarshalAs(UnmanagedType.BStr)> Optional OdbcName As String = "",
										 <MarshalAs(UnmanagedType.BStr)> Optional DbUser As String = "",
										 <MarshalAs(UnmanagedType.BStr)> Optional DbPassword As String = "",
										 <MarshalAs(UnmanagedType.BStr)> Optional ApplicationUser As String = "") As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

		Dim ErrObj As ClsError = Nothing

		Try
			ClsExecuter.InitializeBrowser(BrowserListenerPath, MessagesMaxLength_, CallerHandle, MyLogLevelToExecuterLogLevel(LogLevel), OdbcName, DbUser, DbPassword, ApplicationUser)
		Catch ex As Exception
			Dim InnerEx As Exception = ExtractMustInnerException(ex)
			ErrObj = New ClsError(InnerEx.HResult, $"Message: {InnerEx.Message}{vbCrLf}{vbCrLf}Stack-Trace: {InnerEx.StackTrace}", InnerEx.Message)
		End Try

		Return New ClsReturn(If(ErrObj Is Nothing, 0, 1), If(ErrObj Is Nothing, "", "אירעה שגיאה"), Nothing, ErrObj)
	End Function

	Sub ShutDownCreditGuardBrowser()
		ClsExecuter.ShutDownBrowser()
	End Sub

#End Region

#Region "WebPageTransactions"

#Region "WebPageTransactions.Tokenization"

	Function CreditGuardWebPageTokenize_FullParams(<MarshalAs(UnmanagedType.I4)> CgMerchantID As Integer,
												   <MarshalAs(UnmanagedType.BStr)> CgApiUrl As String,
												   <MarshalAs(UnmanagedType.BStr)> CgApiUser As String,
												   <MarshalAs(UnmanagedType.BStr)> CgApiPassword As String,
												   <MarshalAs(UnmanagedType.BStr)> TerminalNumber As String) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

		Dim Ret As ClsReturn
		Dim Client As ClsExecuter
		Dim Result As ClsResult

		Try
			Client = CreateCGClient(CgMerchantID, CgApiUrl, CgApiUser, CgApiPassword)
			Result = Client.WebPageTokenize(TerminalNumber)
			Ret = CreateWebPageTokenizeReturn(Result)
		Catch Ex As Exception
			Ret = CreateRuntimeErrorReturn(Ex)
		End Try

		Return Ret
	End Function

	Function CreditGuardWebPageTokenize_UsingOdbc(<MarshalAs(UnmanagedType.BStr)> OdbcName As String,
												  <MarshalAs(UnmanagedType.BStr)> Optional DbUser As String = "",
												  <MarshalAs(UnmanagedType.BStr)> Optional DbPassword As String = "") As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

		Dim Ret As ClsReturn
		Dim Client As ClsExecuter
		Dim Result As ClsResult

		Try
			Client = CreateCGClient(OdbcName, DbUser, DbPassword)
			Result = Client.WebPageTokenize()
			Ret = CreateWebPageTokenizeReturn(Result)
		Catch Ex As Exception
			Ret = CreateRuntimeErrorReturn(Ex)
		End Try

		Return Ret
	End Function

	Private Function CreateWebPageTokenizeReturn(CGResult As ClsResult) As ClsReturn
		Dim ResultCategory As EnmResultCategory
		Dim StatusCode As Integer
		Dim StatusText As String = Nothing
		Dim ResultProperties As ClsWebPageTokenization = Nothing

		WebPageResultToMyResult(CGResult, ResultCategory, StatusCode, StatusText, ResultProperties)
		If ResultCategory = EnmResultCategory.CGFailure AndAlso Not String.IsNullOrEmpty(ResultProperties.CardToken) Then
			StatusCode = SuccessCode_
		End If
		If StatusCode = SuccessCode_ Then
			StatusText = ResultProperties.CardToken
		End If

		Return New ClsReturn(StatusCode, StatusText, Nothing, Nothing)
	End Function

#End Region

#Region "WebPageTransactions.Verification"
	Function CreditGuardWebPageVerify_FullParams(<MarshalAs(UnmanagedType.I4)> CgMerchantID As Integer,
												 <MarshalAs(UnmanagedType.BStr)> CgApiUrl As String,
												 <MarshalAs(UnmanagedType.BStr)> CgApiUser As String,
												 <MarshalAs(UnmanagedType.BStr)> CgApiPassword As String,
												 <MarshalAs(UnmanagedType.BStr)> TerminalNumber As String,
												 <MarshalAs(UnmanagedType.BStr)> Optional Expiration As String = "",
												 <MarshalAs(UnmanagedType.U4)> Optional OwnerID As UInteger = 0) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

		Dim Ret As ClsReturn
		Dim Client As ClsExecuter
		Dim Result As ClsResult

		Try
			Client = CreateCGClient(CgMerchantID, CgApiUrl, CgApiUser, CgApiPassword)
			Result = Client.WebPageVerify(TerminalNumber, Expiration, OwnerID)
			Ret = CreateWebPageVerifyReturn(Result)
		Catch Ex As Exception
			Ret = CreateRuntimeErrorReturn(Ex)
		End Try

		Return Ret
	End Function

	Function CreditGuardWebPageVerify_UsingODBC(<MarshalAs(UnmanagedType.BStr)> OdbcName As String,
												<MarshalAs(UnmanagedType.BStr)> DbUser As String,
												<MarshalAs(UnmanagedType.BStr)> DbPassword As String,
												<MarshalAs(UnmanagedType.BStr)> TerminalNumber As String,
												<MarshalAs(UnmanagedType.BStr)> Optional Expiration As String = "",
												<MarshalAs(UnmanagedType.U4)> Optional OwnerID As UInteger = 0) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

		Dim Ret As ClsReturn
		Dim Client As ClsExecuter
		Dim Result As ClsResult

		Try
			Client = CreateCGClient(OdbcName, DbUser, DbPassword)
			Result = Client.WebPageVerify(TerminalNumber, Expiration, OwnerID)
			Ret = CreateWebPageVerifyReturn(Result)
		Catch Ex As Exception
			Ret = CreateRuntimeErrorReturn(Ex)
		End Try

		Return Ret
	End Function

	Private Function CreateWebPageVerifyReturn(CGResult As ClsResult) As ClsReturn
		Dim ResultCategory As EnmResultCategory
		Dim StatusCode As Integer
		Dim StatusText As String = Nothing
		Dim ResultProperties As ClsWebPageVerification = Nothing
		Dim Data As CreditGuardClient.ClsWebPageVerificationResult = Nothing

		WebPageResultToMyResult(CGResult, ResultCategory, StatusCode, StatusText, ResultProperties)
		If ResultCategory = EnmResultCategory.CGSuccess OrElse ResultCategory = EnmResultCategory.CGFailure Then
			Data = New CreditGuardClient.ClsWebPageVerificationResult(ResultProperties.ExtendedResultCode, ResultProperties.ExtendedResultInnerText, ResultProperties.ExtendedResultOuterText,
																	  ResultProperties.CardToken, ResultProperties.ResponseAuthorizationNumber, ResultProperties.Expiration,
																	  ResultProperties.OwnerID, ResultProperties.CardType, ResultProperties.CardIssuer, ResultProperties.CardBrand)
		End If

		Return New ClsReturn(StatusCode, StatusText, Data, Nothing)
	End Function

#End Region

#Region "WebPageTransactions.HKRegitration"

	Function CreditGuardWebPageRegisterHK_FullParams(<MarshalAs(UnmanagedType.I4)> CgMerchantID As Integer,
													 <MarshalAs(UnmanagedType.BStr)> CgApiUrl As String,
													 <MarshalAs(UnmanagedType.BStr)> CgApiUser As String,
													 <MarshalAs(UnmanagedType.BStr)> CgApiPassword As String,
													 <MarshalAs(UnmanagedType.BStr)> TerminalNumber As String,
													 <MarshalAs(UnmanagedType.BStr)> Optional Expiration As String = "",
													 <MarshalAs(UnmanagedType.U4)> Optional OwnerID As UInteger = 0) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

		Dim Ret As ClsReturn
		Dim Client As ClsExecuter
		Dim Result As ClsResult

		Try
			Client = CreateCGClient(CgMerchantID, CgApiUrl, CgApiUser, CgApiPassword)
			Result = Client.WebPageRegisterHK(TerminalNumber, Expiration, OwnerID)
			Ret = CreateWebPageRegisterHKReturn(Result)
		Catch Ex As Exception
			Ret = CreateRuntimeErrorReturn(Ex)
		End Try

		Return Ret
	End Function

	Function CreditGuardWebPageRegisterHK_UsingODBC(<MarshalAs(UnmanagedType.BStr)> OdbcName As String,
													<MarshalAs(UnmanagedType.BStr)> DbUser As String,
													<MarshalAs(UnmanagedType.BStr)> DbPassword As String,
													<MarshalAs(UnmanagedType.BStr)> TerminalNumber As String,
													<MarshalAs(UnmanagedType.BStr)> Optional Expiration As String = "",
													<MarshalAs(UnmanagedType.U4)> Optional OwnerID As UInteger = 0) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

		Dim Ret As ClsReturn
		Dim Client As ClsExecuter
		Dim Result As ClsResult

		Try
			Client = CreateCGClient(OdbcName, DbUser, DbPassword)
			Result = Client.WebPageRegisterHK(TerminalNumber, Expiration, OwnerID)
			Ret = CreateWebPageRegisterHKReturn(Result)
		Catch Ex As Exception
			Ret = CreateRuntimeErrorReturn(Ex)
		End Try

		Return Ret
	End Function

	Private Function CreateWebPageRegisterHKReturn(CGResult As ClsResult) As ClsReturn
		Dim ResultCategory As EnmResultCategory
		Dim StatusCode As Integer
		Dim StatusText As String = Nothing
		Dim ResultProperties As ClsWebPageHKRegistration = Nothing
		Dim Data As CreditGuardClient.ClsWebPageHKRegistrationResult = Nothing

		WebPageResultToMyResult(CGResult, ResultCategory, StatusCode, StatusText, ResultProperties)
		If ResultCategory = EnmResultCategory.CGSuccess OrElse ResultCategory = EnmResultCategory.CGFailure Then
			Data = New CreditGuardClient.ClsWebPageHKRegistrationResult(ResultProperties.ExtendedResultCode, ResultProperties.ExtendedResultInnerText, ResultProperties.ExtendedResultOuterText,
																		ResultProperties.CardToken, ResultProperties.ResponseAuthorizationNumber, ResultProperties.Expiration, ResultProperties.OwnerID,
																		ResultProperties.CardType, ResultProperties.CardIssuer, ResultProperties.CardBrand)
		End If

		Return New ClsReturn(StatusCode, StatusText, Data, Nothing)
	End Function

#End Region

#Region "WebPageTransactions.Payment"

	Function CreditGuardWebPagePay_FullParams(<MarshalAs(UnmanagedType.I4)> CgMerchantID As Integer,
											  <MarshalAs(UnmanagedType.BStr)> CgApiUrl As String,
											  <MarshalAs(UnmanagedType.BStr)> CgApiUser As String,
											  <MarshalAs(UnmanagedType.BStr)> CgApiPassword As String,
											  <MarshalAs(UnmanagedType.BStr)> XField As String,
											  <MarshalAs(UnmanagedType.BStr)> TerminalNumber As String,
											  <MarshalAs(UnmanagedType.U1)> CreditType As CreditGuardClient.EnmCreditType,
											  <MarshalAs(UnmanagedType.R8)> TotalSum As Double,
											  <MarshalAs(UnmanagedType.U1)> NumberPayments As Byte,
											  <MarshalAs(UnmanagedType.Bool)> DefaultPaymentsSpread As Boolean,
											  <MarshalAs(UnmanagedType.R8)> FirstPaymentSum As Double,
											  <MarshalAs(UnmanagedType.R8)> PeriodicalPaymentSum As Double,
											  <MarshalAs(UnmanagedType.BStr)> Optional Authorization As String = "",
											  <MarshalAs(UnmanagedType.BStr)> Optional Expiration As String = "",
											  <MarshalAs(UnmanagedType.U4)> Optional OwnerID As UInteger = 0) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

		Dim Ret As ClsReturn
		Dim Client As ClsExecuter
		Dim Result As ClsResult

		Try
			Client = CreateCGClient(CgMerchantID, CgApiUrl, CgApiUser, CgApiPassword)
			Result = Client.WebPagePay(XField, TerminalNumber, MyCreditTypeToExecuterCreditType(CreditType), TotalSum, NumberPayments, DefaultPaymentsSpread, FirstPaymentSum, PeriodicalPaymentSum, Authorization, Expiration, OwnerID)
			Ret = CreateWebPagePayReturn(Result)
		Catch Ex As Exception
			Ret = CreateRuntimeErrorReturn(Ex)
		End Try

		Return Ret
	End Function

	Function CreditGuardWebPagePay_UsingODBC(<MarshalAs(UnmanagedType.BStr)> OdbcName As String,
											 <MarshalAs(UnmanagedType.BStr)> DbUser As String,
											 <MarshalAs(UnmanagedType.BStr)> DbPassword As String,
											 <MarshalAs(UnmanagedType.BStr)> XField As String,
											 <MarshalAs(UnmanagedType.BStr)> TerminalNumber As String,
											 <MarshalAs(UnmanagedType.U1)> CreditType As CreditGuardClient.EnmCreditType,
											 <MarshalAs(UnmanagedType.R8)> TotalSum As Double,
											 <MarshalAs(UnmanagedType.U1)> NumberPayments As Byte,
											 <MarshalAs(UnmanagedType.Bool)> DefaultPaymentsSpread As Boolean,
											 <MarshalAs(UnmanagedType.R8)> FirstPaymentSum As Double,
											 <MarshalAs(UnmanagedType.R8)> PeriodicalPaymentSum As Double,
											 <MarshalAs(UnmanagedType.BStr)> Optional Authorization As String = "",
											 <MarshalAs(UnmanagedType.BStr)> Optional Expiration As String = "",
											 <MarshalAs(UnmanagedType.U4)> Optional OwnerID As UInteger = 0,
											 <MarshalAs(UnmanagedType.U4)> Optional TerminalSlaveNumber As UInteger = 0,
											 <MarshalAs(UnmanagedType.U4)> Optional Mid As UInteger = 0) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

		Dim Ret As ClsReturn
		Dim Client As ClsExecuter
		Dim Result As ClsResult

		Try
			Client = CreateCGClient(OdbcName, DbUser, DbPassword)
			Result = Client.WebPagePay(XField, TerminalNumber, MyCreditTypeToExecuterCreditType(CreditType), TotalSum, NumberPayments, DefaultPaymentsSpread, FirstPaymentSum, PeriodicalPaymentSum, Authorization, Expiration, OwnerID, TerminalSlaveNumber, Mid)
			Ret = CreateWebPagePayReturn(Result)
		Catch Ex As Exception
			Ret = CreateRuntimeErrorReturn(Ex)
		End Try

		Return Ret
	End Function

	Private Function CreateWebPagePayReturn(CGResult As ClsResult) As ClsReturn
		Dim ResultCategory As EnmResultCategory
		Dim StatusCode As Integer
		Dim StatusText As String = Nothing
		Dim ResultProperties As ClsWebPagePayment = Nothing
		Dim Data As CreditGuardClient.ClsWebPagePaymentResult = Nothing

		WebPageResultToMyResult(CGResult, ResultCategory, StatusCode, StatusText, ResultProperties)
		If ResultCategory = EnmResultCategory.CGSuccess OrElse ResultCategory = EnmResultCategory.CGFailure Then
			Data = New CreditGuardClient.ClsWebPagePaymentResult(ResultProperties.ShvaUID, ResultProperties.ExtendedResultCode, ResultProperties.ExtendedResultInnerText, ResultProperties.ExtendedResultOuterText,
																 ResultProperties.ResponseAuthorizationNumber, ResultProperties.CardToken, ResultProperties.Expiration, ResultProperties.OwnerID,
																 ResultProperties.CardType, ResultProperties.CardAcquirer, ResultProperties.CardIssuer, ResultProperties.CardBrand, ResultProperties.ShovarNumber)
		End If

		Return New ClsReturn(StatusCode, StatusText, Data, Nothing)
	End Function

#End Region

#End Region

#Region "APITransactions"

#Region "APITransactions.WebPageUrl"

	Function CreditGuardAPIGenerateWebPageUrl_FullParams(<MarshalAs(UnmanagedType.U4)> CgMid As Integer,
														 <MarshalAs(UnmanagedType.BStr)> CgApiUrl As String,
														 <MarshalAs(UnmanagedType.BStr)> CgApiUser As String,
														 <MarshalAs(UnmanagedType.BStr)> CgApiPassword As String,
														 <MarshalAs(UnmanagedType.BStr)> ProcessUniqueID As String,
														 <MarshalAs(UnmanagedType.BStr)> TerminalNumber As String,
														 <MarshalAs(UnmanagedType.BStr)> XField As String,
														 <MarshalAs(UnmanagedType.BStr)> SuccessLandingUrl As String,
														 <MarshalAs(UnmanagedType.BStr)> CancelLandingUrl As String,
														 <MarshalAs(UnmanagedType.BStr)> FailureLandingUrl As String,
														 <MarshalAs(UnmanagedType.U1)> TransactionType As CreditGuardClient.EnmWebPageTransactionType,
														 <MarshalAs(UnmanagedType.U1)> CreditType As CreditGuardClient.EnmCreditType,
														 <MarshalAs(UnmanagedType.R8)> TotalSum As Double,
														 <MarshalAs(UnmanagedType.U1)> NumberPayments As Byte,
														 <MarshalAs(UnmanagedType.R8)> FirstPaymentSum As Double,
														 <MarshalAs(UnmanagedType.R8)> PeriodicalPaymentSum As Double,
														 <MarshalAs(UnmanagedType.BStr)> Optional AuthorizationNumber As String = "",
														 <MarshalAs(UnmanagedType.BStr)> Optional Expiration As String = "",
														 <MarshalAs(UnmanagedType.U4)> Optional OwnerID As UInteger = 0) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

		Dim Ret As ClsReturn
		Dim Client As ClsExecuter
		Dim Result As ClsResult

		Try
			Client = CreateCGClient(CgMid, CgApiUrl, CgApiUser, CgApiPassword)
			Result = Client.APIGenerateWebPageURL(ProcessUniqueID, TerminalNumber, XField, SuccessLandingUrl, CancelLandingUrl, FailureLandingUrl, MyWebPageTransactionTypeToExecuterType(TransactionType),
													MyCreditTypeToExecuterCreditType(CreditType), TotalSum, NumberPayments, False, FirstPaymentSum, PeriodicalPaymentSum, AuthorizationNumber, Expiration, OwnerID)
			Ret = CreateAPIWebPageUrlGenerationReturn(Result)
		Catch Ex As Exception
			Ret = CreateRuntimeErrorReturn(Ex)
		End Try

		Return Ret
	End Function

	Function CreditGuardAPIGenerateWebPageUrl_UsingODBC(<MarshalAs(UnmanagedType.BStr)> OdbcName As String,
														<MarshalAs(UnmanagedType.BStr)> DbUser As String,
														<MarshalAs(UnmanagedType.BStr)> DbPassword As String,
														<MarshalAs(UnmanagedType.BStr)> ProcessUniqueID As String,
														<MarshalAs(UnmanagedType.BStr)> TerminalNumber As String,
														<MarshalAs(UnmanagedType.BStr)> XField As String,
														<MarshalAs(UnmanagedType.BStr)> SuccessLandingUrl As String,
														<MarshalAs(UnmanagedType.BStr)> CancelLandingUrl As String,
														<MarshalAs(UnmanagedType.BStr)> FailureLandingUrl As String,
														<MarshalAs(UnmanagedType.U1)> TransactionType As CreditGuardClient.EnmWebPageTransactionType,
														<MarshalAs(UnmanagedType.U1)> CreditType As CreditGuardClient.EnmCreditType,
														<MarshalAs(UnmanagedType.R8)> TotalSum As Double,
														<MarshalAs(UnmanagedType.U1)> NumberPayments As Byte,
														<MarshalAs(UnmanagedType.R8)> FirstPaymentSum As Double,
														<MarshalAs(UnmanagedType.R8)> PeriodicalPaymentSum As Double,
														<MarshalAs(UnmanagedType.BStr)> Optional AuthorizationNumber As String = "",
														<MarshalAs(UnmanagedType.BStr)> Optional Expiration As String = "",
														<MarshalAs(UnmanagedType.U4)> Optional OwnerID As UInteger = 0) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

		Dim Ret As ClsReturn
		Dim Client As ClsExecuter
		Dim Result As ClsResult

		Try
			Client = CreateCGClient(OdbcName, DbUser, DbPassword)
			Result = Client.APIGenerateWebPageURL(ProcessUniqueID, TerminalNumber, XField, SuccessLandingUrl, CancelLandingUrl, FailureLandingUrl, MyWebPageTransactionTypeToExecuterType(TransactionType),
													MyCreditTypeToExecuterCreditType(CreditType), TotalSum, NumberPayments, False, FirstPaymentSum, PeriodicalPaymentSum, AuthorizationNumber, Expiration, OwnerID)
			Ret = CreateAPIWebPageUrlGenerationReturn(Result)
		Catch Ex As Exception
			Ret = CreateRuntimeErrorReturn(Ex)
		End Try

		Return Ret
	End Function

	Private Function CreateAPIWebPageUrlGenerationReturn(CGResult As ClsResult) As ClsReturn
		Dim ResultCategory As EnmResultCategory
		Dim StatusCode As Integer
		Dim StatusText As String = Nothing
		Dim ResultProperties As ClsAPIWebPageUrl = Nothing
		Dim Data As CreditGuardClient.ClsAPIWebPageUrlGenerationResult = Nothing

		APIResultToMyResult(CGResult, ResultCategory, StatusCode, StatusText, ResultProperties)
		If ResultCategory = EnmResultCategory.CGSuccess OrElse ResultCategory = EnmResultCategory.CGFailure Then
			Data = New CreditGuardClient.ClsAPIWebPageUrlGenerationResult(ResultProperties.ExtendedResultCode, ResultProperties.ExtendedResultInnerText, ResultProperties.ExtendedResultOuterText, ResultProperties.WebPageToken, ResultProperties.WebPageUrl)
		End If

		Return New ClsReturn(StatusCode, StatusText, Data, Nothing)
	End Function

#End Region

#Region "APITransactions.Cancelation"

	Function CreditGuardAPICancelDeal_FullParams(<MarshalAs(UnmanagedType.BStr)> CgApiUrl As String,
												 <MarshalAs(UnmanagedType.BStr)> CgApiUser As String,
												 <MarshalAs(UnmanagedType.BStr)> CgApiPassword As String,
												 <MarshalAs(UnmanagedType.BStr)> TerminalNumber As String,
												 <MarshalAs(UnmanagedType.BStr)> XField As String,
												 <MarshalAs(UnmanagedType.R8)> TotalSum As Double) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

		Dim Ret As ClsReturn
		Dim Client As ClsExecuter
		Dim Result As ClsResult

		Try
			Client = CreateCGClient(1, CgApiUrl, CgApiUser, CgApiPassword)
			Result = Client.APICancelDeal(TerminalNumber, XField, TotalSum)
			Ret = CreateAPIDealCancelationReturn(Result)
		Catch Ex As Exception
			Ret = CreateRuntimeErrorReturn(Ex)
		End Try

		Return Ret
	End Function

	Function CreditGuardAPICancelDeal_UsingODBC(<MarshalAs(UnmanagedType.BStr)> OdbcName As String,
												<MarshalAs(UnmanagedType.BStr)> DbUser As String,
												<MarshalAs(UnmanagedType.BStr)> DbPassword As String,
												<MarshalAs(UnmanagedType.BStr)> TerminalNumber As String,
												<MarshalAs(UnmanagedType.BStr)> XField As String,
												<MarshalAs(UnmanagedType.R8)> TotalSum As Double) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

		Dim Ret As ClsReturn
		Dim Client As ClsExecuter
		Dim Result As ClsResult

		Try
			Client = CreateCGClient(OdbcName, DbUser, DbPassword)
			Result = Client.APICancelDeal(TerminalNumber, XField, TotalSum)
			Ret = CreateAPIDealCancelationReturn(Result)
		Catch Ex As Exception
			Ret = CreateRuntimeErrorReturn(Ex)
		End Try

		Return Ret
	End Function

	Private Function CreateAPIDealCancelationReturn(CGResult As ClsResult) As ClsReturn
		Dim ResultCategory As EnmResultCategory
		Dim StatusCode As Integer
		Dim StatusText As String = Nothing
		Dim ResultProperties As ClsAPIDealCancelation = Nothing

		APIResultToMyResult(CGResult, ResultCategory, StatusCode, StatusText, ResultProperties)

		Return New ClsReturn(StatusCode, StatusText, Nothing, Nothing)
	End Function

#End Region

#Region "APITransactions.Tokenization"

	Function CreditGuardAPITokenize_FullParams(<MarshalAs(UnmanagedType.BStr)> CgApiUrl As String,
											   <MarshalAs(UnmanagedType.BStr)> CgApiUser As String,
											   <MarshalAs(UnmanagedType.BStr)> CgApiPassword As String,
											   <MarshalAs(UnmanagedType.BStr)> CardNumber As String,
											   <MarshalAs(UnmanagedType.BStr)> Expiration As String) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

		Dim Ret As ClsReturn
		Dim Client As ClsExecuter
		Dim Result As ClsResult

		Try
			Client = CreateCGClient(1, CgApiUrl, CgApiUser, CgApiPassword)
			Result = Client.APITokenize(CardNumber, Expiration)
			Ret = CreateAPITokenizationReturn(Result)
		Catch Ex As Exception
			Ret = CreateRuntimeErrorReturn(Ex)
		End Try

		Return Ret
	End Function

	Function CreditGuardAPITokenize_UsingODBC(<MarshalAs(UnmanagedType.BStr)> OdbcName As String,
											  <MarshalAs(UnmanagedType.BStr)> DbUser As String,
											  <MarshalAs(UnmanagedType.BStr)> DbPassword As String,
											  <MarshalAs(UnmanagedType.BStr)> CardNumber As String,
											  <MarshalAs(UnmanagedType.BStr)> Expiration As String) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

		Dim Ret As ClsReturn
		Dim Client As ClsExecuter
		Dim Result As ClsResult

		Try
			Client = CreateCGClient(OdbcName, DbUser, DbPassword)
			Result = Client.APITokenize(CardNumber, Expiration)
			Ret = CreateAPITokenizationReturn(Result)
		Catch Ex As Exception
			Ret = CreateRuntimeErrorReturn(Ex)
		End Try

		Return Ret
	End Function

	Private Function CreateAPITokenizationReturn(CGResult As ClsResult) As ClsReturn
		Dim ResultCategory As EnmResultCategory
		Dim StatusCode As Integer
		Dim StatusText As String = Nothing
		Dim ResultProperties As ClsAPITokenization = Nothing

		APIResultToMyResult(CGResult, ResultCategory, StatusCode, StatusText, ResultProperties)
		If ResultCategory = EnmResultCategory.CGFailure AndAlso Not String.IsNullOrEmpty(ResultProperties.CardToken) Then
			StatusCode = SuccessCode_
		End If
		If StatusCode = SuccessCode_ Then
			StatusText = ResultProperties.CardToken
		End If

		Return New ClsReturn(StatusCode, StatusText, Nothing, Nothing)
	End Function

#End Region

#Region "APITransactions.Verification"

	Function CreditGuardAPIVerify_FullParams(<MarshalAs(UnmanagedType.BStr)> CgApiUrl As String,
											 <MarshalAs(UnmanagedType.BStr)> CgApiUser As String,
											 <MarshalAs(UnmanagedType.BStr)> CgApiPassword As String,
											 <MarshalAs(UnmanagedType.BStr)> TerminalNumber As String,
											 <MarshalAs(UnmanagedType.BStr)> CardIdentifier As String,
											 <MarshalAs(UnmanagedType.BStr)> Expiration As String,
											 <MarshalAs(UnmanagedType.U4)> Optional OwnerID As UInteger = 0) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

		Dim Ret As ClsReturn
		Dim Client As ClsExecuter
		Dim Result As ClsResult

		Try
			Client = CreateCGClient(1, CgApiUrl, CgApiUser, CgApiPassword)
			Result = Client.APIVerify(TerminalNumber, CardIdentifier, Expiration, OwnerID)
			Ret = CreateAPIVerifyReturn(Result)
		Catch Ex As Exception
			Ret = CreateRuntimeErrorReturn(Ex)
		End Try

		Return Ret
	End Function

	Function CreditGuardAPIVerify_UsingODBC(<MarshalAs(UnmanagedType.BStr)> OdbcName As String,
											<MarshalAs(UnmanagedType.BStr)> DbUser As String,
											<MarshalAs(UnmanagedType.BStr)> DbPassword As String,
											<MarshalAs(UnmanagedType.BStr)> TerminalNumber As String,
											<MarshalAs(UnmanagedType.BStr)> CardIdentifier As String,
											<MarshalAs(UnmanagedType.BStr)> Expiration As String,
											<MarshalAs(UnmanagedType.U4)> Optional OwnerID As UInteger = 0) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

		Dim Ret As ClsReturn
		Dim Client As ClsExecuter
		Dim Result As ClsResult

		Try
			Client = CreateCGClient(OdbcName, DbUser, DbPassword)
			Result = Client.APIVerify(TerminalNumber, CardIdentifier, Expiration, OwnerID)
			Ret = CreateAPIVerifyReturn(Result)
		Catch Ex As Exception
			Ret = CreateRuntimeErrorReturn(Ex)
		End Try

		Return Ret
	End Function

	Private Function CreateAPIVerifyReturn(CGResult As ClsResult) As ClsReturn
		Dim ResultCategory As EnmResultCategory
		Dim StatusCode As Integer
		Dim StatusText As String = Nothing
		Dim ResultProperties As ClsAPIVerification = Nothing
		Dim Data As CreditGuardClient.ClsAPIVerificationResult = Nothing

		APIResultToMyResult(CGResult, ResultCategory, StatusCode, StatusText, ResultProperties)
		If ResultCategory = EnmResultCategory.CGSuccess OrElse ResultCategory = EnmResultCategory.CGFailure Then
			Data = New CreditGuardClient.ClsAPIVerificationResult(ResultProperties.ExtendedResultCode, ResultProperties.ExtendedResultInnerText, ResultProperties.ExtendedResultOuterText, ResultProperties.CardToken, ResultProperties.CardType, ResultProperties.ResponseAuthorizationNumber)
		End If

		Return New ClsReturn(StatusCode, StatusText, Data, Nothing)
	End Function

#End Region

#Region "APITransactions.HKRegistration"

	Function CreditGuardAPIRegisterHK_FullParams(<MarshalAs(UnmanagedType.BStr)> CgApiUrl As String,
												 <MarshalAs(UnmanagedType.BStr)> CgApiUser As String,
												 <MarshalAs(UnmanagedType.BStr)> CgApiPassword As String,
												 <MarshalAs(UnmanagedType.BStr)> TerminalNumber As String,
												 <MarshalAs(UnmanagedType.BStr)> CardIdentifier As String,
												 <MarshalAs(UnmanagedType.BStr)> Expiration As String,
												 <MarshalAs(UnmanagedType.U4)> Optional OwnerID As UInteger = 0) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

		Dim Ret As ClsReturn
		Dim Client As ClsExecuter
		Dim Result As ClsResult

		Try
			Client = CreateCGClient(1, CgApiUrl, CgApiUser, CgApiPassword)
			Result = Client.APIRegisterHK(TerminalNumber, CardIdentifier, Expiration, OwnerID)
			Ret = CreateAPIRegisterHKReturn(Result)
		Catch Ex As Exception
			Ret = CreateRuntimeErrorReturn(Ex)
		End Try

		Return Ret
	End Function

	Function CreditGuardAPIRegisterHK_UsingODBC(<MarshalAs(UnmanagedType.BStr)> OdbcName As String,
												<MarshalAs(UnmanagedType.BStr)> DbUser As String,
												<MarshalAs(UnmanagedType.BStr)> DbPassword As String,
												<MarshalAs(UnmanagedType.BStr)> TerminalNumber As String,
												<MarshalAs(UnmanagedType.BStr)> CardIdentifier As String,
												<MarshalAs(UnmanagedType.BStr)> Expiration As String,
												<MarshalAs(UnmanagedType.U4)> Optional OwnerID As UInteger = 0) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

		Dim Ret As ClsReturn
		Dim Client As ClsExecuter
		Dim Result As ClsResult

		Try
			Client = CreateCGClient(OdbcName, DbUser, DbPassword)
			Result = Client.APIRegisterHK(TerminalNumber, CardIdentifier, Expiration, OwnerID)
			Ret = CreateAPIRegisterHKReturn(Result)
		Catch Ex As Exception
			Ret = CreateRuntimeErrorReturn(Ex)
		End Try

		Return Ret
	End Function

	Private Function CreateAPIRegisterHKReturn(CGResult As ClsResult) As ClsReturn
		Dim ResultCategory As EnmResultCategory
		Dim StatusCode As Integer
		Dim StatusText As String = Nothing
		Dim ResultProperties As ClsAPIHKRegistration = Nothing
		Dim Data As CreditGuardClient.ClsAPIHKRegistrationResult = Nothing

		APIResultToMyResult(CGResult, ResultCategory, StatusCode, StatusText, ResultProperties)
		If ResultCategory = EnmResultCategory.CGSuccess OrElse ResultCategory = EnmResultCategory.CGFailure Then
			Data = New CreditGuardClient.ClsAPIHKRegistrationResult(ResultProperties.ExtendedResultCode, ResultProperties.ExtendedResultInnerText, ResultProperties.ExtendedResultOuterText, ResultProperties.CardToken, ResultProperties.CardType, ResultProperties.ResponseAuthorizationNumber)
		End If

		Return New ClsReturn(StatusCode, StatusText, Data, Nothing)
	End Function

#End Region

#Region "APITransactions.Payment"

	Function CreditGuardAPIPay_FullParams(<MarshalAs(UnmanagedType.BStr)> CgApiUrl As String,
										  <MarshalAs(UnmanagedType.BStr)> CgApiUser As String,
										  <MarshalAs(UnmanagedType.BStr)> CgApiPassword As String,
									 	  <MarshalAs(UnmanagedType.BStr)> TerminalNumber As String,
									 	  <MarshalAs(UnmanagedType.BStr)> CardIdentifier As String,
										  <MarshalAs(UnmanagedType.BStr)> Expiration As String,
										  <MarshalAs(UnmanagedType.U4)> OwnerID As UInteger,
										  <MarshalAs(UnmanagedType.BStr)> AuthorizationNumber As String,
										  <MarshalAs(UnmanagedType.U1)> CreditType As CreditGuardClient.EnmCreditType,
										  <MarshalAs(UnmanagedType.R8)> TotalSum As Double,
										  <MarshalAs(UnmanagedType.U1)> NumberPayments As Byte,
										  <MarshalAs(UnmanagedType.BStr)> XField As String) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

		Dim Ret As ClsReturn
		Dim Client As ClsExecuter
		Dim Result As ClsResult

		Try
			Client = CreateCGClient(1, CgApiUrl, CgApiUser, CgApiPassword)
			Result = Client.APIPay(TerminalNumber, CardIdentifier, Expiration, OwnerID, AuthorizationNumber, CreditType, TotalSum, NumberPayments, XField)
			Ret = CreateAPIPayReturn(Result)
		Catch Ex As Exception
			Ret = CreateRuntimeErrorReturn(Ex)
		End Try

		Return Ret
	End Function

	Function CreditGuardAPIPay_UsingODBC(<MarshalAs(UnmanagedType.BStr)> OdbcName As String,
										 <MarshalAs(UnmanagedType.BStr)> DbUser As String,
										 <MarshalAs(UnmanagedType.BStr)> DbPassword As String,
									 	 <MarshalAs(UnmanagedType.BStr)> TerminalNumber As String,
									 	 <MarshalAs(UnmanagedType.BStr)> CardIdentifier As String,
										 <MarshalAs(UnmanagedType.BStr)> Expiration As String,
										 <MarshalAs(UnmanagedType.U4)> OwnerID As UInteger,
										 <MarshalAs(UnmanagedType.BStr)> AuthorizationNumber As String,
										 <MarshalAs(UnmanagedType.U1)> CreditType As CreditGuardClient.EnmCreditType,
										 <MarshalAs(UnmanagedType.R8)> TotalSum As Double,
										 <MarshalAs(UnmanagedType.U1)> NumberPayments As Byte,
										 <MarshalAs(UnmanagedType.BStr)> XField As String) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

		Dim Ret As ClsReturn
		Dim Client As ClsExecuter
		Dim Result As ClsResult

		Try
			Client = CreateCGClient(OdbcName, DbUser, DbPassword)
			Result = Client.APIPay(TerminalNumber, CardIdentifier, Expiration, OwnerID, AuthorizationNumber, CreditType, TotalSum, NumberPayments, XField)
			Ret = CreateAPIPayReturn(Result)
		Catch Ex As Exception
			Ret = CreateRuntimeErrorReturn(Ex)
		End Try

		Return Ret
	End Function

	Private Function CreateAPIPayReturn(CGResult As ClsResult) As ClsReturn
		Dim ResultCategory As EnmResultCategory
		Dim StatusCode As Integer
		Dim StatusText As String = Nothing
		Dim ResultProperties As ClsAPIPayment = Nothing
		Dim Data As CreditGuardClient.ClsAPIPaymentResult = Nothing

		APIResultToMyResult(CGResult, ResultCategory, StatusCode, StatusText, ResultProperties)
		If ResultCategory = EnmResultCategory.CGSuccess OrElse ResultCategory = EnmResultCategory.CGFailure Then
			Data = New CreditGuardClient.ClsAPIPaymentResult(ResultProperties.ShvaUID, ResultProperties.ExtendedResultCode, ResultProperties.ExtendedResultInnerText, ResultProperties.ExtendedResultOuterText,
															 ResultProperties.ResponseAuthorizationNumber, ResultProperties.CardToken, ResultProperties.CardType,
															 ResultProperties.CreditType, ResultProperties.CardAcquirer, ResultProperties.CardIssuer, ResultProperties.CardBrand, ResultProperties.ShovarNumber)
		End If

		Return New ClsReturn(StatusCode, StatusText, Data, Nothing)
	End Function

#End Region

#Region "APITransactions.InquireTransaction"

	Function CreditGuardAPIInquireTransaction_FullParams(<MarshalAs(UnmanagedType.U4)> CgMid As Integer,
														 <MarshalAs(UnmanagedType.BStr)> CgApiUrl As String,
														 <MarshalAs(UnmanagedType.BStr)> CgApiUser As String,
														 <MarshalAs(UnmanagedType.BStr)> CgApiPassword As String,
														 <MarshalAs(UnmanagedType.BStr)> TerminalNumber As String,
														 <MarshalAs(UnmanagedType.BStr)> WebPageToken As String) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

		Dim Ret As ClsReturn
		Dim Client As ClsExecuter
		Dim Result As ClsResult

		Try
			Client = CreateCGClient(CgMid, CgApiUrl, CgApiUser, CgApiPassword)
			Result = Client.APIInquireTransaction(TerminalNumber, WebPageToken)
			Ret = CreateAPIInquiryReturn(Result)
		Catch Ex As Exception
			Ret = CreateRuntimeErrorReturn(Ex)
		End Try

		Return Ret
	End Function

	Function CreditGuardAPIInquireTransaction_UsingODBC(<MarshalAs(UnmanagedType.BStr)> OdbcName As String,
														<MarshalAs(UnmanagedType.BStr)> DbUser As String,
														<MarshalAs(UnmanagedType.BStr)> DbPassword As String,
														<MarshalAs(UnmanagedType.BStr)> TerminalNumber As String,
														<MarshalAs(UnmanagedType.BStr)> WebPageToken As String) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

		Dim Ret As ClsReturn
		Dim Client As ClsExecuter
		Dim Result As ClsResult

		Try
			Client = CreateCGClient(OdbcName, DbUser, DbPassword)
			Result = Client.APIInquireTransaction(TerminalNumber, WebPageToken)
			Ret = CreateAPIInquiryReturn(Result)
		Catch Ex As Exception
			Ret = CreateRuntimeErrorReturn(Ex)
		End Try

		Return Ret
	End Function

	Private Function CreateAPIInquiryReturn(CGResult As ClsResult) As ClsReturn
		Dim ResultCategory As EnmResultCategory
		Dim StatusCode As Integer
		Dim StatusText As String = Nothing
		Dim ResultProperties As ClsAPITransactionInquiry = Nothing
		Dim Data As CreditGuardClient.ClsAPITransactionInquiryResult = Nothing

		APIResultToMyResult(CGResult, ResultCategory, StatusCode, StatusText, ResultProperties)
		If ResultCategory = EnmResultCategory.CGSuccess OrElse ResultCategory = EnmResultCategory.CGFailure Then
			Data = New CreditGuardClient.ClsAPITransactionInquiryResult(WebPageTransactionResultToMyResult(ResultProperties.WebPageTransactionResult), ResultProperties.XField, ResultProperties.ShvaUID, ResultProperties.TerminalNumber,
																		ResultProperties.ExtendedResultCode, ResultProperties.ExtendedResultInnerText, ResultProperties.ExtendedResultOuterText,
																		ResultProperties.CardToken, ResultProperties.Expiration, ResultProperties.OwnerID,
																		ResultProperties.CardType, ResultProperties.CreditType, ResultProperties.CardAcquirer, ResultProperties.CardIssuer, ResultProperties.CardBrand,
																		ResultProperties.ResponseAuthorizationNumber, ResultProperties.ShovarNumber)
		End If

		Return New ClsReturn(StatusCode, StatusText, Data, Nothing)
	End Function

#End Region

#Region "APITransactions.TransactionsRangeInquiry"

	Function CreditGuardAPIInquirePaymentTransactionsIntoSql_UsingODBC(<MarshalAs(UnmanagedType.BStr)> OdbcName As String,
																	   <MarshalAs(UnmanagedType.BStr)> DbUser As String,
																	   <MarshalAs(UnmanagedType.BStr)> DbPassword As String,
																	   <MarshalAs(UnmanagedType.LPArray)> TerminalNumbers As String(),
																	   <MarshalAs(UnmanagedType.Struct)> From As Date,
																	   <MarshalAs(UnmanagedType.Struct)> [To] As Date,
																	   <MarshalAs(UnmanagedType.BStr)> UserName As String,
																	   <MarshalAs(UnmanagedType.BStr)> AccessVersion As String) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

		Dim Ret As ClsReturn
		Dim Client As ClsExecuter
		Dim ErrorMessage As String = Nothing
		Dim SqlRowsUniqueID As Integer
		Dim Result As EnmInquireTransactionRangeResult
		Try
			Client = CreateCGClient(OdbcName, DbUser, DbPassword)
			Result = Client.APIInquirePaymentTransactionsIntoSql(TerminalNumbers, From, [To], UserName, AccessVersion, ErrorMessage, SqlRowsUniqueID)
			Select Case Result
				Case EnmInquireTransactionRangeResult.Success
					Ret = New ClsReturn(SqlRowsUniqueID, "", Nothing, Nothing)
				Case EnmInquireTransactionRangeResult.NoResults
					Ret = New ClsReturn(-1, "", Nothing, Nothing)
				Case Else
					Ret = New ClsReturn(0, ErrorMessage, Nothing, Nothing)
			End Select
		Catch ex As Exception
			Ret = CreateRuntimeErrorReturn(ex)
		End Try

		Return Ret
	End Function

#End Region

#End Region

End Class
