Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Security.Authentication
Imports Helper.EmailManager

Partial Public Class ClsMain
    'Function SendSingleEmail(smtpServer As String, smtpPort As Integer, smtpEncrypt As Boolean, smtpAuthenticate As Boolean, smtpUserName As String,
    '				smtpPassword As String, msgUniqueId As Integer, msgEmailFrom As String, msgEmailTo As String(), Optional msgSubject As String = "",
    '				Optional msgBody As String = "", Optional isHtml As Boolean = True, Optional attachments As String() = Nothing) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

    Function SendSingleEmail(ByVal msgEmailTo As String, ByVal msgSubject As String,
                    ByVal msgBody As String, isHtml As Boolean, attachments As String(), ByVal operationCode As Integer,
                    ByVal userName As String, ByVal odbcName As String, ByVal Optional dbUser As String = Nothing, ByVal Optional dbPassword As String = Nothing) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

        Dim settings As EmailSender.Settings.ClsSettings
        Dim sender As EmailSender.ClsSender
        Dim msg As New EmailSender.Settings.ClsMessage
        Dim result As EmailSender.Result.ClsProcessResult
        Dim errObj As ClsError = Nothing
        Dim StatusCode As Integer = 0, StatusText As String = ""
        Dim ds As New DsEmail
        Dim dal As DalEmail
        Try
            dal = New DalEmail(odbcName, dbUser, dbPassword)

            If Not dal.GetEmailParams(ds, operationCode, userName) Then
                Throw New Exception($"EmailParams not found for operationCode={operationCode}, userName={userName}")
            End If

            If attachments.Length = 1 AndAlso attachments(0) = "" Then
                attachments = Nothing
            End If

            With msg
                .UniqueId = 1
                .EmailFrom = ds.SendEmailsParams(0).EmailFrom
                .EmailTo = New String() {msgEmailTo}
                .Subject = msgSubject
                .Body = msgBody
                .IsBodyHtml = isHtml
                .Attachments = attachments
            End With

            settings = New EmailSender.Settings.ClsSettings() With
                                             {
                                             .SmtpSettings = New EmailSender.Settings.ClsSmtp() With
                                                {.Server = ds.SendEmailsParams(0).SmtpServer,
                                                .Port = ds.SendEmailsParams(0).ServerPort,
                                                .Encrypt = ds.SendEmailsParams(0).Encrypt,
                                                .AllowedEncryptionProtocols = New SslProtocols() {SslProtocols.Tls, SslProtocols.Tls11, SslProtocols.Tls12, 12288},
                                                .Authenticate = ds.SendEmailsParams(0).Authenticate,
                                                .UserName = ds.SendEmailsParams(0).UserName,
                                                .Password = ds.SendEmailsParams(0).Password},
                                             .Messages = New EmailSender.Settings.ClsMessage() {msg}
                                             }

            sender = New EmailSender.ClsSender(settings)

            result = sender.Send()

            'If Not result.MessageResults(0).Success Then
            If result.GeneralResult <> EmailSender.Result.EnmGeneralResult.CompletedSuccessfully Then
                StatusCode = result.GeneralResult
                StatusText = result.FatalException.Message
            End If

        Catch ex As Exception
            Dim InnerEx As Exception = ExtractMustInnerException(ex)
            errObj = New ClsError(0, $"Message: {InnerEx.Message}{vbCrLf}{vbCrLf}Stack-Trace: {InnerEx.StackTrace}", InnerEx.Message)
            StatusCode = -1
        End Try

        Return New ClsReturn(StatusCode, StatusText, Nothing, errObj)

    End Function

    Function CreateKblOperation(ByVal odbcName As String, ByVal Operation As Integer, ByVal ProcUserID As Integer, Optional ByVal mnt As Integer = 0, Optional ByVal thum As Integer = 0, Optional ByVal kabala As Integer = 0)
        Dim errObj As ClsError = Nothing
        Dim StatusCode As Integer = 0, StatusText As String = ""
        Try
            Return SendDocument(Operation, New DalEmail(odbcName).CreateKblOperation(Operation, ProcUserID, mnt, thum, kabala), odbcName)
        Catch ex As Exception
            Dim InnerEx As Exception = ExtractMustInnerException(ex)
            errObj = New ClsError(0, $"Message: {InnerEx.Message}{vbCrLf}{vbCrLf}Stack-Trace: {InnerEx.StackTrace}", InnerEx.Message)
            StatusCode = -1
        End Try
        Return New ClsReturn(StatusCode, StatusText, Nothing, errObj)
    End Function

    Function SendDocument(ByVal Operation As Integer, ByVal ControlIndex As Integer, ByVal odbcName As String)
        Dim errObj As ClsError = Nothing
        Dim StatusCode As Integer = 0, StatusText As String = ""
        Try
            Dim exePath As String = $"{Directory.GetCurrentDirectory()}\\DocumentsSender"
            Dim prc As New Process
            With prc.StartInfo
                .UseShellExecute = False
                .FileName = $"{exePath}\\DocumentsSender.exe"
                .Arguments = $"{exePath}@2@{Operation}@{ControlIndex}@{odbcName}"
            End With
            prc.Start()
        Catch ex As Exception
            Dim InnerEx As Exception = ExtractMustInnerException(ex)
            errObj = New ClsError(0, $"Message: {InnerEx.Message}{vbCrLf}{vbCrLf}Stack-Trace: {InnerEx.StackTrace}", InnerEx.Message)
            StatusCode = -1
        End Try
        Return New ClsReturn(StatusCode, StatusText, Nothing, errObj)
    End Function

    'Function GetEmailMessage(msgId As Integer, msgEmailFrom As String, msgEmailTo As String(), Optional msgSubject As String = "", Optional msgBody As String = "",
    '				 Optional isHtml As Boolean = True, Optional attachments As String() = Nothing) As EmailSender.Settings.ClsMessage
    '	Dim msg As New EmailSender.Settings.ClsMessage
    '	With msg
    '		.UniqueId = msgId
    '		.EmailFrom = msgEmailFrom
    '		.EmailTo = msgEmailTo
    '		.Subject = msgSubject
    '		.Body = msgBody
    '		.IsBodyHtml = isHtml
    '		.Attachments = attachments
    '	End With

    '	Return msg
    'End Function

End Class
