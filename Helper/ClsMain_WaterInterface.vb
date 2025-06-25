
Imports System.Runtime.InteropServices
Imports Helper.SmsManager.SmsEnum

Partial Public Class ClsMain

    Public Function GetAradGroupRead(ByVal monewtr As String, ByVal odbcName As String, ByVal Optional readdate As String = Nothing, ByVal Optional user As String = Nothing,
                                   ByVal Optional password As String = Nothing) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn


        Dim errDesc As String
        Dim AradKriotReader As WaterInterface.AradCls
        Dim errObj As ClsError = Nothing
        Dim result As Boolean
        Dim readDateTime As DateTime, withdate As Boolean
        Try

            errDesc = ""

            withdate = DateTime.TryParse(readdate, readDateTime)

            AradKriotReader = New WaterInterface.AradCls(odbcName, user, password)
            result = AradKriotReader.GetAradGroupRead(monewtr, IIf(withdate, readDateTime, Nothing))

            ' Sms = New SmsManager.SmsTypeFunctions(odbcName, user, password)

            'If smsType = SmsTypeEnum.ImutPhonePass Then
            '    result = Sms.SendImutPhonePassSms(Rnd, tel, smsStatus, smsStatusDesc, errDesc)
            'ElseIf smsType = SmsTypeEnum.UpdateInfoPass Then
            '    result = Sms.SendUpdateInfoPassSms(Rnd, tel, smsStatus, smsStatusDesc, errDesc)
            'End If

            If result Then

            Else
                errObj = New ClsError(0, $"Message: {errDesc}", errDesc)
            End If
        Catch ex As Exception
            Dim InnerEx As Exception = ExtractMustInnerException(ex)
            errObj = New ClsError(0, $"Message: {InnerEx.Message}{vbCrLf}{vbCrLf}Stack-Trace: {InnerEx.StackTrace}", InnerEx.Message)
        End Try

        Return New ClsReturn(0, "", Nothing, errObj)
    End Function

End Class
