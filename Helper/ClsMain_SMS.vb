Imports System.Runtime.InteropServices
Imports Helper.SmsManager.SmsEnum

Partial Public Class ClsMain

    Function SendHinHarshama(ByVal yldName As String, ByVal tel As String, ByVal odbcName As String, ByVal Optional user As String = Nothing,
                        ByVal Optional password As String = Nothing) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

        Dim smsStatus As Integer
        Dim smsStatusDesc, errDesc As String
        Dim Sms As SmsManager.SmsTypeFunctions
        Dim StatusCode As Integer = 0, StatusText As String = ""
        Dim errObj As ClsError = Nothing

        Try
            smsStatusDesc = ""
            errDesc = ""
            Sms = New SmsManager.SmsTypeFunctions(odbcName, user, password)
            If Sms.SendHinHarshama(yldName, tel, smsStatus, smsStatusDesc, errDesc) Then
                StatusCode = smsStatus
                StatusText = smsStatusDesc
            Else
                errObj = New ClsError(0, $"Message: {errDesc}", errDesc)
            End If
        Catch ex As Exception
            Dim InnerEx As Exception = ExtractMustInnerException(ex)
            errObj = New ClsError(0, $"Message: {InnerEx.Message}{vbCrLf}{vbCrLf}Stack-Trace: {InnerEx.StackTrace}", InnerEx.Message)
        End Try

        Return New ClsReturn(StatusCode, StatusText, Nothing, errObj)
    End Function

    Function SendInvoiceSms(ByVal invoiceRashut As Integer, ByVal invoiceNum As Integer, ByVal tel As String, ByVal pdfLocation As String,
                        ByVal shouldOverwrite As Boolean, ByVal odbcName As String, ByVal Optional user As String = Nothing,
                        ByVal Optional password As String = Nothing) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

        Dim smsStatus As Integer
        Dim smsStatusDesc, errDesc As String
        Dim Sms As SmsManager.SmsTypeFunctions
        Dim StatusCode As Integer = 0, StatusText As String = ""
        Dim errObj As ClsError = Nothing

        Try
            smsStatusDesc = ""
            errDesc = ""
            Sms = New SmsManager.SmsTypeFunctions(odbcName, user, password)
            If Sms.SendInvoiceLinkSms(invoiceRashut, invoiceNum, tel, pdfLocation, shouldOverwrite, smsStatus, smsStatusDesc, errDesc) Then
                StatusCode = smsStatus
                StatusText = smsStatusDesc
            Else
                errObj = New ClsError(0, $"Message: {errDesc}", errDesc)
            End If
        Catch ex As Exception
            Dim InnerEx As Exception = ExtractMustInnerException(ex)
            errObj = New ClsError(0, $"Message: {InnerEx.Message}{vbCrLf}{vbCrLf}Stack-Trace: {InnerEx.StackTrace}", InnerEx.Message)
        End Try

        Return New ClsReturn(StatusCode, StatusText, Nothing, errObj)
    End Function
    Function SendShovarHovLinkSms(ByVal manaHovNum As Integer, ByVal odbcName As String,
                              ByVal Optional isRecovery As Boolean = False,
                              ByVal Optional user As String = Nothing,
                              ByVal Optional password As String = Nothing) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn
        Dim ShovarType As SmsTypeEnum
        ShovarType = SmsTypeEnum.ShovarHovLink
        Dim totalInBatch As Integer = 0
        Dim newSuccessCount As Integer = 0
        Dim totalSuccessCount As Integer = 0
        Dim failedCount As Integer = 0
        Dim errDesc As String = ""
        Dim Sms As SmsManager.SmsTypeFunctions
        Dim StatusCode As Integer = -1, StatusText As String = ""
        Dim errObj As ClsError = Nothing
        'SendShovarHovLink
        Try
            Sms = New SmsManager.SmsTypeFunctions(odbcName, user, password)
            If Sms.ProcessShovarBatch(manaHovNum, ShovarType, isRecovery, totalInBatch, newSuccessCount, totalSuccessCount, failedCount, errDesc) Then
                StatusCode = newSuccessCount

                If isRecovery Then
                    ' במצב התאוששות - הודעה מפורטת
                    StatusText = $"סה""כ נשלחו {totalSuccessCount} שוברים מתוך {totalInBatch} במנת חוב מספר {manaHovNum} ({newSuccessCount} חדשים בריצה זו)"
                    If failedCount > 0 Then
                        StatusText &= $", {failedCount} עדיין נכשלו"
                    End If
                Else
                    ' מצב רגיל - הודעה רגילה
                    StatusText = $"נשלחו {newSuccessCount} מתוך {totalInBatch} שוברים בהצלחה במנת חוב מספר {manaHovNum}"
                    If failedCount > 0 Then
                        StatusText &= $", {failedCount} נכשלו"
                    End If
                End If
            Else
                errObj = New ClsError(0, $"Message: {errDesc}", errDesc)
            End If
        Catch ex As Exception
            Dim InnerEx As Exception = ExtractMustInnerException(ex)
            errObj = New ClsError(0, $"Message: {InnerEx.Message}{vbCrLf}{vbCrLf}Stack-Trace: {InnerEx.StackTrace}", InnerEx.Message)
        End Try

        Return New ClsReturn(StatusCode, StatusText, Nothing, errObj)
    End Function


    Function SendShovarLinkSms(ByVal shovarRashut As Integer, ByVal shovarNum As Integer, ByVal tz As Long, ByVal tel As String,
                              ByVal odbcName As String, ByVal Optional user As String = Nothing,
                                   ByVal Optional password As String = Nothing) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

        Dim ShovarType As SmsTypeEnum
        ShovarType = SmsTypeEnum.ShovarLink
        SendShovarLinkSms = SendAllShovarLinkSms(ShovarType, shovarRashut, shovarNum, tz, tel,
                               odbcName, user,
                             password)


    End Function


    Function SendAllShovarLinkSms(ShovarType As SmsTypeEnum, ByVal shovarRashut As Integer, ByVal shovarNum As Integer, ByVal tz As Long, ByVal tel As String,
                              ByVal odbcName As String, ByVal Optional user As String = Nothing,
                               ByVal Optional password As String = Nothing) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn
        Dim smsStatus As Integer
        Dim smsStatusDesc, errDesc As String
        Dim Sms As SmsManager.SmsTypeFunctions
        Dim StatusCode As Integer = 0, StatusText As String = ""
        Dim errObj As ClsError = Nothing

        Try
            smsStatusDesc = ""
            errDesc = ""
            Sms = New SmsManager.SmsTypeFunctions(odbcName, user, password)
            If Sms.SendAllShovarLinkSms(ShovarType, shovarRashut, shovarNum, tz, tel, smsStatus, smsStatusDesc, errDesc) Then
                StatusCode = smsStatus
                StatusText = smsStatusDesc
            Else
                errObj = New ClsError(0, $"Message: {errDesc}", errDesc)
            End If
        Catch ex As Exception
            Dim InnerEx As Exception = ExtractMustInnerException(ex)
            errObj = New ClsError(0, $"Message: {InnerEx.Message}{vbCrLf}{vbCrLf}Stack-Trace: {InnerEx.StackTrace}", InnerEx.Message)
        End Try

        Return New ClsReturn(StatusCode, StatusText, Nothing, errObj)
    End Function

    Private Function SendPassSms(smsType As SmsTypeEnum, rnd As String, ByVal tel As String, ByVal odbcName As String, ByVal Optional user As String = Nothing,
                               ByVal Optional password As String = Nothing) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn
        Dim smsStatus As Integer
        Dim smsStatusDesc, errDesc As String
        Dim Sms As SmsManager.SmsTypeFunctions
        Dim StatusCode As Integer = 0, StatusText As String = ""
        Dim errObj As ClsError = Nothing
        Dim result As Boolean

        Try
            smsStatusDesc = ""
            errDesc = ""
            Sms = New SmsManager.SmsTypeFunctions(odbcName, user, password)

            If smsType = SmsTypeEnum.ImutPhonePass Then
                result = Sms.SendImutPhonePassSms(rnd, tel, smsStatus, smsStatusDesc, errDesc)
            ElseIf smsType = SmsTypeEnum.UpdateInfoPass Then
                result = Sms.SendUpdateInfoPassSms(rnd, tel, smsStatus, smsStatusDesc, errDesc)
            End If

            If result Then
                StatusCode = smsStatus
                StatusText = smsStatusDesc
            Else
                errObj = New ClsError(0, $"Message: {errDesc}", errDesc)
            End If
        Catch ex As Exception
            Dim InnerEx As Exception = ExtractMustInnerException(ex)
            errObj = New ClsError(0, $"Message: {InnerEx.Message}{vbCrLf}{vbCrLf}Stack-Trace: {InnerEx.StackTrace}", InnerEx.Message)
        End Try

        Return New ClsReturn(StatusCode, StatusText, Nothing, errObj)
    End Function

    Function SendImutPhonePassSms(rnd As String, ByVal tel As String, ByVal odbcName As String, ByVal Optional user As String = Nothing,
                               ByVal Optional password As String = Nothing) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn
        Return SendPassSms(SmsTypeEnum.ImutPhonePass, rnd, tel, odbcName, user, password)
    End Function

    Function SendUpdateInfoPassSms(rnd As String, ByVal tel As String, ByVal odbcName As String, ByVal Optional user As String = Nothing,
                           ByVal Optional password As String = Nothing) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn
        Return SendPassSms(SmsTypeEnum.UpdateInfoPass, rnd, tel, odbcName, user, password)
    End Function

    Function SendGeneralSms(ByVal Content As String, ByVal tel As String, ByVal odbcName As String, ByVal Optional user As String = Nothing,
                        ByVal Optional password As String = Nothing) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

        Dim smsStatus As Integer
        Dim smsStatusDesc, errDesc As String
        Dim Sms As SmsManager.SmsTypeFunctions
        Dim StatusCode As Integer = 0, StatusText As String = ""
        Dim errObj As ClsError = Nothing

        Try
            smsStatusDesc = ""
            errDesc = ""
            Sms = New SmsManager.SmsTypeFunctions(odbcName, user, password)
            If Sms.SendGenearlMessage(Content, tel, smsStatus, smsStatusDesc, errDesc) Then
                StatusCode = smsStatus
                StatusText = smsStatusDesc
            Else
                errObj = New ClsError(0, $"Message: {errDesc}", errDesc)
            End If
        Catch ex As Exception
            Dim InnerEx As Exception = ExtractMustInnerException(ex)
            errObj = New ClsError(0, $"Message: {InnerEx.Message}{vbCrLf}{vbCrLf}Stack-Trace: {InnerEx.StackTrace}", InnerEx.Message)
        End Try

        Return New ClsReturn(StatusCode, StatusText, Nothing, errObj)
    End Function

    Function SendShvaHkUnacceptableSms(ByVal thum As Integer, ByVal SendNum As Integer, ByVal SendUniqueId As String, UserName As String, ByVal odbcName As String, ByVal Optional user As String = Nothing,
                        ByVal Optional password As String = Nothing) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

        Dim smsStatus As Integer
        Dim smsStatusDesc, errDesc As String
        Dim Sms As SmsManager.SmsTypeFunctions
        Dim StatusCode As Integer = 0, StatusText As String = ""
        Dim errObj As ClsError = Nothing

        Try
            smsStatusDesc = ""
            errDesc = ""
            Sms = New SmsManager.SmsTypeFunctions(odbcName, user, password)
            If Sms.SendShvaHkUnacceptableSms(thum, SendNum, SendUniqueId, UserName, smsStatus, smsStatusDesc, errDesc).IsCompleted Then
                StatusCode = 1
                StatusText = smsStatusDesc
            Else
                errObj = New ClsError(0, $"Message: {errDesc}", errDesc)
            End If
        Catch ex As Exception
            Dim InnerEx As Exception = ExtractMustInnerException(ex)
            errObj = New ClsError(0, $"Message: {InnerEx.Message}{vbCrLf}{vbCrLf}Stack-Trace: {InnerEx.StackTrace}", InnerEx.Message)
        End Try

        Return New ClsReturn(StatusCode, StatusText, Nothing, errObj)
    End Function

End Class