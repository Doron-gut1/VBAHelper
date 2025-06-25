Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Access
Imports SendEmailManager

Partial Public Class ClsMain

    Function MtfEmailProcess(ByVal moazaCode As Integer, ByVal userName As String, ByVal odbcName As String, ByVal fromMnt As Integer, ByVal toMnt As Integer,
                             ByVal hskod As String, ByVal mspkod As Integer, ByVal kvuzaShovar As Integer, ByVal ishuv As Integer, ByVal operationCode As Integer) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

        Dim errDesc As String = ""
        Dim mngr As MtfEmailManager
        Dim StatusCode As Integer = 0, StatusText As String = ""
        Dim errObj As ClsError = Nothing
        Dim ctrlIndex As Integer
        Try
            mngr = New SendEmailManager.MtfEmailManager(moazaCode, userName, odbcName, False, operationCode)
            If mngr.MtfEmailProcess(fromMnt, toMnt, hskod, mspkod, kvuzaShovar, ctrlIndex, ishuv, errDesc) Then
                StatusCode = ctrlIndex
                StatusText = "מעטפיות נשלחו בהצלחה"
            Else
                errObj = New ClsError(0, $"Message: {errDesc}", errDesc)
            End If
        Catch ex As Exception
            Dim InnerEx As Exception = ExtractMustInnerException(ex)
            errObj = New ClsError(0, $"Message: {InnerEx.Message}{vbCrLf}{vbCrLf}Stack-Trace: {InnerEx.StackTrace}", InnerEx.Message)
        End Try

        Return New ClsReturn(StatusCode, StatusText, Nothing, errObj)
    End Function


    Function ShowRecipients(ByVal moazaCode As Integer, ByVal userName As String, ByVal odbcName As String, ByVal fromMnt As Integer, ByVal toMnt As Integer,
                             ByVal hskod As String, ByVal mspkod As Integer, ByVal kvuzaShovar As Integer, ByVal ishuv As Integer) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

        Dim errDesc As String = ""
        Dim mngr As MtfEmailManager
        Dim StatusCode As Integer = 0, StatusText As String = ""
        Dim errObj As ClsError = Nothing
        Dim ctrlIndex As Integer
        Try
            mngr = New SendEmailManager.MtfEmailManager(moazaCode, userName, odbcName)
            If mngr.ShowRecipients(fromMnt, toMnt, hskod, mspkod, kvuzaShovar, ctrlIndex, ishuv, errDesc) Then
                StatusCode = ctrlIndex
                StatusText = "הוכנו נמענים בהצלחה"
            Else
                errObj = New ClsError(0, $"Message: {errDesc}", errDesc)
            End If
        Catch ex As Exception
            Dim InnerEx As Exception = ExtractMustInnerException(ex)
            errObj = New ClsError(0, $"Message: {InnerEx.Message}{vbCrLf}{vbCrLf}Stack-Trace: {InnerEx.StackTrace}", InnerEx.Message)
        End Try

        Return New ClsReturn(StatusCode, StatusText, Nothing, errObj)
    End Function

    Function ShowMtfProcess(ByVal moazaCode As Integer, ByVal userName As String, ByVal odbcName As String, ByVal fromMnt As Integer, ByVal toMnt As Integer,
                             ByVal hskod As String, ByVal mspkod As Integer, ByVal kvuzaShovar As Integer, ByVal noEmail As Boolean,
                            ByVal midagm As Boolean, ByVal midgamAuto As Boolean, ByVal ishuv As Integer, ByVal operationCode As Integer) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

        Dim errDesc As String = ""
        Dim mngr As MtfEmailManager
        Dim StatusCode As Integer = 0, StatusText As String = ""
        Dim errObj As ClsError = Nothing
        Dim ctrlIndex As Integer = 0

        Try
            mngr = New SendEmailManager.MtfEmailManager(moazaCode, userName, odbcName, False, operationCode)
            If mngr.ShowMtfProcess(fromMnt, toMnt, hskod, mspkod, kvuzaShovar, noEmail, midagm, midgamAuto, ctrlIndex, ishuv, errDesc) Then
                StatusText = IIf(errDesc <> "", "הוכנו קבצי מעטפיות להצגה בהצלחה", "בעיה בהצגת מעטפיות")
                StatusCode = ctrlIndex
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