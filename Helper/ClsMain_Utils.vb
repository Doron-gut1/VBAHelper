Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Access
Imports CreditCardsManager
Imports SendEmailManager
Imports ExcelImportManager
Imports CalcArnProcess

Partial Public Class ClsMain

    Function ReplaceCreditCards(ByVal userName As String, ByVal creditCardCompany As Integer, ByVal replacementFileName As String,
                                ByVal odbcName As String, ByVal Optional user As String = Nothing,
                               ByVal Optional password As String = Nothing) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

        Dim fileIndex As Integer = -1
        Dim errDesc As String = ""
        Dim mngr As ReplacementFileManager
        Dim StatusCode As Integer = 0, StatusText As String = ""
        Dim errObj As ClsError = Nothing
        Try
            mngr = New ReplacementFileManager(userName, creditCardCompany, odbcName, user, password)
            If mngr.ReplaceCreditCards(replacementFileName, fileIndex, errDesc) Then
                StatusCode = fileIndex
                StatusText = "קובץ מוחלפים עודכן בהצלחה"
            Else
                errObj = New ClsError(0, $"Message: {errDesc}", errDesc)
            End If
        Catch ex As Exception
            Dim InnerEx As Exception = ExtractMustInnerException(ex)
            errObj = New ClsError(0, $"Message: {InnerEx.Message}{vbCrLf}{vbCrLf}Stack-Trace: {InnerEx.StackTrace}", InnerEx.Message)
        End Try

        Return New ClsReturn(StatusCode, StatusText, Nothing, errObj)
    End Function

    Function CalcArnProcessManager(ByVal moazaCode As Integer, ByVal userName As String, ByVal odbcName As String, ByVal mnt As Integer, ByVal valdt As String, ByVal semel As Integer,
                                   ByVal Optional Sgrnum As Integer = 0, ByVal Optional Zmani As Integer = 0, ByVal Optional OneSugts As Integer = 0,
                                   ByVal Optional Hskod As String = "", ByVal Optional NoArn As Boolean = False, ByVal Optional Orgmkdmmadad As Decimal = 0,
                                   ByVal Optional Jobnum As Integer = 0, ByVal Optional processType As Integer = 0) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

        Dim errDesc As String = ""
        Dim mngr As CalcArnProcessManager
        Dim StatusCode As Integer = 0, StatusText As String = ""
        Dim errObj As ClsError = Nothing
        Try
            mngr = New CalcArnProcessManager(moazaCode, userName, odbcName, mnt, valdt, semel, Sgrnum, Zmani, OneSugts, Hskod, NoArn, Orgmkdmmadad, Jobnum, processType, , Nothing, 0)
            If mngr.CalcArnProcess(errDesc, Hskod) Then
                StatusCode = 1
                StatusText = "חישוב ארנונה בוצע בהצלחה"
            Else
                errObj = New ClsError(0, $"Message: {errDesc}", errDesc)
            End If
        Catch ex As Exception
            Dim InnerEx As Exception = ExtractMustInnerException(ex)
            errObj = New ClsError(0, $"Message: {InnerEx.Message}{vbCrLf}{vbCrLf}Stack-Trace: {InnerEx.StackTrace}", InnerEx.Message)
        End Try

        Return New ClsReturn(StatusCode, StatusText, Nothing, errObj)
    End Function

    Function CalcRetroProcessManager(ByVal moazaCode As Integer, ByVal userName As String, ByVal odbcName As String,
                                     ByVal Optional Jobnum As Integer = 0, ByVal Optional processType As Integer = 0,
                                     ByVal Optional Hskod As String = "") As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

        Dim retro As Retro
        Dim StatusCode As Integer = 0
        Dim StatusText As String = ""
        Dim errObj As ClsError = Nothing

        Try
            retro = New Retro(moazaCode, userName, odbcName, Jobnum, processType, Hskod)

            Dim result = retro.CalculateRetro()

            If result.Success Then
                StatusCode = 1
                StatusText = "חישוב רטרו בוצע בהצלחה"
            Else
                errObj = New ClsError(0, $"Message: {result.ErrorDescription}", result.ErrorDescription)
            End If
        Catch ex As Exception
            Dim InnerEx As Exception = ExtractMustInnerException(ex)
            errObj = New ClsError(0, $"Message: {InnerEx.Message}{vbCrLf}{vbCrLf}Stack-Trace: {InnerEx.StackTrace}", InnerEx.Message)
        End Try

        Return New ClsReturn(StatusCode, StatusText, Nothing, errObj)
    End Function
    Function ProcessExcel(ByVal moazaCode As Integer, ByVal jobNum As Integer, ByVal FileName As String, ByVal fileType As Integer,
            ByVal odbcName As String, ByVal userName As String, ByVal Optional password As String = Nothing) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

        Dim errDesc As String = ""
        Dim mngr As ExcelImportManager.ExcelImportManager
        Dim StatusCode As Integer = 0, StatusText As String = ""
        Dim errObj As ClsError = Nothing
        Try
            mngr = New ExcelImportManager.ExcelImportManager(moazaCode, jobNum, odbcName, userName, password)
            If mngr.ProcessExcel(FileName, fileType, errDesc) Then
                StatusText = "קובץ אקסל עודכן בהצלחה"
            Else
                errObj = New ClsError(0, $"Message: {errDesc}", errDesc)
            End If
        Catch ex As Exception
            Dim InnerEx As Exception = ExtractMustInnerException(ex)
            errObj = New ClsError(0, $"Message: {InnerEx.Message}{vbCrLf}{vbCrLf}Stack-Trace: {InnerEx.StackTrace}", InnerEx.Message)
        End Try

        Return New ClsReturn(StatusCode, StatusText, Nothing, errObj)
    End Function

    Function ExecuteHkRegistrationAndHkPayment(ByVal sugts As Integer, ByVal mntOrMana As Integer, ByVal userName As String, ByVal source As String,
                              ByVal sendUniqueID As String, ByVal odbcName As String,
                              ByVal Optional user As String = Nothing, ByVal Optional password As String = Nothing) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn

        Dim errDesc As String = ""
        Dim mngr As ShvaSendHkManager
        Dim StatusCode As Integer = 0
        Dim StatusText As String = ""
        Dim serverName As String
        Dim dbName As String
        Dim errObj As ClsError = Nothing
        Dim odbcConvert As OdbcConverter.OdbcConverter = New OdbcConverter.OdbcConverter()
        Dim connString As String
        Try

            connString = odbcConvert.GetSqlConnectionString(odbcName)
            serverName = odbcConvert.ServerName
            dbName = odbcConvert.DbName

            mngr = New ShvaSendHkManager(sugts, mntOrMana, userName, source, sendUniqueID, serverName, dbName, user, password)
            If mngr.RunProcess(errDesc) Then
                StatusCode = 1
                StatusText = "תהליך הסתיים בהצלחה"
            Else
                errObj = New ClsError(0, $"Message: {errDesc}", errDesc)
            End If
        Catch ex As Exception
            Dim InnerEx As Exception = ExtractMustInnerException(ex)
            errObj = New ClsError(0, $"Message: {InnerEx.Message}{vbCrLf}{vbCrLf}Stack-Trace: {InnerEx.StackTrace}", InnerEx.Message)
        End Try

        Return New ClsReturn(StatusCode, StatusText, Nothing, errObj)
    End Function

    Function FixMainAccessReferenceToExternal(<MarshalAs(UnmanagedType.BStr)> MainAccessPath As String,
                                              <MarshalAs(UnmanagedType.BStr)> ExternalAccessPath As String,
                                              <MarshalAs(UnmanagedType.BStr)> ReferenceName As String) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn
        Dim Status As Integer = 0
        Dim AccessApp As Application = Nothing
        Dim References As References = Nothing
        Dim ExternalRef As Reference = Nothing
        Dim [Error] As ClsError = Nothing

        Try
            AccessApp = New Application()

            Status = 1

            AccessApp.OpenCurrentDatabase(MainAccessPath, True)

            Status = 2

            References = AccessApp.References
            For Each Ref As Reference In References
                If Ref.Name.ToLower() = ReferenceName.ToLower() AndAlso Ref.FullPath.ToLower() <> ExternalAccessPath Then
                    ExternalRef = Ref
                    Exit For
                End If
            Next

            If ExternalRef IsNot Nothing Then
                AccessApp.References.Remove(ExternalRef)
                AccessApp.References.AddFromFile(ExternalAccessPath)
            End If

            Status = 3
        Catch Ex As Exception
            Dim InnerEx As Exception = ExtractMustInnerException(Ex)
            [Error] = New ClsError(0, $"Message: {InnerEx.Message}{vbCrLf}{vbCrLf}Stack-Trace: {InnerEx.StackTrace}", InnerEx.Message)
        Finally
            If Status > 0 Then
                If Status > 1 Then
                    Do Until Marshal.ReleaseComObject(References) = 0
                    Loop

                    If ExternalRef IsNot Nothing Then
                        Do Until Marshal.ReleaseComObject(ExternalRef) = 0
                        Loop
                    End If

                    AccessApp.CloseCurrentDatabase()
                End If

                AccessApp.Quit(If(Status = 3, AcQuitOption.acQuitSaveAll, AcQuitOption.acQuitSaveNone))

                Do Until Marshal.ReleaseComObject(AccessApp) = 0
                Loop
            End If
        End Try

        Return New ClsReturn(If(Status = 3, 1, 0), "", Nothing, [Error])
    End Function

    Function GenerateGuid() As <MarshalAs(UnmanagedType.BStr)> String
        Return Guid.NewGuid().ToString()
    End Function

End Class