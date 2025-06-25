Imports System.Runtime.InteropServices
Imports TabuRequestClient

Partial Public Class ClsMain

    Public Function TabuCheckHs(ByVal BranchId As Long, ByVal ApiUrl As String, ByVal Usr As String, ByVal Password As String, ByVal Sernum As String, ByVal odbcName As String, ByVal Optional dbUser As String = Nothing, ByVal Optional dbPass As String = Nothing) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn
        Dim clsResult As ClsResult = Nothing
        Dim errObj As ClsError = Nothing
        Dim t As TabuMain

        t = New TabuMain(odbcName, dbUser, dbPass)
        clsResult = t.CheckPropertyExistence(BranchId, ApiUrl, Usr, Password, Sernum)
        Return ConvertResultToClsReturn(clsResult)
    End Function

    Public Function LackOfDebts(ByVal BranchId As Long, ByVal ApiUrl As String, ByVal Usr As String, ByVal Password As String, ByVal Sernum As String, ByVal odbcName As String, ByVal Optional dbUser As String = Nothing, ByVal Optional dbPass As String = Nothing) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn
        Dim clsResult As ClsResult = Nothing
        Dim errObj As ClsError = Nothing
        Dim t As TabuMain

        t = New TabuMain(odbcName, dbUser, dbPass)
        clsResult = t.CheckLackOfDebtsInProperty(BranchId, ApiUrl, Usr, Password, Sernum)
        Return ConvertResultToClsReturn(clsResult)
    End Function

    Private Function ConvertResultToClsReturn(ByVal result As ClsResult) As <MarshalAs(UnmanagedType.IDispatch)> ClsReturn
        Return New ClsReturn(result.StatusCode, result.StatusText, Nothing, New Helper.ClsError(result.Error.ErrNum, result.Error.ErrInnerDesc, result.Error.ErrOuterDesc))
    End Function
End Class