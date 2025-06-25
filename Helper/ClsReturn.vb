Imports System.Runtime.InteropServices

<ComVisible(True), ClassInterface(ClassInterfaceType.AutoDual)>
Public Class ClsReturn

    ReadOnly Property StatusCode As <MarshalAs(UnmanagedType.I4)> Integer

    ReadOnly Property StatusText As <MarshalAs(UnmanagedType.BStr)> String

    ReadOnly Property DataObject As <MarshalAs(UnmanagedType.IDispatch)> Object

    ReadOnly Property [Error] As <MarshalAs(UnmanagedType.IDispatch)> ClsError

    Sub New(StatusCode As Integer, StatusText As String, DataObject As Object, [Error] As ClsError)
        Me.StatusCode = StatusCode
        Me.StatusText = StatusText
        Me.DataObject = DataObject
        Me.Error = [Error]
    End Sub

End Class
