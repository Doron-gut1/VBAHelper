Imports System.Runtime.InteropServices

<ComVisible(True), ClassInterface(ClassInterfaceType.AutoDual)>
Public Class ClsError

    ReadOnly Property Number As <MarshalAs(UnmanagedType.I4)> Integer

    ReadOnly Property InnerDescription As <MarshalAs(UnmanagedType.BStr)> String

    ReadOnly Property OuterDescription As <MarshalAs(UnmanagedType.BStr)> String

    Sub New(Number As Integer, InnerDescription As String, OuterDescription As String)
        Me.Number = Number
        Me.InnerDescription = InnerDescription
        Me.OuterDescription = OuterDescription
    End Sub

End Class
