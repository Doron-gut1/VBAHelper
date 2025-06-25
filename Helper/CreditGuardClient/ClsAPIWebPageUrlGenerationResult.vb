Imports System.Runtime.InteropServices

Namespace CreditGuardClient

    <ComVisible(True), ClassInterface(ClassInterfaceType.AutoDual)>
    Public Class ClsAPIWebPageUrlGenerationResult

        ReadOnly Property ExtendedResultCode As <MarshalAs(UnmanagedType.BStr)> String

        ReadOnly Property ExtendedResultInnerText As <MarshalAs(UnmanagedType.BStr)> String

        ReadOnly Property ExtendedResultOuterText As <MarshalAs(UnmanagedType.BStr)> String

        ReadOnly Property WebPageToken As <MarshalAs(UnmanagedType.BStr)> String

        ReadOnly Property WebPageUrl As <MarshalAs(UnmanagedType.BStr)> String

        Sub New(ExtendedResultCode As String, ExtendedResultInnerText As String, ExtendedResultOuterText As String, WebPageToken As String, WebPageUrl As String)
            Me.ExtendedResultCode = ExtendedResultCode
            Me.ExtendedResultInnerText = ExtendedResultInnerText
            Me.ExtendedResultOuterText = ExtendedResultOuterText
            Me.WebPageToken = WebPageToken
            Me.WebPageUrl = WebPageUrl
        End Sub

    End Class
End Namespace