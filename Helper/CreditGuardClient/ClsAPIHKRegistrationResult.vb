Imports System.Runtime.InteropServices
Imports CreditGuardGlobals

Namespace CreditGuardClient

    <ComVisible(True), ClassInterface(ClassInterfaceType.AutoDual)>
    Public Class ClsAPIHKRegistrationResult

        ReadOnly Property ExtendedResultCode As <MarshalAs(UnmanagedType.BStr)> String

        ReadOnly Property ExtendedResultInnerText As <MarshalAs(UnmanagedType.BStr)> String

        ReadOnly Property ExtendedResultOuterText As <MarshalAs(UnmanagedType.BStr)> String

        ReadOnly Property CardToken As <MarshalAs(UnmanagedType.BStr)> String

        ReadOnly Property CardType As <MarshalAs(UnmanagedType.IDispatch)> Object = Nothing

        ReadOnly Property AuthorizationNumber As <MarshalAs(UnmanagedType.BStr)> String

        Sub New(ExtendedResultCode As String, ExtendedResultInnerText As String, ExtendedResultOuterText As String, CardToken As String, CardType As EnmCardType?, AuthorizationNumber As String)
            Me.ExtendedResultCode = ExtendedResultCode
            Me.ExtendedResultInnerText = ExtendedResultInnerText
            Me.ExtendedResultOuterText = ExtendedResultOuterText
            Me.CardToken = CardToken
            If CardType.HasValue Then
                Me.CardType = CInt(CardType.Value)
            End If
            Me.AuthorizationNumber = AuthorizationNumber
        End Sub

    End Class
End Namespace