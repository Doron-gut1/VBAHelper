Imports System.Runtime.InteropServices
Imports CreditGuardGlobals

Namespace CreditGuardClient

    <ComVisible(True), ClassInterface(ClassInterfaceType.AutoDual)>
    Public Class ClsWebPagePaymentResult

        ReadOnly Property ExtendedResultCode As <MarshalAs(UnmanagedType.BStr)> String

        ReadOnly Property ExtendedResultInnerText As <MarshalAs(UnmanagedType.BStr)> String

        ReadOnly Property ExtendedResultOuterText As <MarshalAs(UnmanagedType.BStr)> String

        ReadOnly Property ShvaUID As <MarshalAs(UnmanagedType.BStr)> String

        ReadOnly Property AuthorizationNumber As <MarshalAs(UnmanagedType.BStr)> String

        ReadOnly Property CardToken As <MarshalAs(UnmanagedType.BStr)> String

        ReadOnly Property Expiration As <MarshalAs(UnmanagedType.BStr)> String

        ReadOnly Property OwnerID As <MarshalAs(UnmanagedType.U4)> UInteger

        ReadOnly Property CardType As <MarshalAs(UnmanagedType.IDispatch)> Object = Nothing

        ReadOnly Property CardAcquirer As <MarshalAs(UnmanagedType.IDispatch)> Object = Nothing

        ReadOnly Property CardIssuer As <MarshalAs(UnmanagedType.IDispatch)> Object = Nothing

        ReadOnly Property CardBrand As <MarshalAs(UnmanagedType.IDispatch)> Object = Nothing

        ReadOnly Property ShovarNumber As <MarshalAs(UnmanagedType.BStr)> String

        Sub New(ShvaUID As String,
                ExtendedResultCode As String, ExtendedResultInnerText As String, ExtendedResultOuterText As String,
                AuthorizationNumber As String, CardToken As String, Expiration As String, OwnerID As UInteger, CardType As EnmCardType?,
                CardAcquirer As EnmCardAcquirer?, CardIssuer As EnmCardIssuer?, CardBrand As EnmCardBrand?, ShovarNumber As String)
            Me.ShvaUID = ShvaUID
            Me.ExtendedResultCode = ExtendedResultCode
            Me.ExtendedResultInnerText = ExtendedResultInnerText
            Me.ExtendedResultOuterText = ExtendedResultOuterText
            Me.AuthorizationNumber = AuthorizationNumber
            Me.CardToken = CardToken
            Me.Expiration = Expiration
            Me.OwnerID = OwnerID
            If CardType.HasValue Then
                Me.CardType = CInt(CardType.Value)
            End If
            If CardAcquirer.HasValue Then
                Me.CardAcquirer = CByte(CardAcquirer.Value)
            End If
            If CardIssuer.HasValue Then
                Me.CardIssuer = CByte(CardIssuer.Value)
            End If
            If CardBrand.HasValue Then
                Me.CardBrand = CByte(CardBrand.Value)
            End If
            Me.ShovarNumber = ShovarNumber
        End Sub

    End Class
End Namespace