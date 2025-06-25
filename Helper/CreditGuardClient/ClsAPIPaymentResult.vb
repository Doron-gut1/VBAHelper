Imports System.Runtime.InteropServices
Imports CreditGuardGlobals

Namespace CreditGuardClient

    <ComVisible(True), ClassInterface(ClassInterfaceType.AutoDual)>
    Public Class ClsAPIPaymentResult

        ReadOnly Property ShvaUID As <MarshalAs(UnmanagedType.BStr)> String

        ReadOnly Property ExtendedResultCode As <MarshalAs(UnmanagedType.BStr)> String

        ReadOnly Property ExtendedResultInnerText As <MarshalAs(UnmanagedType.BStr)> String

        ReadOnly Property ExtendedResultOuterText As <MarshalAs(UnmanagedType.BStr)> String

        ReadOnly Property ResponseAuthorizationNumber As <MarshalAs(UnmanagedType.BStr)> String

        ReadOnly Property CardToken As <MarshalAs(UnmanagedType.BStr)> String

        ReadOnly Property CardType As <MarshalAs(UnmanagedType.IDispatch)> Object = Nothing

        ReadOnly Property CreditType As <MarshalAs(UnmanagedType.IDispatch)> Object = Nothing

        ReadOnly Property CardAcquirer As <MarshalAs(UnmanagedType.IDispatch)> Object = Nothing

        ReadOnly Property CardIssuer As <MarshalAs(UnmanagedType.IDispatch)> Object = Nothing

        ReadOnly Property CardBrand As <MarshalAs(UnmanagedType.IDispatch)> Object = Nothing

        ReadOnly Property ShovarNumber As <MarshalAs(UnmanagedType.BStr)> String

        Friend Sub New(ShvaUID As String,
                       ExtendedResultCode As String, ExtendedResultInnerText As String, ExtendedResultOuterText As String,
                       ResponseAuthorizationNumber As String, CardToken As String,
                       CardType As EnmCardType?, CreditType As EnmCreditType?, CardAcquirer As EnmCardAcquirer?, CardIssuer As EnmCardIssuer?, CardBrand As EnmCardBrand?,
                       ShovarNumber As String)

            Me.ShvaUID = ShvaUID
            Me.ExtendedResultCode = ExtendedResultCode
            Me.ExtendedResultInnerText = ExtendedResultInnerText
            Me.ExtendedResultOuterText = ExtendedResultOuterText
            Me.ResponseAuthorizationNumber = ResponseAuthorizationNumber
            Me.CardToken = CardToken
            If CardType.HasValue Then
                Me.CardType = CInt(CardType.Value)
            End If
            If CreditType.HasValue Then
                Me.CreditType = CByte(CreditType.Value)
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