Imports System.Runtime.InteropServices
Imports CreditGuardGlobals

Namespace CreditGuardClient

    <ComVisible(True), ClassInterface(ClassInterfaceType.AutoDual)>
    Public Class ClsAPITransactionInquiryResult

        ReadOnly Property WebPageTransactionResult As <MarshalAs(UnmanagedType.U1)> EnmWebPageTransactionResult

        ReadOnly Property XField As <MarshalAs(UnmanagedType.BStr)> String

        ReadOnly Property ShvaUID As <MarshalAs(UnmanagedType.BStr)> String

        ReadOnly Property TerminalNumber As <MarshalAs(UnmanagedType.BStr)> String

        ReadOnly Property ExtendedResultCode As <MarshalAs(UnmanagedType.BStr)> String

        ReadOnly Property ExtendedResultInnerText As <MarshalAs(UnmanagedType.BStr)> String

        ReadOnly Property ExtendedResultOuterText As <MarshalAs(UnmanagedType.BStr)> String

        ReadOnly Property CardToken As <MarshalAs(UnmanagedType.BStr)> String

        ReadOnly Property Expiration As <MarshalAs(UnmanagedType.BStr)> String

        ReadOnly Property OwnerID As <MarshalAs(UnmanagedType.U4)> UInteger

        ReadOnly Property CardType As <MarshalAs(UnmanagedType.IDispatch)> Object = Nothing

        ReadOnly Property CreditType As <MarshalAs(UnmanagedType.IDispatch)> Object = Nothing

        ReadOnly Property CardAcquirer As <MarshalAs(UnmanagedType.IDispatch)> Object = Nothing

        ReadOnly Property CardIssuer As <MarshalAs(UnmanagedType.IDispatch)> Object = Nothing

        ReadOnly Property CardBrand As <MarshalAs(UnmanagedType.IDispatch)> Object = Nothing

        ReadOnly Property ResponseAuthorizationNumber As <MarshalAs(UnmanagedType.BStr)> String

        ReadOnly Property ShovarNumber As <MarshalAs(UnmanagedType.BStr)> String

        Friend Sub New(WebPageTransactionResult As EnmWebPageTransactionResult,
                       XField As String, ShvaUID As String, TerminalNumber As String,
                       ExtendedResultCode As String, ExtendedResultInnerText As String, ExtendedResultOuterText As String,
                       CardToken As String, Expiration As String, OwnerID As UInteger,
                       CardType As EnmCardType?, CreditType As EnmCreditType?, CardAcquirer As EnmCardAcquirer?, CardIssuer As EnmCardIssuer?, CardBrand As EnmCardBrand?, ResponseAuthorizationNumber As String, ShovarNumber As String)
            Me.WebPageTransactionResult = WebPageTransactionResult
            Me.XField = XField
            Me.ShvaUID = ShvaUID
            Me.TerminalNumber = TerminalNumber
            Me.ExtendedResultCode = ExtendedResultCode
            Me.ExtendedResultInnerText = ExtendedResultInnerText
            Me.ExtendedResultOuterText = ExtendedResultOuterText
            Me.CardToken = CardToken
            Me.Expiration = Expiration
            Me.OwnerID = OwnerID
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
            Me.ResponseAuthorizationNumber = ResponseAuthorizationNumber
            Me.ShovarNumber = ShovarNumber
        End Sub

    End Class
End Namespace