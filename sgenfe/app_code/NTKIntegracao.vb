'classe para importacao de pedidos do aplicativo da NTK (tokecompre)

Imports System
Imports System.IO
Imports System.Web.Script.Serialization
Imports System.Collections
Imports System.Collections.Generic

Public Class ClassNTKAddress

    Public street As String
    Public number As String
    Public neighborhood As String
    Public complement As String
    Public zipcode As String
    Public city As String
    Public uf As String
    Public reference As String

End Class

Public Class ClassNTKItem

    Public codigo As String 'do produto na loja
    Public id As String 'do produto na NTK
    Public name As String
    Public qtd As Double
    Public price_cents As Long
    Public subtotal_cents As Long
    Public complements As List(Of ClassNTKItem)
    Public obs As String
    Public discount_percentage As Double
    Public discounted_price_cents As Long

End Class

Public Class ClassNTKPayment

    Public total As Long
    Public total_paid As Long
    Public value As Long
    Public payment_method As String            'Dinheiro, Cartão de Crédito

End Class

Public Class ClassNTKOrder

    Public id As String
    Public order_number As String
    Public status As Integer
    Public status_description As String
    Public payment_type As String 'Dinheiro, Cartão, ONLINE (pré-pago no app), PAGAMENTO PARCIAL
    Public cpf As String
    Public troco As Long
    Public total As Long
    Public delivery_fee As Long
    Public items_description As String
    Public merchant_id As String
    Public order_date As String
    Public customer_name As String
    Public customer_phone As String
    Public address As ClassNTKAddress
    Public Items As List(Of ClassNTKItem)
    Public payments As List(Of ClassNTKPayment)

End Class

Public Class ClassNTKOrderExt
    Public order As ClassNTKOrder
End Class

Public Class ClassNTKOrders

    Public orders As List(Of ClassNTKOrder)

End Class

<ComClass(NTKIntegracao.ClassId, NTKIntegracao.InterfaceId, NTKIntegracao.EventsId)> _
Public Class NTKIntegracao

    Private Function ObtemTextoURL(ByVal sURL As String, ByRef Texto As String, ByRef ErrorMsg As String) As Long

        Dim lErro As Long = 1

        Try

            Dim webClient As New System.Net.WebClient
            webClient.Encoding = Encoding.UTF8

            Texto = webClient.DownloadString(sURL)

            lErro = SUCESSO

        Catch ex As Exception

            ErrorMsg = ex.Message & " " & sURL & "."

        Finally

            ObtemTextoURL = lErro

        End Try

    End Function

    Private Function CopiarPedido(ByVal objPedido As ClassNTKOrder, ByRef objOrder As GlobaisLoja.ClassNTKOrder) As Long
        'copia informacoes de objPedido para objOrder

        Try

            Dim lErro As Long = 1

            Dim objItem As ClassNTKItem
            Dim objPagamento As ClassNTKPayment
            Dim objComplemento As ClassNTKItem

            Dim objItem2 As GlobaisLoja.ClassNTKItem
            Dim objPayment As GlobaisLoja.ClassNTKPayment
            Dim objComplement As GlobaisLoja.ClassNTKItem

            Dim bCompleto As Boolean = False

            objOrder = New GlobaisLoja.ClassNTKOrder

            With objOrder

                .id = objPedido.id
                .order_number = objPedido.order_number
                .status = objPedido.status
                .order_date = CDate(Left(objPedido.order_date, 10))
                .order_time = TimeValue(Right(objPedido.order_date, 5) & ":00").ToOADate

                If Not (objPedido.address Is Nothing) Then

                    bCompleto = True

                    .payment_type = objPedido.payment_type
                    .cpf = Replace(Replace(Replace(objPedido.cpf, "-", ""), ".", ""), "/", "")
                    .troco = objPedido.troco
                    .total = objPedido.total
                    .delivery_fee = objPedido.delivery_fee
                    .customer_name = objPedido.customer_name
                    .customer_phone = objPedido.customer_phone
                    .address = New GlobaisLoja.ClassNTKAddress
                    .address.street = objPedido.address.street
                    .address.number = objPedido.address.number
                    .address.complement = objPedido.address.complement
                    .address.neighborhood = objPedido.address.neighborhood
                    .address.city = objPedido.address.city
                    .address.uf = objPedido.address.uf
                    .address.zipcode = objPedido.address.zipcode
                    .address.reference = objPedido.address.reference

                End If

            End With

            If bCompleto Then

                For Each objItem In objPedido.Items

                    objItem2 = New GlobaisLoja.ClassNTKItem

                    With objItem2

                        .codigo = objItem.codigo
                        If .codigo = "" Then .codigo = objItem.id
                        .Name = objItem.name
                        .qtd = objItem.qtd
                        .price_cents = objItem.price_cents
                        .subtotal_cents = objItem.subtotal_cents
                        .obs = objItem.obs

                    End With

                    For Each objComplemento In objItem.complements

                        objComplement = New GlobaisLoja.ClassNTKItem

                        With objComplement

                            .Name = objComplemento.name
                            .obs = objComplemento.obs

                        End With

                        objItem2.complements.Add(objComplement)

                    Next

                    objOrder.items.Add(objItem2)

                Next

                For Each objPagamento In objPedido.payments

                    objPayment = New GlobaisLoja.ClassNTKPayment

                    With objPayment

                        .total = objPagamento.total
                        .total_paid = objPagamento.total_paid
                        .value = objPagamento.value
                        .payment_method = objPagamento.payment_method

                    End With

                    objOrder.payments.Add(objPayment)

                Next

            End If

            CopiarPedido = SUCESSO

        Catch ex As Exception

            MsgBox(ex.Message)

            CopiarPedido = 1

        Finally

        End Try

    End Function

    Public Function ObterPedido(ByVal sToken As String, ByVal sURLBase As String, ByVal sOrderId As String, ByRef objOrder As GlobaisLoja.ClassNTKOrder) As Long

        Try

            Dim Texto As String = "", ErrorMsg As String
            Dim lErro As Long = 1, sURL As String
            Dim iTentativas As Integer

            sURL = sURLBase & "/api/ext/orders/" & sOrderId & "/status?token=" & sToken
            ErrorMsg = "Erro na obtenção do pedido " & sURL

            For iTentativas = 30 To 0 Step -1

                lErro = ObtemTextoURL(sURL, Texto, ErrorMsg)
                If lErro = SUCESSO Then Exit For

                Sleep(2000)

            Next

            If lErro <> SUCESSO Then Throw New System.Exception(ErrorMsg)

            Texto = Replace(Texto, """value"":null", """value"":0")
            Texto = Replace(Texto, "null", """""")

            Dim objPedido As New ClassNTKOrder
            Dim order As New ClassNTKOrderExt

            Dim JS1 As JavaScriptSerializer = New JavaScriptSerializer

            If InStr(Texto, "{""order"":") = 0 Then
                Texto = "{""order"":" & Texto & "}"
            End If
            order = JS1.Deserialize(Of ClassNTKOrderExt)(Texto)

            objPedido = order.order

            objOrder = New GlobaisLoja.ClassNTKOrder

            lErro = CopiarPedido(objPedido, objOrder)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            ObterPedido = SUCESSO

        Catch ex As Exception

            MsgBox(ex.Message)

            ObterPedido = 1

        Finally

        End Try

    End Function

    Public Function ObterPedidos(ByVal sToken As String, ByVal sMerchantId As String, ByVal sURLBase As String, ByVal sFiltro As String, ByRef colPedidos As VBA.Collection) As Long

        Try

            Dim Texto As String = "", ErrorMsg As String
            Dim lErro As Long = 1, sURL As String
            Dim iTentativas As Integer

            sURL = sURLBase & "/api/ext/orders?token=" & sToken & "&merchant_id=" & sMerchantId & sFiltro & "&order_by_asc=0"

            ErrorMsg = "Erro na obtenção de pedidos de " & sURL

            For iTentativas = 30 To 0 Step -1

                lErro = ObtemTextoURL(sURL, Texto, ErrorMsg)
                If lErro = SUCESSO Then Exit For

                Sleep(2000)

            Next

            If lErro <> SUCESSO Then Throw New System.Exception(ErrorMsg)

            Texto = Replace(Texto, """value"":null", """value"":0")
            Texto = Replace(Texto, "null", """""")

            Dim objPedidos As New ClassNTKOrders
            Dim objPedido As New ClassNTKOrder

            Dim objOrder As GlobaisLoja.ClassNTKOrder

            Dim JS1 As JavaScriptSerializer = New JavaScriptSerializer
            objPedidos = JS1.Deserialize(Of ClassNTKOrders)(Texto)

            For Each objPedido In objPedidos.orders

                lErro = CopiarPedido(objPedido, objOrder)
                If lErro <> SUCESSO Then Throw New System.Exception("")

                colPedidos.Add(objOrder)

            Next

            ObterPedidos = SUCESSO

        Catch ex As Exception

            MsgBox(ex.Message)

            ObterPedidos = 1

        Finally

        End Try

    End Function

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "0CF83384-EA2D-4023-8761-7EB937D77809"
    Public Const InterfaceId As String = "4C56089B-8B1C-471a-BA98-BBDFF27A5360"
    Public Const EventsId As String = "E639AC19-E7DD-48b2-A4D7-62CD7EA5A06B"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

End Class


