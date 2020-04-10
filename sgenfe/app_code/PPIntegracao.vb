'classe para importacao de pedidos do aplicativo de pesquisa de precos

Imports System
Imports System.IO
Imports System.Web.Script.Serialization
Imports System.Collections
Imports System.Collections.Generic

Public Class ClassItemPedidoImportaInfo

    Public Seq As Integer
    Public ProdutoEan As String
    Public ProdutoCodigo As String
    Public Descricao As String
    Public Quantidade As Double
    Public PrecoUnitario As Double
    Public Observacao As String

End Class

Public Class ClassPedidoImportaInfo

    Public IdCarrinho As Long

    Public LojaCNPJ As String

    Public Status As Integer

    Public Email As String
    Public Nome As String

    Public DataRegistro As Date
    Public HoraRegistro As Double

    Public ValorTotal As Double

    Public TelConfirmacao As String
    Public Observacao As String
    Public EnderecoLogradouro As String
    Public EnderecoNumero As String
    Public EnderecoComplemento As String
    Public EnderecoBairro As String
    Public EnderecoCidade As String
    Public EnderecoUF As String
    Public EnderecoCEP As String
    Public EnderecoFone As String
    Public EnderecoReferencia As String

    Public TaxaEntrega As Double

    Public FormaPagto As Integer
    Public MeioPagto As Integer
    Public LocalPagto As Integer
    Public NumParcelas As Integer
    Public NumeroCartao As String
    Public Tid As String
    Public AutorizacaoARP As String
    Public NumeroBanco As String
    Public NumeroAgencia As String
    Public NumeroCheque As String
    Public TrocoPara As Double
    Public StatusPagto As Integer
    Public FormaEntrega As Integer

    Public colItens As New System.Collections.Generic.List(Of ClassItemPedidoImportaInfo)

End Class

Public Class ClassPedidoImporta

    Public CodLoja As Long
    Public LojaVirtual As Integer

End Class

Public Class ClassPedidoImportaRet

    Public objPedidoImportaInfo As New ClassPedidoImportaInfo

End Class

Public Class ClassPedidoConfirmaImportacao

    Public CodLoja As Long

    Public IdCarrinho As Long

End Class

Public Class ClassPedidoConfirmaImportacaoRet

    Public IdCarrinho As Long

End Class

<ComClass(PPIntegracao.ClassId, PPIntegracao.InterfaceId, PPIntegracao.EventsId)> _
Public Class PPIntegracao

    Private Function CopiarPedido(ByVal objPedido As ClassPedidoImportaInfo, ByRef objPPPedido As GlobaisLoja.ClassPPPedido) As Long
        'copia informacoes de objPedido para objOrder

        Dim lErro As Long = SUCESSO

        Try

            With objPPPedido

                .IdCarrinho = objPedido.IdCarrinho

                .DataRegistro = objPedido.DataRegistro
                .HoraRegistro = objPedido.HoraRegistro

                .Email = objPedido.Email
                .Nome = objPedido.Nome

                .ValorTotal = objPedido.ValorTotal
                .Observacao = objPedido.Observacao
                .EnderecoLogradouro = objPedido.EnderecoLogradouro
                .EnderecoNumero = objPedido.EnderecoNumero
                .EnderecoComplemento = objPedido.EnderecoComplemento
                .EnderecoBairro = objPedido.EnderecoBairro
                .EnderecoCidade = objPedido.EnderecoCidade
                .EnderecoUF = objPedido.EnderecoUF
                .EnderecoCEP = objPedido.EnderecoCEP
                .EnderecoFone = objPedido.EnderecoFone
                .EnderecoReferencia = objPedido.EnderecoReferencia
                .TaxaEntrega = objPedido.TaxaEntrega
                .FormaPagto = objPedido.FormaPagto
                .MeioPagto = objPedido.MeioPagto
                .NumeroCartao = objPedido.NumeroCartao
                .AutorizacaoARP = objPedido.AutorizacaoARP
                .NumeroBanco = objPedido.NumeroBanco
                .NumeroAgencia = objPedido.NumeroAgencia
                .NumeroCheque = objPedido.NumeroCheque
                .TrocoPara = objPedido.TrocoPara
                .StatusPagto = objPedido.StatusPagto
                .FormaEntrega = objPedido.FormaEntrega

            End With

            Dim objPPItem As GlobaisLoja.ClassPPItemPedido

            For Each objItem In objPedido.colItens

                objPPItem = New GlobaisLoja.ClassPPItemPedido

                With objPPItem

                    .Seq = objItem.Seq
                    .ProdutoEan = objItem.ProdutoEan
                    .ProdutoCodigo = objItem.ProdutoCodigo
                    .Descricao = objItem.Descricao
                    .Quantidade = objItem.Quantidade
                    .PrecoUnitario = objItem.PrecoUnitario
                    .Observacao = objItem.Observacao

                End With

                objPPPedido.colItens.Add(objPPItem)

            Next

        Catch ex As Exception

            'MsgBox(ex.Message)

            If lErro = SUCESSO Then lErro = 1

        Finally

            CopiarPedido = lErro

        End Try

    End Function

    Public Function Pedido_Importa(ByVal Ambiente As String, ByVal Versao As String, ByVal EmailLojista As String, ByVal SenhaLojista As String, ByVal CodLoja As Long, ByRef objPPPedido As GlobaisLoja.ClassPPPedido) As Long
        'obter proximo carrinho a ser importado

        Dim lErro As Long = SUCESSO

        Try

            Dim xmlEntrada As String = ""
            Dim xmlSaida As String = ""
            Dim xmlErro As String = ""

            Dim objPedidoImporta As New ClassPedidoImporta

            objPedidoImporta.CodLoja = CodLoja
            objPedidoImporta.LojaVirtual = 1

            lErro = Serializar(GetType(ClassPedidoImporta), objPedidoImporta, xmlEntrada)
            If lErro <> SUCESSO Then Throw New System.Exception("Erro na serializacao de ClassPedidoImporta")

            Dim objPP As New pesquisadepreco.IntegracaoLojista

            lErro = objPP.FuncaoGenerica2(Versao, "Pedido_Importa", xmlEntrada, xmlSaida, xmlErro, Ambiente, EmailLojista, SenhaLojista)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            Dim objPedidoImportaRet As New ClassPedidoImportaRet

            lErro = Deserializar(xmlSaida, GetType(ClassPedidoImportaRet), objPedidoImportaRet)
            If lErro <> SUCESSO Then Throw New System.Exception("Erro na deserializacao de " & xmlSaida)

            lErro = CopiarPedido(objPedidoImportaRet.objPedidoImportaInfo, objPPPedido)
            If lErro <> SUCESSO Then Throw New System.Exception("")

        Catch ex As Exception

            'MsgBox(ex.Message)

            If lErro = SUCESSO Then lErro = 1

        Finally

            Pedido_Importa = lErro

        End Try

    End Function

    Public Function Pedido_Confirma_Importacao(ByVal Ambiente As String, ByVal Versao As String, ByVal EmailLojista As String, ByVal SenhaLojista As String, ByVal CodLoja As Long, ByVal IdCarrinho As Long) As Long
        'confirmar que o pedido foi importado corretamente para o bd de orçamentos do Corporator

        Dim lErro As Long = SUCESSO

        Try

            Dim xmlEntrada As String = ""
            Dim xmlSaida As String = ""
            Dim xmlErro As String = ""

            Dim objPedidoConfirmaImportacao As New ClassPedidoConfirmaImportacao

            objPedidoConfirmaImportacao.CodLoja = CodLoja
            objPedidoConfirmaImportacao.IdCarrinho = IdCarrinho

            lErro = Serializar(GetType(ClassPedidoConfirmaImportacao), objPedidoConfirmaImportacao, xmlEntrada)
            If lErro <> SUCESSO Then Throw New System.Exception("Erro na serializacao de ClassPedidoConfirmaImportacao")

            Dim objPP As New pesquisadepreco.IntegracaoLojista

            lErro = objPP.FuncaoGenerica2(Versao, "Pedido_Confirma_Importacao", xmlEntrada, xmlSaida, xmlErro, Ambiente, EmailLojista, SenhaLojista)
            If lErro <> SUCESSO Then Throw New System.Exception("")

        Catch ex As Exception

            'MsgBox(ex.Message)

            If lErro = SUCESSO Then lErro = 1

        Finally

            Pedido_Confirma_Importacao = lErro

        End Try

    End Function

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "4CB832F6-3270-481f-A0EC-A9D72FFFDE85"
    Public Const InterfaceId As String = "E4E0D907-B7D4-45c5-AC9D-15F54601CCF9"
    Public Const EventsId As String = "9CEADE9F-2DD9-43ab-B5D0-B0BBE32D9F6F"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

End Class


