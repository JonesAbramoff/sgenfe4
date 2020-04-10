Option Strict Off
Option Explicit On

Imports System
Imports System.IO
Imports System.Xml.Serialization
Imports System.Collections
Imports System.Collections.Generic
'''<remarks/>
<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.18020"), _
 System.SerializableAttribute(), _
 System.Diagnostics.DebuggerStepThroughAttribute(), _
 System.ComponentModel.DesignerCategoryAttribute("code"), _
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True), _
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="", IsNullable:=False)> _
Partial Public Class ClassiFoodwspdv
    '''<remarks/>
    Public status As SByte
    '''<remarks/>
    Public message As String
    '''<remarks/>
    Public pedido As ClassiFoodPed
    '''<remarks/>
    Public [date] As String
    '''<remarks/>
    <System.Xml.Serialization.XmlArrayItemAttribute("item", IsNullable:=False)> _
    Public list() As ClassiFoodItem
End Class

'''<remarks/>
<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.18020"), _
 System.SerializableAttribute(), _
 System.Diagnostics.DebuggerStepThroughAttribute(), _
 System.ComponentModel.DesignerCategoryAttribute("code"), _
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True)> _
Partial Public Class ClassiFoodPed
    '''<remarks/>
    Public codPedido As Long
    '''<remarks/>
    Public idPedidoCurto As Short
    '''<remarks/>
    Public togo As String
    '''<remarks/>
    Public dataEntrega As String
    '''<remarks/>
    Public vlrPratos As Single
    '''<remarks/>
    Public vlrTaxa As Single
    '''<remarks/>
    Public vlrDesconto As Single
    '''<remarks/>
    Public vlrTotal As Single
    '''<remarks/>
    Public obsPedido As String
    '''<remarks/>
    Public condicaoPgto As String
    '''<remarks/>
    <System.Xml.Serialization.XmlArrayItemAttribute("pagamento", IsNullable:=False)> _
    Public pagamentos() As ClassiFoodPagto
    '''<remarks/>
    <System.Xml.Serialization.XmlArrayItemAttribute("telefone", IsNullable:=False)> _
    Public telefones() As ClassiFoodFone
    '''<remarks/>
    Public vlrTroco As Single
    '''<remarks/>
    Public codCampanha As String
    '''<remarks/>
    Public idCliente As Integer
    '''<remarks/>
    Public nome As String
    '''<remarks/>
    Public email As String
    '''<remarks/>
    Public clienteNovo As String
    '''<remarks/>
    Public referenciaGuia As String
    '''<remarks/>
    Public referenciaXy As String
    '''<remarks/>
    Public tipoLogradouro As String
    '''<remarks/>
    Public logradouro As String
    '''<remarks/>
    Public bairro As String
    '''<remarks/>
    Public logradouroNum As SByte
    '''<remarks/>
    Public complemento As String
    '''<remarks/>
    Public referencia As String
    '''<remarks/>
    Public cidade As String
    '''<remarks/>
    Public estado As String
    '''<remarks/>
    Public pais As String
    '''<remarks/>
    Public cep As Integer
    '''<remarks/>
    Public locale As String
    '''<remarks/>
    Public codFornecedor As Short
    '''<remarks/>
    Public codEmpresa As String
    '''<remarks/>
    Public nomeFornecedor As String
    '''<remarks/>
    Public status As String
    '''<remarks/>
    Public dataAlteracaoStatus As String
    '''<remarks/>
    Public nomeAtendente As String
    '''<remarks/>
    Public dataPedidoComanda As String
    '''<remarks/>
    Public dataPrevista As String
    '''<remarks/>
    Public agendado As String
End Class

'''<remarks/>
<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.18020"), _
 System.SerializableAttribute(), _
 System.Diagnostics.DebuggerStepThroughAttribute(), _
 System.ComponentModel.DesignerCategoryAttribute("code"), _
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True)> _
Partial Public Class ClassiFoodPagto
    '''<remarks/>
    Public codFormaPagto As String
    '''<remarks/>
    Public descricaoFormaPagto As String
    '''<remarks/>
    Public codTipoCondPagto As SByte
    '''<remarks/>
    Public bundleKey As String
    '''<remarks/>
    Public valor As Single
End Class

'''<remarks/>
<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.18020"), _
 System.SerializableAttribute(), _
 System.Diagnostics.DebuggerStepThroughAttribute(), _
 System.ComponentModel.DesignerCategoryAttribute("code"), _
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True)> _
Partial Public Class ClassiFoodFone
    '''<remarks/>
    Public ddd As SByte
    '''<remarks/>
    Public numero As Integer
End Class

'''<remarks/>
<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.18020"), _
 System.SerializableAttribute(), _
 System.Diagnostics.DebuggerStepThroughAttribute(), _
 System.ComponentModel.DesignerCategoryAttribute("code"), _
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True)> _
Partial Public Class ClassiFoodItem
    '''<remarks/>
    Public sequencia As SByte
    '''<remarks/>
    Public quantidade As Single
    '''<remarks/>
    Public vlrUnitLiq As Single
    '''<remarks/>
    Public codCardapio As String
    '''<remarks/>
    Public codProdutoPdv As String
    '''<remarks/>
    Public descricao As String
    '''<remarks/>
    Public descricaoCardapio As String
    '''<remarks/>
    Public codTipoProdutoPdv As String
    '''<remarks/>
    Public vlrUnitBruto As Single
    '''<remarks/>
    Public descPromocional As Single
    '''<remarks/>
    Public descCampanha As Single
    '''<remarks/>
    Public codpai As String

End Class

<ComClass(IFoodIntegracao.ClassId, IFoodIntegracao.InterfaceId, IFoodIntegracao.EventsId)> _
Public Class IFoodIntegracao

    Const SUCESSO = 0

    Private Function CopiarPedido(ByVal objPedidoRet As ClassiFoodwspdv, ByRef objVb6Ped As GlobaisLoja.ClassiFoodPedido) As Long

        Dim lErro As Long = SUCESSO

        Try
            Dim objItem As ClassiFoodItem
            Dim objFone As ClassiFoodFone
            Dim objPedido As ClassiFoodPed
            Dim objPagto As ClassiFoodPagto

            Dim objVb6Item As GlobaisLoja.ClassiFoodItem
            Dim objVb6Fone As GlobaisLoja.ClassiFoodFone
            Dim objVb6Pagto As GlobaisLoja.ClassiFoodPagto

            objPedido = objPedidoRet.pedido

            objVb6Ped.agendado = objPedido.agendado
            objVb6Ped.bairro = objPedido.bairro
            objVb6Ped.cep = objPedido.cep
            objVb6Ped.Cidade = objPedido.cidade
            objVb6Ped.clienteNovo = objPedido.clienteNovo
            objVb6Ped.codCampanha = objPedido.codCampanha
            objVb6Ped.codEmpresa = objPedido.codEmpresa
            objVb6Ped.codFornecedor = objPedido.codFornecedor
            objVb6Ped.codPedido = objPedido.codPedido
            objVb6Ped.complemento = objPedido.complemento
            objVb6Ped.condicaoPgto = objPedido.condicaoPgto
            objVb6Ped.dataAlteracaoStatus = objPedido.dataAlteracaoStatus
            objVb6Ped.dataEntrega = objPedido.dataEntrega
            objVb6Ped.dataPedidoComanda = objPedido.dataPedidoComanda
            objVb6Ped.dataPrevista = objPedido.dataPrevista
            objVb6Ped.email = objPedido.email
            objVb6Ped.estado = objPedido.estado
            objVb6Ped.idCliente = objPedido.idCliente
            objVb6Ped.idPedidoCurto = objPedido.idPedidoCurto
            objVb6Ped.locale = objPedido.locale
            objVb6Ped.logradouro = objPedido.logradouro
            objVb6Ped.logradouroNum = objPedido.logradouroNum
            objVb6Ped.Nome = objPedido.nome
            objVb6Ped.nomeAtendente = objPedido.nomeAtendente
            objVb6Ped.nomeFornecedor = objPedido.nomeFornecedor
            objVb6Ped.obsPedido = objPedido.obsPedido
            objVb6Ped.pais = objPedido.pais
            objVb6Ped.referencia = objPedido.referencia
            objVb6Ped.referenciaGuia = objPedido.referenciaGuia
            objVb6Ped.referenciaXy = objPedido.referenciaXy
            objVb6Ped.Status = objPedido.status
            objVb6Ped.tipoLogradouro = objPedido.tipoLogradouro
            objVb6Ped.togo = objPedido.togo
            objVb6Ped.vlrDesconto = objPedido.vlrDesconto
            objVb6Ped.vlrPratos = objPedido.vlrPratos
            objVb6Ped.vlrTaxa = objPedido.vlrTaxa
            objVb6Ped.vlrTotal = objPedido.vlrTotal
            objVb6Ped.vlrTroco = objPedido.vlrTroco

            For Each objItem In objPedidoRet.list

                objVb6Item = New GlobaisLoja.ClassiFoodItem

                objVb6Item.codCardapio = objItem.codCardapio
                objVb6Item.codPai = objItem.codPai
                objVb6Item.codProdutoPdv = objItem.codProdutoPdv
                objVb6Item.codTipoProdutoPdv = objItem.codTipoProdutoPdv
                objVb6Item.descCampanha = objItem.descCampanha
                objVb6Item.descPromocional = objItem.descPromocional
                objVb6Item.Descricao = objItem.descricao
                objVb6Item.descricaoCardapio = objItem.descricaoCardapio
                objVb6Item.Quantidade = objItem.quantidade
                objVb6Item.sequencia = objItem.sequencia

                objVb6Ped.colItens.Add(objVb6Item)

            Next

            For Each objPagto In objPedido.pagamentos

                objVb6Pagto = New GlobaisLoja.ClassiFoodPagto

                objVb6Pagto.bundleKey = objPagto.bundleKey
                objVb6Pagto.codFormaPagto = objPagto.codFormaPagto
                objVb6Pagto.codTipoCondPagto = objPagto.codTipoCondPagto
                objVb6Pagto.DescricaoFormaPagto = objPagto.descricaoFormaPagto
                objVb6Pagto.Valor = objPagto.valor

                objVb6Ped.pagamentos.Add(objVb6Pagto)

            Next

            For Each objFone In objPedido.telefones

                objVb6Fone = New GlobaisLoja.ClassiFoodFone

                objVb6Fone.ddd = objFone.ddd
                objVb6Fone.numero = objFone.numero

                objVb6Ped.telefones.Add(objVb6Fone)

            Next

        Catch ex As Exception

            MsgBox(ex.Message)

            If lErro = SUCESSO Then lErro = 1

        Finally

            CopiarPedido = lErro

        End Try

    End Function

    Public Function Pedido_Importa(ByVal sArquivoXML As String, ByVal objVb6Ped As GlobaisLoja.ClassiFoodPedido) As Long

        Dim lErro As Long = SUCESSO

        Try

            Dim objPedido As New ClassiFoodwspdv
            Dim DocDados2 As IO.StreamReader
            Dim XMLString1 As String
            Dim XMLStreamDados = New MemoryStream(10000)
            Dim xDados1 As Byte()

            DocDados2 = New IO.StreamReader(sArquivoXML)
            XMLString1 = DocDados2.ReadToEnd
            DocDados2.Close()

            'Altera as tags porque 
            '1 -> o - no nome estava atrapalhando 
            '2 -> ter duas coisas diferentes com a tag response-body acabava fazendo com que os itens não fossem lidos
            XMLString1 = Replace(XMLString1, "wspdv-response", "ClassiFoodwspdv")
            XMLString1 = Replace(XMLString1, "response-status", "status")
            XMLString1 = Replace(XMLString1, "response-message", "message")
            XMLString1 = Replace(XMLString1, "response-body class=""pedido""", "pedido")
            XMLString1 = Replace(XMLString1, "response-date", "date")
            XMLString1 = Replace(XMLString1, "response-body class=""list""", "list")
            XMLString1 = Replace(XMLString1, "response-body", "pedido", , 1)
            XMLString1 = Replace(XMLString1, "response-body", "list", , 1)

            xDados1 = System.Text.Encoding.UTF8.GetBytes(XMLString1)

            XMLStreamDados.Write(xDados1, 0, xDados1.Length)

            Dim mySerializerImportaXML As New XmlSerializer(GetType(ClassiFoodwspdv))

            XMLStreamDados.Position = 0

            objPedido = mySerializerImportaXML.Deserialize(XMLStreamDados)

            lErro = CopiarPedido(objPedido, objVb6Ped)
            If lErro <> SUCESSO Then Throw New System.Exception("")

        Catch ex As Exception

            If Not ex.InnerException Is Nothing Then
                MsgBox(ex.Message & vbNewLine & ex.InnerException.Message)
            Else
                MsgBox(ex.Message)
            End If

            If lErro = SUCESSO Then lErro = 1

        Finally

            Pedido_Importa = lErro

        End Try

    End Function

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "3730F6C3-838E-4beb-BD6D-40A8AAE3CAF9"
    Public Const InterfaceId As String = "98B11B24-3E43-444d-B610-F6FA9C33E611"
    Public Const EventsId As String = "ED57465D-E1CE-47f0-B127-4D4519756F23"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

End Class


