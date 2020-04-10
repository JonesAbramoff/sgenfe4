Imports System
Imports sgenfe4.NFeXsd

<ComClass(NFeEnvio.ClassId, NFeEnvio.InterfaceId, NFeEnvio.EventsId)>
Public Class NFeEnvio

    Private sMetodo As String
    Private gobjApp As ClassGlobalApp
    Private gobjVenda As GlobaisLoja.ClassVenda
    Private sArquivoLoteEnvio As String
    Private lNumNotaFiscal As Long
    Private sSerie As String
    Private lLote As Long
    Private sModelo As String ' "NFe" ou "NFCe"
    Private bComISSQN As Boolean
    Private bComICMS As Boolean

    'para qrcode
    Private scDest As String
    Private sdhEmi As String
    Private svNF As String
    Private svICMS As String

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "6FD89F22-8C60-40a3-B355-5A44C2E7B683"
    Public Const InterfaceId As String = "63BE0341-CAFC-4fd0-9C76-C809ED1CCE95"
    Public Const EventsId As String = "64C3F9D2-5E0A-4259-B395-A7F83676E11C"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
        'MsgBox("new sgenfe")
        'gobjApp.GravarLog("New NFeEnvio")
    End Sub

    Public Declare Sub Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)

    Private Function Email_Monta_Texto_Venda(ByRef sTexto As String) As Long

        Try

            sTexto = gobjVenda.objNFeInfo.sEmitRazaoSocial & " - " & gobjVenda.objNFeInfo.sEmitNomeReduzido & vbCrLf
            sTexto = sTexto & vbCrLf & vbCrLf
            sTexto = sTexto & "Nota Fiscal de Consumidor Eletronica - NFCe" & vbCrLf
            sTexto = sTexto & vbCrLf & vbCrLf
            sTexto = sTexto & "Link para consulta pelo QRCode: " & gobjVenda.objCupomFiscal.sNFCeQRCode & vbCrLf
            sTexto = sTexto & vbCrLf & vbCrLf
            sTexto = sTexto & "Chave de Acesso: " & gobjVenda.objCupomFiscal.sNFeChaveAcesso & vbCrLf
            sTexto = sTexto & vbCrLf & vbCrLf
            sTexto = sTexto & "Emitida através do Sistema Corporator"

            Email_Monta_Texto_Venda = SUCESSO

        Catch ex As Exception

            Email_Monta_Texto_Venda = 1

        Finally

        End Try

    End Function

    Private Function Lote_Prepara(ByVal XMLStringNFes As String, ByRef XMLString2 As String) As Long

        Try

            Dim envioNFe As TEnviNFe = New TEnviNFe
            Dim XMLStream1 As MemoryStream = New MemoryStream(10000)
            Dim XMLString1 As String
            Dim iPos As Integer
            Dim XMLStreamDados As MemoryStream = New MemoryStream(10000)

            envioNFe.versao = NFE_VERSAO_XML
            envioNFe.idLote = lLote

            If sModelo = "NFCe" Then envioNFe.indSinc = TEnviNFeIndSinc.Item1

            Dim mySerializerw As New XmlSerializer(GetType(TEnviNFe))

            XMLStream1 = New MemoryStream(10000)

            mySerializerw.Serialize(XMLStream1, envioNFe)

            Dim xmw As Byte()
            xmw = XMLStream1.ToArray

            XMLString1 = System.Text.Encoding.UTF8.GetString(xmw)

            XMLString2 = Mid(XMLString1, 1, Len(XMLString1) - 10) & XMLStringNFes & Mid(XMLString1, Len(XMLString1) - 10)

            XMLString2 = Mid(XMLString2, 1, 19) & " encoding=""UTF-8"" " & Mid(XMLString2, 20)

            iPos = InStr(XMLString2, "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema""")

            If iPos <> 0 Then

                XMLString2 = Mid(XMLString2, 1, iPos - 1) & Mid(XMLString2, iPos + 99)

            End If


            '************* valida dados antes do envio **********************
            Dim xDados As Byte()

            xDados = System.Text.Encoding.UTF8.GetBytes(XMLString2)

            XMLStreamDados = New MemoryStream(10000)

            XMLStreamDados.Write(xDados, 0, xDados.Length)

            Dim DocDados As XmlDocument = New XmlDocument
            XMLStreamDados.Position = 0
            DocDados.Load(XMLStreamDados)
            sArquivoLoteEnvio = gobjApp.sDirXml & sModelo & "_" & CStr(gobjVenda.objCupomFiscal.iFilialEmpresa) & "_" & CStr(gobjVenda.objCupomFiscal.iCodCaixa) & "_" & envioNFe.idLote & "-env-lot.xml"
            DocDados.Save(sArquivoLoteEnvio)

            Lote_Prepara = SUCESSO

        Catch ex As Exception

            Lote_Prepara = 1

            Call gobjApp.GravarLog(ex.Message)

        End Try

    End Function

    Private Function Lote_Envia(ByVal DocDados1 As XmlDocument, ByRef objRetEnviNFE As TRetEnviNFe) As Long

        Dim lErro As Long, XMLStringRetEnvNFE As String = ""

        Try

            Dim NfeAutorizacao As New nfeautorizacao.NFeAutorizacao4
            'Dim NFecabec_EnvioLote As Object New nfeautorizacao.nfeCabecMsg

            Call gobjApp.GravarLog("Chamando Lote_Envia1", "INFO", False, False, True, False, "")

            lErro = Lote_Envia1(NfeAutorizacao, XMLStringRetEnvNFE, DocDados1)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            Call gobjApp.GravarLog("Chamando Lote_Envia2", "INFO", False, False, True, False, "")

            lErro = Lote_Envia2(XMLStringRetEnvNFE, objRetEnviNFE)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            Lote_Envia = SUCESSO

        Catch ex As Exception

            Lote_Envia = 1

            Call gobjApp.GravarLog(ex.Message)

            If Len(Trim(XMLStringRetEnvNFE)) > 0 Then Call gobjApp.GravarLog(XMLStringRetEnvNFE, "ERRO", False, False, False, True, "ERRO_Lote_Envia_" & Year(Now) & Month(Now) & Day(Now) & Hour(Now) & Minute(Now) & Second(Now) & ".xml")

        End Try

    End Function

    Private Function Lote_Envia1(ByVal NfeAutorizacao As nfeautorizacao.NFeAutorizacao4, ByRef XMLStringRetEnvNFE As String, ByVal DocDados1 As XmlDocument) As Long

        Try

            Dim xmlNode1 As XmlNode

            'NFecabec_EnvioLote.cUF = CStr(gobjApp.iUFCodIBGE)
            'NFecabec_EnvioLote.versaoDados = NFE_VERSAO_XML

            'NfeAutorizacao.nfeCabecMsgValue = NFecabec_EnvioLote


            If gobjApp.iDebug = 1 Then MsgBox("39.6")
            gobjApp.sErro = "39.6"
            gobjApp.sMsg1 = "vai enviar a nota"

            Dim sURL As String
            sURL = ""
            Call WS_Obter_URL(sURL, gobjApp.iNFeAmbiente = NFE_AMBIENTE_HOMOLOGACAO, gobjApp.sSiglaEstado, "NFeAutorizacao", sModelo)

            NfeAutorizacao.Url = sURL

            NfeAutorizacao.ClientCertificates.Add(gobjApp.cert)
            xmlNode1 = NfeAutorizacao.nfeAutorizacaoLote(DocDados1)

            XMLStringRetEnvNFE = xmlNode1.OuterXml

            If gobjApp.iDebug = 1 Then
                MsgBox("39.7")
                MsgBox(XMLStringRetEnvNFE)

            End If
            gobjApp.sErro = "39.7"
            gobjApp.sMsg1 = "enviou a nota"

            Lote_Envia1 = SUCESSO

        Catch ex As Exception

            Lote_Envia1 = 1

            Call gobjApp.GravarLog(ex.Message)

        End Try

    End Function

    Private Function Lote_Envia2(ByVal XMLStringRetEnvNFE As String, ByRef objRetEnviNFE As TRetEnviNFe) As Long

        Try

            Dim XMLStreamRet As MemoryStream = New MemoryStream(10000)
            Dim xRet As Byte()

            xRet = System.Text.Encoding.UTF8.GetBytes(XMLStringRetEnvNFE)

            XMLStreamRet = New MemoryStream(10000)

            XMLStreamRet.Write(xRet, 0, xRet.Length)

            Dim mySerializerRetEnvNFe As New XmlSerializer(GetType(TRetEnviNFe))

            XMLStreamRet.Position = 0

            objRetEnviNFE = mySerializerRetEnvNFe.Deserialize(XMLStreamRet)

            gobjApp.GravarLog("Retorno do envio do lote - " & objRetEnviNFE.xMotivo)

            Lote_Envia2 = SUCESSO

        Catch ex As Exception

            Lote_Envia2 = 1

            Call gobjApp.GravarLog(ex.Message)

        End Try

    End Function

    Private Function Lote_Consulta(ByVal objRetEnviNFE As TRetEnviNFe) As Long

        Dim XMLStringRetConsReciNFE As String = ""

        Try

            Dim NFeRetAutorizacao As New nferetautorizacao.NFeRetAutorizacao4
            Dim DocDados1 = New XmlDocument
            Dim lErro As Long, bAutorizou As Boolean
            Dim xmlNode1 As XmlNode
            Dim xRet As Byte(), objconsRetReciNFe As TRetConsReciNFe
            Dim XMLStreamRet As MemoryStream
            Dim objconsReciNFe As TConsReciNFe = New TConsReciNFe

            Dim bSincrono As Boolean

            Call gobjApp.GravarLog("Chamando Lote_Consulta1", "INFO", False, False, True, False, "")

            lErro = Lote_Consulta1(objRetEnviNFE, objconsReciNFe, NFeRetAutorizacao, DocDados1, bSincrono)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            If Not bSincrono Then

                Dim i1 As Integer

                i1 = 1

                Do While i1 < 11

                    Sleep(2000)

                    Call gobjApp.GravarLog("Chamando Lote_Consulta2", "INFO", False, False, True, False, "")

                    lErro = Lote_Consulta2(NFeRetAutorizacao)
                    If lErro <> SUCESSO Then Throw New System.Exception("")

                    NFeRetAutorizacao.ClientCertificates.Add(gobjApp.cert)
                    xmlNode1 = NFeRetAutorizacao.nfeRetAutorizacaoLote(DocDados1)

                    XMLStringRetConsReciNFE = xmlNode1.OuterXml

                    xRet = System.Text.Encoding.UTF8.GetBytes(XMLStringRetConsReciNFE)

                    XMLStreamRet = New MemoryStream(10000)
                    XMLStreamRet.Write(xRet, 0, xRet.Length)

                    Dim mySerializerRetConsReciNFe As New XmlSerializer(GetType(TRetConsReciNFe))

                    objconsRetReciNFe = New TRetConsReciNFe

                    XMLStreamRet.Position = 0

                    objconsRetReciNFe = mySerializerRetConsReciNFe.Deserialize(XMLStreamRet)

                    If objconsRetReciNFe.cStat = "105" Then

                        gobjApp.GravarLog("Lote em Processamento - Aguarde - Tentativa " & i1 & "/10")

                    Else

                        Call gobjApp.GravarLog("Chamando Lote_Consulta3", "INFO", False, False, True, False, "")

                        lErro = Lote_Consulta3(objconsRetReciNFe, bAutorizou)
                        If lErro <> SUCESSO Then Throw New System.Exception("")

                        If bAutorizou Then Exit Do

                    End If

                    i1 = i1 + 1

                Loop

                If i1 = 11 Then

                    gobjApp.GravarLog("Tente a consulta a este lote mais tarde")
                    Throw New System.Exception("")

                End If

            End If

            Lote_Consulta = SUCESSO

        Catch ex As Exception

            Lote_Consulta = 1

            Call gobjApp.GravarLog(ex.Message)

            If Len(Trim(XMLStringRetConsReciNFE)) > 0 Then Call gobjApp.GravarLog(XMLStringRetConsReciNFE, "ERRO", False, False, False, True, "ERRO_Lote_Consulta_" & Year(Now) & Month(Now) & Day(Now) & Hour(Now) & Minute(Now) & Second(Now) & ".xml")

        End Try

    End Function

    Private Function Lote_Consulta2(ByVal NFeRetAutorizacao As nferetautorizacao.NFeRetAutorizacao4) As Long

        Try

            Dim sURL As String = ""

            Call WS_Obter_URL(sURL, gobjApp.iNFeAmbiente = NFE_AMBIENTE_HOMOLOGACAO, gobjApp.sSiglaEstado, "NFeRetAutorizacao", sModelo)

            NFeRetAutorizacao.Url = sURL

            Lote_Consulta2 = SUCESSO

        Catch ex As Exception

            Lote_Consulta2 = 1

            Call gobjApp.GravarLog(ex.Message)

        End Try

    End Function

    Private Function protNFe_Processa(ByVal objProtNFe As TProtNFe, ByRef bAutorizou As Boolean) As Long

        Try

            Dim sArquivo As String, XMLString1 As String, XMLStream1 As MemoryStream
            Dim XMLString3 As String, XMLString As String, iPos As Integer
            Dim XMLStreamDados As MemoryStream, sChaveNFe As String, sArq As String, lErro As Long

            bAutorizou = False

            If String.IsNullOrEmpty(objProtNFe.infProt.nProt) Then
                objProtNFe.infProt.nProt = ""
            End If

            If sMetodo = "Envio" Then

                If Format(CInt(sSerie), "000") <> Mid(objProtNFe.infProt.chNFe, 23, 3) Or
                       Format(lNumNotaFiscal, "000000000") <> Mid(objProtNFe.infProt.chNFe, 26, 9) Then
                    Throw New System.Exception("A nota nao corresponde à nota enviada. Serie Consulta = " & Mid(objProtNFe.infProt.chNFe, 23, 3) & " Numero Consulta = " & Mid(objProtNFe.infProt.chNFe, 26, 9))
                End If

            End If

            gobjApp.GravarLog("Nota Fiscal  - " & objProtNFe.infProt.chNFe & " - " & objProtNFe.infProt.xMotivo)

            If objProtNFe.infProt.cStat = "100" Or objProtNFe.infProt.cStat = "150" Then

                If objProtNFe.versao <> "1.10" And objProtNFe.versao <> "2.00" And objProtNFe.versao <> "3.10" And objProtNFe.versao <> "4.00" Then
                    Throw New System.Exception("Versao não tratada. Versao = " & objProtNFe.versao)
                End If

                Dim DocDados2 As XmlDocument = New XmlDocument
                Dim XMLStreamDados1 = New MemoryStream(10000)

                sArquivo = gobjApp.sDirXml & objProtNFe.infProt.chNFe & "-pre.xml"
                DocDados2.Load(sArquivo)
                DocDados2.Save(XMLStreamDados1)

                Dim xm As Byte()

                'pega a parte do xml que fica entre <NFe> e </NFe>
                xm = XMLStreamDados1.ToArray

                XMLString1 = System.Text.Encoding.UTF8.GetString(xm)

                'cria uma versao do que vai ser armazenado somente com o que é possivel, ou seja,
                'versao e protNFE. O <NFe> vai ser inserido depois (XMLString1)
                Dim objNFeProc As TNfeProc = New TNfeProc

                'objNFeProc.NFe = objNFe
                objNFeProc.versao = NFE_VERSAO_XML
                objNFeProc.protNFe = objProtNFe

                Dim mySerializer As New XmlSerializer(GetType(TNfeProc))

                XMLStream1 = New MemoryStream(10000)

                mySerializer.Serialize(XMLStream1, objNFeProc)

                Dim xm3 As Byte()
                xm3 = XMLStream1.ToArray

                XMLString3 = System.Text.Encoding.UTF8.GetString(xm3)


                iPos = InStr(XMLString3, "<protNFe")

                'criacao da string completa
                XMLString = Mid(XMLString3, 1, iPos - 1) & XMLString1 & Mid(XMLString3, iPos)

                XMLString = Mid(XMLString, 1, 19) & " encoding=""UTF-8"" " & Mid(XMLString, 20)


                iPos = InStr(XMLString, "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema""")

                If iPos <> 0 Then

                    XMLString = Mid(XMLString, 1, iPos - 1) & Mid(XMLString, iPos + 99)

                End If


                XMLStreamDados = New MemoryStream(10000)

                Dim xDados1 As Byte()

                xDados1 = System.Text.Encoding.UTF8.GetBytes(XMLString)

                XMLStreamDados.Write(xDados1, 0, xDados1.Length)


                Dim DocDados3 As XmlDocument = New XmlDocument

                XMLStreamDados.Position = 0
                DocDados3.Load(XMLStreamDados)
                'sArquivo = sDir & objNFe.infNFe.Id & ".xml"
                sArquivo = gobjApp.sDirXml & objProtNFe.infProt.chNFe & "-procNfe.xml"
                '                                    DocDados3.Save(sArquivo)

                Dim writer As New XmlTextWriter(sArquivo, Nothing)

                writer.Formatting = Formatting.None
                DocDados3.WriteTo(writer)
                writer.Close()

                If sMetodo = "Envio" Then

                    gobjVenda.objCupomFiscal.sNFeChaveAcesso = objProtNFe.infProt.chNFe
                    gobjVenda.objCupomFiscal.sNFeArqXml = sArquivo
                    gobjVenda.objCupomFiscal.lNumero = lNumNotaFiscal
                    gobjVenda.objCupomFiscal.lCCF = lNumNotaFiscal
                    gobjVenda.objCupomFiscal.sNumSerieECF = sSerie

                    gobjVenda.objCupomFiscal.sNFecStat = objProtNFe.infProt.cStat
                    gobjVenda.objCupomFiscal.sNFenProt = objProtNFe.infProt.nProt
                    gobjVenda.objCupomFiscal.dtNFeData = UTCParaDate(objProtNFe.infProt.dhRecbto)
                    gobjVenda.objCupomFiscal.dNFeHora = UTCParaHora(objProtNFe.infProt.dhRecbto)
                    gobjVenda.objCupomFiscal.sNFEversao = NFE_VERSAO_XML
                    gobjVenda.objCupomFiscal.iNFetpAmb = CInt(GetXmlAttrNameFromEnumValue(Of TAmb)(objProtNFe.infProt.tpAmb))

                    Dim sDigVal As String

                    Dim ns As New XmlNamespaceManager(DocDados3.NameTable)
                    ns.AddNamespace("nfe", "http://www.portalfiscal.inf.br/nfe")
                    Dim xpathNav As XPathNavigator = DocDados3.CreateNavigator()
                    Dim node As XPathNavigator = xpathNav.SelectSingleNode("//nfe:protNFe/nfe:infProt/nfe:digVal", ns)
                    sDigVal = node.InnerXml

                    'TODO: Trocada chamada do QRCode Online 1
                    gobjApp.GravarLog("protNFe_Processa NFCE_Gera_QRCode2_Online",, False)
                    'gobjVenda.objCupomFiscal.sNFCeQRCode = NFCE_Gera_QRCode(objProtNFe.infProt.chNFe, "100", GetXmlAttrNameFromEnumValue(Of TAmb)(objProtNFe.infProt.tpAmb), scDest, sdhEmi, svNF, svICMS, sDigVal, gobjVenda.objNFeInfo.sidNFCECSC, gobjVenda.objNFeInfo.sNFCECSC)

                    gobjVenda.objCupomFiscal.sNFCeQRCode = NFCE_Gera_QRCode2_Online(gobjApp, objProtNFe.infProt.chNFe, GetXmlAttrNameFromEnumValue(Of TAmb)(objProtNFe.infProt.tpAmb), gobjVenda.objNFeInfo.sidNFCECSC, gobjVenda.objNFeInfo.sNFCECSC)
                Else

                    'envio do xml gerado em contingencia offline

                    gobjVenda.objCupomFiscal.sNFeChaveAcesso = objProtNFe.infProt.chNFe
                    gobjVenda.objCupomFiscal.sNFeArqXml = sArquivo
                    gobjVenda.objCupomFiscal.sNFecStat = objProtNFe.infProt.cStat
                    gobjVenda.objCupomFiscal.sNFenProt = objProtNFe.infProt.nProt
                    gobjVenda.objCupomFiscal.dtNFeData = UTCParaDate(objProtNFe.infProt.dhRecbto)
                    gobjVenda.objCupomFiscal.dNFeHora = UTCParaHora(objProtNFe.infProt.dhRecbto)
                    gobjVenda.objCupomFiscal.iNFetpAmb = CInt(GetXmlAttrNameFromEnumValue(Of TAmb)(objProtNFe.infProt.tpAmb))

                End If

                bAutorizou = True

                'se for duplicidade pesquisa a chave e faz consulta
            ElseIf objProtNFe.infProt.cStat = "204" Or objProtNFe.infProt.cStat = "539" Then

                Dim colArq As New Collection

                sChaveNFe = Left(objProtNFe.infProt.chNFe, 34)

                sArq = My.Computer.FileSystem.GetName(gobjApp.sDirXml & sChaveNFe & "*-pre.xml")

                sArq = Dir(gobjApp.sDirXml & sArq)

                Do While sArq <> ""

                    colArq.Add(sArq)

                    sArq = Dir()

                Loop


                For Each sArq In colArq

                    sChaveNFe = Left(sArq, 44)

                    lErro = Consulta_NFe(sChaveNFe, bAutorizou)
                    If lErro <> SUCESSO Then Throw New System.Exception("")

                    If bAutorizou Then Exit For

                Next

                'se for uma nota denegada vai guardar a informacao em NFeFedDenegada
            ElseIf (objProtNFe.infProt.cStat = "205" Or objProtNFe.infProt.cStat = "110" Or objProtNFe.infProt.cStat = "301" Or objProtNFe.infProt.cStat = "302") Then



                If objProtNFe.versao <> "1.10" And objProtNFe.versao <> "2.00" And objProtNFe.versao <> "3.10" And objProtNFe.versao <> "4.00" Then
                    Throw New System.Exception("Versao não tratada. Versao = " & objProtNFe.versao)
                End If

                '??? lErro = objConsultaNFe.Consulta_NFe(sEmpresa, objProtNFe.infProt.chNFe, iFilialEmpresa, 0)

            Else

                gobjVenda.objCupomFiscal.bEditavel = True
                Throw New System.Exception("")

            End If

            protNFe_Processa = SUCESSO

        Catch ex As Exception

            protNFe_Processa = 1

        Finally

        End Try

    End Function

    Private Function Lote_Consulta3(ByVal objconsRetReciNFe As TRetConsReciNFe, ByRef bAutorizou As Boolean) As Long

        Try
            Dim lErro As Long

            bAutorizou = False

            gobjApp.GravarLog("Retorno da consulta do lote - " & objconsRetReciNFe.xMotivo & IIf(objconsRetReciNFe.cMsg <> "0", " - código de Msg = " & objconsRetReciNFe.cMsg & " - " & objconsRetReciNFe.xMsg, ""))

            If Not objconsRetReciNFe.protNFe Is Nothing Then

                For i = 0 To objconsRetReciNFe.protNFe.Length - 1

                    Dim objProtNFe As TProtNFe
                    objProtNFe = objconsRetReciNFe.protNFe(i)

                    lErro = protNFe_Processa(objProtNFe, bAutorizou)
                    If lErro <> SUCESSO Then Throw New System.Exception("")

                Next

            End If

            Lote_Consulta3 = SUCESSO

        Catch ex As Exception

            Lote_Consulta3 = 1

            Call gobjApp.GravarLog(ex.Message)

        End Try

    End Function

    Private Function Lote_Consulta1(ByVal objRetEnviNFE As TRetEnviNFe, ByVal objconsReciNFe As TConsReciNFe, ByVal NFeRetAutorizacao As nferetautorizacao.NFeRetAutorizacao4, ByRef DocDados1 As XmlDocument, ByRef bSincrono As Boolean) As Long

        Try

            Dim XMLStream1 As MemoryStream
            Dim XMLString1 As String
            Dim iPos As Integer, lErro As Long
            'Dim NFecabec_RetEnvioLote As New nferetautorizacao.nfeCabecMsg

            Dim snRec As String
            Dim infRec As TRetEnviNFeInfRec

            Dim sAux As String

            If objRetEnviNFE.cStat = 225 Then gobjVenda.objCupomFiscal.bEditavel = True
            sAux = objRetEnviNFE.Item.GetType.FullName.ToString
            Do While InStr(sAux, ".") <> 0
                sAux = Mid(sAux, InStr(sAux, ".") + 1)
            Loop

            gobjApp.GravarLog("Lote_Consulta1: objRetEnviNFE.Item.GetType.FullName.ToString = " & sAux, "INFO", True, True, True, False, "")

            Select Case sAux

                Case "TRetEnviNFeInfRec"
                    infRec = objRetEnviNFE.Item

                Case "TProtNFe"
                    bSincrono = True
                    Dim objprotNFe As TProtNFe, bAutorizou As Boolean
                    objprotNFe = objRetEnviNFE.Item
                    gobjApp.GravarLog("Retorno do envio da nfce - " & objRetEnviNFE.xMotivo & IIf(objRetEnviNFE.cStat <> "0", " - cStat = " & objRetEnviNFE.cStat, ""))
                    lErro = protNFe_Processa(objprotNFe, bAutorizou)
                    If lErro <> SUCESSO Then Throw New System.Exception("")

                Case Else
                    Throw New System.Exception("Tipo " & sAux & " não tratado.")

            End Select

            If Not bSincrono Then

                If infRec Is Nothing Then
                    snRec = ""
                Else
                    snRec = infRec.nRec
                End If

                If gobjApp.iNFeAmbiente = NFE_AMBIENTE_HOMOLOGACAO Then
                    objconsReciNFe.tpAmb = TAmb.Item2
                Else
                    objconsReciNFe.tpAmb = TAmb.Item1
                End If
                objconsReciNFe.versao = NFE_VERSAO_XML
                objconsReciNFe.nRec = snRec

                Dim mySerializerx As New XmlSerializer(GetType(TConsReciNFe))

                XMLStream1 = New MemoryStream(10000)
                mySerializerx.Serialize(XMLStream1, objconsReciNFe)

                Dim xm1 As Byte()
                xm1 = XMLStream1.ToArray

                XMLString1 = System.Text.Encoding.UTF8.GetString(xm1)

                XMLString1 = Mid(XMLString1, 1, 19) & " encoding=""UTF-8"" " & Mid(XMLString1, 20)

                iPos = InStr(XMLString1, "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema""")

                If iPos <> 0 Then

                    XMLString1 = Mid(XMLString1, 1, iPos - 1) & Mid(XMLString1, iPos + 99)

                End If

                Call Salva_Arquivo(DocDados1, XMLString1)

                'NFecabec_RetEnvioLote.cUF = gobjApp.iUFCodIBGE
                'NFecabec_RetEnvioLote.versaoDados = NFE_VERSAO_XML

                'NFeRetAutorizacao.nfeCabecMsgValue = NFecabec_RetEnvioLote

                gobjApp.GravarLog("Iniciando a consulta do status do lote - Aguarde")

                '                System.Windows.Forms.Application.DoEvents()

            End If

            Lote_Consulta1 = SUCESSO

        Catch ex As Exception

            Lote_Consulta1 = 1

            Call gobjApp.GravarLog(ex.Message)

        End Try

    End Function

    Public Function Envia_NFe(ByVal objVenda As GlobaisLoja.ClassVenda) As Long

        Dim XMLString2 As String
        Dim objRetEnviNFE As TRetEnviNFe = New TRetEnviNFe
        Dim lErro As Long
        Dim XMLStringNFes As String

        Try

            sModelo = "NFe"

            lErro = Inicia_NF(objVenda, sModelo)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            sSerie = objVenda.objNFeInfo.sNFeSerie
            lNumNotaFiscal = objVenda.objNFeInfo.lNFeProximoNum
            lLote = objVenda.objNFeInfo.lNFeProximoLote

            Dim a5 = New TNFe
            XMLStringNFes = ""
            lErro = Monta_NFiscal_Xml(XMLStringNFes, a5)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            XMLString2 = ""
            lErro = Lote_Prepara(XMLStringNFes, XMLString2)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            lErro = gobjApp.XML_Valida(sArquivoLoteEnvio, "enviNFe_v4.00.xsd")
            If lErro <> SUCESSO Then Throw New System.Exception("")

            '****************  salva o arquivo 
            Dim DocDados1 = New XmlDocument
            Call Salva_Arquivo(DocDados1, XMLString2)

            lErro = gobjApp.Verifica_Status_Servico()
            If lErro <> SUCESSO Then Throw New System.Exception("")

            gobjApp.GravarLog("Iniciando o envio do lote")

            '            System.Windows.Forms.Application.DoEvents()

            lErro = Lote_Envia(DocDados1, objRetEnviNFE)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            lErro = Lote_Consulta(objRetEnviNFE)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            gobjApp.GravarLog("Encerrado o processamento do lote " & CStr(lLote))

            '            System.Windows.Forms.Application.DoEvents()

            Envia_NFe = SUCESSO

        Catch ex As Exception

            Envia_NFe = 1

            Dim sMsg2 As String

            If ex.InnerException Is Nothing Then
                sMsg2 = ""
            Else
                sMsg2 = " - " & ex.InnerException.Message
            End If

            gobjApp.GravarLog("ERRO - " & ex.Message & sMsg2 & IIf(lNumNotaFiscal <> 0, "Serie = " & sSerie & " Nota Fiscal = " & lNumNotaFiscal, ""))
            gobjApp.GravarLog("ERRO - " & gobjApp.sErro & " - " & gobjApp.sMsg1 & IIf(lNumNotaFiscal <> 0, " Serie = " & sSerie & " Nota Fiscal = " & lNumNotaFiscal, ""))
            gobjApp.GravarLog("ERRO - o envio do lote " & CStr(lLote) & " foi encerrado por erro.")

            gobjApp.Form1.Show()

            MsgBox("Verifique a mensagem de erro.")

        Finally

            gobjVenda = Nothing
            gobjApp.Terminar()
            gobjApp = Nothing

        End Try

    End Function

    Private Function Inicia_NF(ByVal objVenda As GlobaisLoja.ClassVenda, ByVal sModelo As String) As Long

        Try

            Dim lErro As Long

            sMetodo = "Envio"
            gobjVenda = objVenda

            bComISSQN = False
            bComICMS = False

            gobjApp = New ClassGlobalApp(objVenda, sModelo)

            gobjApp.GravarLog("Inicia_NF")

            lErro = gobjApp.Obtem_Certificado()
            If lErro <> SUCESSO Then Throw New System.Exception("")

            scDest = ""
            sdhEmi = ""
            svNF = ""
            svICMS = ""

            Inicia_NF = SUCESSO

        Catch ex As Exception

            Inicia_NF = 1

            Call gobjApp.GravarLog(ex.Message)

        End Try

    End Function

    Public Function Envia_NFCE_Offline(ByVal sArqXml As String, ByVal objVenda As GlobaisLoja.ClassVenda) As Long
        'enviar nfce que foi gerada durante contingencia offline

        Try

            Dim lErro As Long
            Dim objRetEnviNFE As TRetEnviNFe = New TRetEnviNFe
            Dim DocDados As XmlDocument = New XmlDocument

            DocDados.Load(sArqXml)

            sModelo = "NFCe"
            sMetodo = "EnvioContingenciaOffline"

            gobjVenda = objVenda

            gobjApp = New ClassGlobalApp(objVenda, sModelo)

            lErro = gobjApp.Obtem_Certificado()
            If lErro <> SUCESSO Then Throw New System.Exception("")

            'ajusta xml invalido
            Dim mySerializerImportaXML As New XmlSerializer(GetType(TEnviNFe))
            Dim objEnviNFe As New TEnviNFe
            Dim XMLString1 As String
            Dim xm As Byte()
            Dim XMLStreamDados = New MemoryStream(10000)
            Dim XMLStreamDados1 = New MemoryStream(10000)
            Dim xDados1 As Byte()
            Dim objNFe As TNFe

            Dim writer1 As New XmlTextWriter(XMLStreamDados1, Nothing)

            writer1.Formatting = Formatting.None
            DocDados.WriteTo(writer1)
            writer1.Flush()

            xm = XMLStreamDados1.ToArray

            XMLString1 = Replace(Replace(System.Text.Encoding.UTF8.GetString(xm), "BRASIL", "Brasil"), "<?xml version=""1.0"" encoding=""UTF-8""?>", "")

            xDados1 = System.Text.Encoding.UTF8.GetBytes(XMLString1)

            XMLStreamDados.Write(xDados1, 0, xDados1.Length)

            XMLStreamDados.Position = 0

            objEnviNFe = mySerializerImportaXML.Deserialize(XMLStreamDados)
            objNFe = objEnviNFe.NFe(0)
            gobjApp.GravarLog("Envia_NFCE_Offline 1", True, True, True, True, True)
            If (objNFe.infNFe.ide.cNF = Format(CLng(objNFe.infNFe.ide.nNF), "00000000")) Or Len(Trim(objNFe.infNFe.ide.xJust)) < 15 Or (UTCParaDate(objNFe.infNFe.ide.dhEmi) < #12/1/2015# And Not (objNFe.infNFeSupl Is Nothing) And gobjApp.iNFeAmbiente <> NFE_AMBIENTE_HOMOLOGACAO) Or InStr(XMLString1, "<NCM>48232090</NCM>") <> 0 Or InStr(XMLString1, "<NCM>22029000</NCM>") <> 0 Or InStr(XMLString1, "<NCM>20098000</NCM>") <> 0 Then
                gobjApp.GravarLog("Envia_NFCE_Offline 2", True, True, True, True, True)
                Dim XMLStream As New MemoryStream
                Dim XMLString As String
                Dim iPos As Integer
                Dim iPos2 As Integer
                Dim iPos3 As String

                '                MsgBox("2")

                'SP nao tem contingencia offline
                If Mid(objNFe.infNFe.Id, 4, 2) = "35" Then

                    objNFe.infNFe.ide.tpEmis = TNFeInfNFeIdeTpEmis.Item1
                    objNFe.infNFe.ide.dhCont = Nothing
                    objNFe.infNFe.ide.xJust = Nothing

                Else

                    objNFe.infNFe.ide.xJust = "falta de acesso a internet"

                End If

                'tirar NFe do inicio e o digito verificador do final
                If objNFe.infNFe.ide.cNF = Format(CLng(objNFe.infNFe.ide.nNF), "00000000") Then
                    gobjApp.GravarLog("Envia_NFCE_Offline 3", True, True, True, True, True)
                    Dim sArquivoOld = gobjApp.sDirXml & Mid(objNFe.infNFe.Id, 4) & "-pre.xml"

                    objNFe.infNFe.Id = Mid(objNFe.infNFe.Id, 4, 43)
                    objNFe.infNFe.ide.cNF = Format(lNumNotaFiscal + 11111, "00000000")
                    objNFe.infNFe.Id = Strings.Left(objNFe.infNFe.Id, (Len(objNFe.infNFe.Id) - 8)) & objNFe.infNFe.ide.cNF
                    Dim iDigito As Integer

                    CalculaDV_Modulo11(objNFe.infNFe.Id, iDigito)

                    objNFe.infNFe.Id = "NFe" & objNFe.infNFe.Id & iDigito
                    objNFe.infNFe.ide.cDV = iDigito

                    Dim sArquivoNew = gobjApp.sDirXml & Mid(objNFe.infNFe.Id, 4) & "-pre.xml"
                    gobjApp.GravarLog("Envia_NFCE_Offline 4 " & sArquivoOld & " " & sArquivoNew & " " & objNFe.infNFe.Id, True, True, True, True, True)
                    If System.IO.File.Exists(sArquivoOld) And Not System.IO.File.Exists(sArquivoNew) Then
                        My.Computer.FileSystem.RenameFile(sArquivoOld, Mid(objNFe.infNFe.Id, 4) & "-pre.xml")
                    End If

                End If

                objNFe.Signature = Nothing

                objNFe.infNFeSupl = New TNFeInfNFeSupl
                objNFe.infNFeSupl.qrCode = QRCODE_PROVISORIO

                Dim AD As AssinaturaDigital = New AssinaturaDigital

                Dim mySerializer As New XmlSerializer(GetType(TNFe))

                XMLStream = New MemoryStream(10000)

                mySerializer.Serialize(XMLStream, objNFe)

                Dim xm2 As Byte()
                xm2 = XMLStream.ToArray

                XMLString = System.Text.Encoding.UTF8.GetString(xm2)

                XMLString = Replace(XMLString, "<ICMS100>", "<ICMS40>")
                XMLString = Replace(XMLString, "</ICMS100>", "</ICMS40>")

                '***********************************
                'acerto de ncms invalidos na mercado romano
                XMLString = Replace(XMLString, "<NCM>48232090</NCM>", "<NCM>48232099</NCM>")
                XMLString = Replace(XMLString, "<NCM>22029000</NCM>", "<NCM>22021000</NCM>")
                XMLString = Replace(XMLString, "<NCM>20098000</NCM>", "<NCM>22021000</NCM>")

                iPos = InStr(XMLString, "xmlns:xsi")
                iPos2 = InStr(XMLString, """>")

                XMLString = Mid(XMLString, 1, iPos - 1) & Mid(XMLString, iPos2 + 1)

                'retirado em 31/03/2010 pois estava dando erro no xml
                iPos3 = InStr(XMLString, "<NFe >")

                XMLString = Mid(XMLString, 1, iPos3 + 4) & "xmlns = ""http://www.portalfiscal.inf.br/nfe""" & Mid(XMLString, iPos3 + 5)

                '****************************************


                iPos = InStr(XMLString, "<infNFe")

                If iPos <> 0 Then

                    Dim iPosAux As Integer

                    iPosAux = InStr(Mid(XMLString, iPos), "xmlns=""http://www.portalfiscal.inf.br/nfe""")

                    If iPosAux <> 0 Then

                        XMLString = Mid(XMLString, 1, iPos + iPosAux - 2) & Mid(XMLString, iPos + iPosAux + 41)

                    End If

                End If


                If gobjApp.iDebug = 1 Then MsgBox("37")
                gobjApp.sErro = "37"
                gobjApp.sMsg1 = "vai assinar a nota"


                lErro = AD.Assinar(XMLString, "infNFe", gobjApp.cert, gobjApp.iDebug)
                If lErro <> SUCESSO Then Throw New System.Exception("")

                Dim xString As String
                xString = AD.XMLStringAssinado

                'MsgBox("3 " & xString)

                If InStr(xString, "<qrCode>") <> 0 Then

                    Dim sNFCeQRCode As String, sNFCeQRCodeAnterior As String

                    Dim iPosAux1 As Integer, iPosAux2 As Integer
                    iPosAux1 = InStr(xString, "<qrCode>") + Len("<qrCode>")
                    iPosAux2 = InStr(xString, "</qrCode>")
                    sNFCeQRCodeAnterior = Mid(xString, iPosAux1, iPosAux2 - iPosAux1)

                    Dim sDigVal As String = ""
                    iPosAux1 = InStr(xString, "<DigestValue>") + Len("<DigestValue>")
                    iPosAux2 = InStr(xString, "</DigestValue>")
                    sDigVal = Mid(xString, iPosAux1, iPosAux2 - iPosAux1)

                    Dim sCPFCGC As String = ""

                    If Not (objNFe.infNFe.dest Is Nothing) Then

                        If objNFe.infNFe.dest.ItemElementName = ItemChoiceType5.CNPJ Or objNFe.infNFe.dest.ItemElementName = ItemChoiceType5.CPF Then
                            sCPFCGC = objNFe.infNFe.dest.Item
                        End If

                    End If

                    'TODO: Trocada chamada do QRCode Offline 2
                    gobjApp.GravarLog("Envia_NFCE_Offline NFCE_Gera_QRCode2_Offline",, False)

                    'sNFCeQRCode = NFCE_Gera_QRCode(Mid(objNFe.infNFe.Id, 4), "100", GetXmlAttrNameFromEnumValue(Of TAmb)(objNFe.infNFe.ide.tpAmb), sCPFCGC, objNFe.infNFe.ide.dhEmi, objNFe.infNFe.total.ICMSTot.vNF, objNFe.infNFe.total.ICMSTot.vICMS, sDigVal, gobjVenda.objNFeInfo.sidNFCECSC, gobjVenda.objNFeInfo.sNFCECSC)
                    sNFCeQRCode = NFCE_Gera_QRCode2_Offline(gobjApp, Mid(objNFe.infNFe.Id, 4), GetXmlAttrNameFromEnumValue(Of TAmb)(objNFe.infNFe.ide.tpAmb), objNFe.infNFe.ide.dhEmi, objNFe.infNFe.total.ICMSTot.vNF, sDigVal, gobjVenda.objNFeInfo.sidNFCECSC, gobjVenda.objNFeInfo.sNFCECSC)
                    sNFCeQRCode = "<![CDATA[" & sNFCeQRCode & "]]>"

                    xString = Replace(xString, sNFCeQRCodeAnterior, sNFCeQRCode)
                    xString = Replace(xString, "</qrCode>", "</qrCode><urlChave>www.fazenda.rj.gov.br/nfce/consulta</urlChave>")

                End If

                Dim XMLStringNFes As String
                Dim XMLString2 As String
                Dim XMLStream1 As New MemoryStream

                XMLStringNFes = Mid(xString, 22) & " "

                Dim mySerializerw As New XmlSerializer(GetType(TEnviNFe))

                XMLStream1 = New MemoryStream(10000)

                objEnviNFe.NFe = Nothing
                mySerializerw.Serialize(XMLStream1, objEnviNFe)

                Dim xmw As Byte()
                xmw = XMLStream1.ToArray

                XMLString1 = System.Text.Encoding.UTF8.GetString(xmw)

                XMLString2 = Mid(XMLString1, 1, Len(XMLString1) - 10) & XMLStringNFes & Mid(XMLString1, Len(XMLString1) - 10)

                XMLString2 = Mid(XMLString2, 1, 19) & " encoding=""UTF-8"" " & Mid(XMLString2, 20)

                iPos = InStr(XMLString2, "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema""")

                If iPos <> 0 Then

                    XMLString2 = Mid(XMLString2, 1, iPos - 1) & Mid(XMLString2, iPos + 99)

                End If

                'MsgBox("4 " & XMLString2)
                gobjApp.GravarLog(XMLString2)

                '************* valida dados antes do envio **********************
                Dim xDados As Byte()

                xDados = System.Text.Encoding.UTF8.GetBytes(XMLString2)

                XMLStreamDados = New MemoryStream(10000)

                XMLStreamDados.Write(xDados, 0, xDados.Length)

                XMLStreamDados.Position = 0
                DocDados.Load(XMLStreamDados)


            End If

            lErro = Lote_Envia(DocDados, objRetEnviNFE)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            lErro = Lote_Consulta(objRetEnviNFE)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            If objVenda.objCupomFiscal.sNFenProt = "" Then Throw New System.Exception("")

            Envia_NFCE_Offline = SUCESSO

        Catch ex As Exception

            Envia_NFCE_Offline = 1

            Dim sMsg2 As String

            If ex.InnerException Is Nothing Then
                sMsg2 = ""
            Else
                sMsg2 = " - " & ex.InnerException.Message
            End If

            gobjApp.GravarLog("ERRO - " & ex.Message & sMsg2)
            gobjApp.GravarLog("ERRO - " & gobjApp.sErro & " - " & gobjApp.sMsg1 & "arquivo: " & sArqXml)
            gobjApp.Form1.Show()

            MsgBox("Verifique a mensagem de erro.")

        Finally

            gobjVenda = Nothing
            gobjApp.Terminar()
            gobjApp = Nothing

        End Try

    End Function

    Public Function Envia_NFCE(ByVal objVenda As GlobaisLoja.ClassVenda) As Long

        Dim XMLString2 As String
        Dim objRetEnviNFE As TRetEnviNFe = New TRetEnviNFe
        Dim lErro As Long
        Dim XMLStringNFes As String

        Try

            sModelo = "NFCe"
            sSerie = objVenda.objNFeInfo.sNFCeSerie
            lNumNotaFiscal = objVenda.objNFeInfo.lNFCeProximoNum
            lLote = objVenda.objNFeInfo.lNFCeProximoLote

            lErro = Inicia_NF(objVenda, sModelo)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            Dim a5 = New TNFe
            XMLStringNFes = ""
            lErro = Monta_NFiscal_Xml(XMLStringNFes, a5)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            XMLString2 = ""
            lErro = Lote_Prepara(XMLStringNFes, XMLString2)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            lErro = gobjApp.XML_Valida(sArquivoLoteEnvio, "enviNFe_v4.00.xsd")
            If lErro <> SUCESSO Then
                objVenda.objCupomFiscal.bEditavel = True
                Throw New System.Exception("")
            End If

            '****************  salva o arquivo 
            Dim DocDados1 = New XmlDocument
            Call Salva_Arquivo(DocDados1, XMLString2)

            If gobjVenda.objNFeInfo.iEmContingencia = 0 Or (sModelo = "NFCe" And gobjVenda.objNFeInfo.sEmitUF = "SP") Then

                gobjApp.GravarLog("Iniciando o envio do lote")

                '                System.Windows.Forms.Application.DoEvents()

                lErro = Lote_Envia(DocDados1, objRetEnviNFE)
                If lErro <> SUCESSO Then Throw New System.Exception("")

                lErro = Lote_Consulta(objRetEnviNFE)
                If lErro <> SUCESSO Then Throw New System.Exception("")

                If gobjVenda.objCupomFiscal.sNFenProt = "" Then Throw New System.Exception("")

            Else

                'contingencia offline
                gobjVenda.objCupomFiscal.sNFeChaveAcesso = Right(a5.infNFe.Id, 44)
                gobjVenda.objCupomFiscal.sNFeArqXml = sArquivoLoteEnvio
                gobjVenda.objCupomFiscal.lNumero = lNumNotaFiscal
                gobjVenda.objCupomFiscal.lCCF = lNumNotaFiscal
                gobjVenda.objCupomFiscal.sNumSerieECF = sSerie

                gobjVenda.objCupomFiscal.sNFecStat = ""
                gobjVenda.objCupomFiscal.sNFenProt = ""
                gobjVenda.objCupomFiscal.dtNFeData = DATA_NULA
                gobjVenda.objCupomFiscal.dNFeHora = UTCParaHora(0)
                gobjVenda.objCupomFiscal.sNFEversao = NFE_VERSAO_XML
                gobjVenda.objCupomFiscal.iNFetpAmb = CInt(GetXmlAttrNameFromEnumValue(Of TAmb)(a5.infNFe.ide.tpAmb))

                Dim sDigVal As String

                Dim ns As New XmlNamespaceManager(DocDados1.NameTable)
                ns.AddNamespace("nfe", "http://www.portalfiscal.inf.br/nfe")
                ns.AddNamespace("sig", "http://www.w3.org/2000/09/xmldsig#")
                Dim xpathNav As XPathNavigator = DocDados1.CreateNavigator()
                Dim node As XPathNavigator = xpathNav.SelectSingleNode("//nfe:NFe/sig:Signature/sig:SignedInfo/sig:Reference/sig:DigestValue", ns)
                sDigVal = node.InnerXml

                'TODO: Trocada chamada do QRCode Offline 1
                gobjApp.GravarLog("Envia_NFCE NFCE_Gera_QRCode2_Offline",, False)

                'gobjVenda.objCupomFiscal.sNFCeQRCode = NFCE_Gera_QRCode(gobjVenda.objCupomFiscal.sNFeChaveAcesso, "100", GetXmlAttrNameFromEnumValue(Of TAmb)(a5.infNFe.ide.tpAmb), scDest, sdhEmi, svNF, svICMS, sDigVal, gobjVenda.objNFeInfo.sidNFCECSC, gobjVenda.objNFeInfo.sNFCECSC)
                gobjVenda.objCupomFiscal.sNFCeQRCode = NFCE_Gera_QRCode2_Offline(gobjApp, gobjVenda.objCupomFiscal.sNFeChaveAcesso, GetXmlAttrNameFromEnumValue(Of TAmb)(a5.infNFe.ide.tpAmb), gobjVenda.objCupomFiscal.dtDataEmissao, svNF, sDigVal, gobjVenda.objNFeInfo.sidNFCECSC, gobjVenda.objNFeInfo.sNFCECSC)

            End If

            If gobjVenda.objNFeInfo.iNFCeEnviarEmail <> 0 And gobjVenda.objCupomFiscal.sNFCeQRCode <> "" Then
                Dim sEMailTexto As String = ""
                Call Email_Monta_Texto_Venda(sEMailTexto)
                Call Email_Enviar(gobjVenda.objNFeInfo.sSMTP, gobjVenda.objNFeInfo.sSMTPUsu, gobjVenda.objNFeInfo.sSMTPSenha, CStr(gobjVenda.objNFeInfo.lSMTPPorta), "Comprovante de Emissão de NFCe", sEMailTexto, IIf(gobjVenda.objCupomFiscal.sEndEntEmail <> "", gobjVenda.objCupomFiscal.sEndEntEmail, gobjVenda.objNFeInfo.sSMTPUsu))
            End If

            gobjApp.GravarLog("Encerrado o processamento do lote " & CStr(lLote))

            'System.Windows.Forms.Application.DoEvents()

            Envia_NFCE = SUCESSO

        Catch ex As Exception

            Envia_NFCE = 1

            Dim sMsg2 As String

            If ex.InnerException Is Nothing Then
                sMsg2 = ""
            Else
                sMsg2 = " - " & ex.InnerException.Message
            End If

            gobjApp.GravarLog("ERRO - " & ex.Message & sMsg2 & IIf(lNumNotaFiscal <> 0, "Serie = " & sSerie & " Nota Fiscal = " & lNumNotaFiscal, ""))
            gobjApp.GravarLog("ERRO - " & gobjApp.sErro & " - " & gobjApp.sMsg1 & IIf(lNumNotaFiscal <> 0, " Serie = " & sSerie & " Nota Fiscal = " & lNumNotaFiscal, ""))
            gobjApp.GravarLog("ERRO - o envio do lote " & CStr(lLote) & " foi encerrado por erro.")

            gobjApp.Form1.Show()

            MsgBox("Verifique a mensagem de erro.")

        Finally

            gobjVenda = Nothing
            gobjApp.Terminar()
            gobjApp = Nothing

        End Try

    End Function

    Private Function Monta_NFiscal_Xml1(ByVal a5 As TNFe) As Long

        Try

            Dim infNFe As TNFeInfNFe = New TNFeInfNFe
            a5.infNFe = infNFe

            a5.infNFe.versao = NFE_VERSAO_XML

            Dim infNFeIde As TNFeInfNFeIde = New TNFeInfNFeIde
            a5.infNFe.ide = infNFeIde

            If gobjApp.iDebug = 1 Then MsgBox("5")

            a5.infNFe.ide.cUF = GetCode(Of TCodUfIBGE)(CStr(gobjApp.iUFCodIBGE))

            If gobjApp.iDebug = 1 Then MsgBox("6")

            gobjApp.sErro = "6"
            gobjApp.sMsg1 = "vai acessar Cidades, Paises"

            If gobjApp.iNFeAmbiente = NFE_AMBIENTE_HOMOLOGACAO Then
                a5.infNFe.ide.tpAmb = TAmb.Item2
            Else
                a5.infNFe.ide.tpAmb = TAmb.Item1
            End If

            'Randomize()
            'a5.infNFe.ide.cNF = Format(Rnd() * 100000000, "00000000")
            a5.infNFe.ide.cNF = Format(lNumNotaFiscal + 11111, "00000000") 'para nao ficar alterando a chave a cada envio

            a5.infNFe.ide.natOp = DesacentuaTexto("Venda") '???? sDescrNF

            'a5.infNFe.ide.indPag = TNFeInfNFeIdeIndPag.Item2


            a5.infNFe.ide.mod = IIf(sModelo = "NFe", TMod.Item55, TMod.Item65)
            a5.infNFe.ide.serie = sSerie
            a5.infNFe.ide.nNF = lNumNotaFiscal

            If gobjApp.iDebug = 1 Then MsgBox("8")

            gobjApp.sErro = "8"
            gobjApp.sMsg1 = "vai completar a estrutura a5.infNFe.ide"

            If gobjVenda.objNFeInfo.iEmContingencia = 0 Or (sModelo = "NFCe" And gobjVenda.objNFeInfo.sEmitUF = "SP") Then
                a5.infNFe.ide.tpEmis = TNFeInfNFeIdeTpEmis.Item1
            Else
                If sModelo = "NFCe" Then
                    a5.infNFe.ide.tpEmis = TNFeInfNFeIdeTpEmis.Item9
                    a5.infNFe.ide.dhCont = DataHoraParaUTC(gobjVenda.objNFeInfo.dtContingenciaDataEntrada, gobjVenda.objNFeInfo.dContingenciaHoraEntrada)
                    a5.infNFe.ide.xJust = gobjVenda.objNFeInfo.sContigenciaxJust
                Else
                    a5.infNFe.ide.tpEmis = TNFeInfNFeIdeTpEmis.Item3
                End If
            End If

            If sModelo = "NFe" Then
                a5.infNFe.ide.tpImp = TNFeInfNFeIdeTpImp.Item1
            Else
                a5.infNFe.ide.tpImp = TNFeInfNFeIdeTpImp.Item4
            End If

            a5.infNFe.ide.dhEmi = DataHoraParaUTC(gobjVenda.objCupomFiscal.dtDataEmissao, gobjVenda.objCupomFiscal.dHoraEmissao)
            sdhEmi = a5.infNFe.ide.dhEmi

            If gobjApp.iDebug = 1 Then MsgBox("9")

            gobjApp.sErro = "9"
            gobjApp.sMsg1 = "preenche a data/hora de entrada/saida"

            If sModelo <> "NFCe" Then a5.infNFe.ide.dhSaiEnt = a5.infNFe.ide.dhEmi
            a5.infNFe.ide.tpNF = TNFeInfNFeIdeTpNF.Item1

            a5.infNFe.ide.finNFe = TFinNFe.Item1

            a5.infNFe.ide.indFinal = TNFeInfNFeIdeIndFinal.Item1

            a5.infNFe.ide.procEmi = TProcEmi.Item0
            a5.infNFe.ide.verProc = "Corporator"

            Monta_NFiscal_Xml1 = SUCESSO

        Catch ex As Exception

            Monta_NFiscal_Xml1 = 1

            Call gobjApp.GravarLog(ex.Message)

        End Try

    End Function

    Private Function Monta_NFiscal_Xml2(ByVal a5 As TNFe) As Long

        Try

            Dim sIE As String, sIM As String, sCNAE As String
            Dim infNFeEmit As TNFeInfNFeEmit = New TNFeInfNFeEmit
            a5.infNFe.emit = infNFeEmit

            a5.infNFe.emit.xNome = DesacentuaTexto(Trim(gobjVenda.objNFeInfo.sEmitRazaoSocial))
            a5.infNFe.emit.xFant = DesacentuaTexto(Trim(gobjVenda.objNFeInfo.sEmitNomeReduzido))

            Dim objEndereco As GlobaisAdm.ClassEndereco

            objEndereco = gobjVenda.objNFeInfo.objEndereco

            a5.infNFe.emit.ItemElementName = ItemChoiceType.CNPJ
            a5.infNFe.emit.Item = gobjVenda.objNFeInfo.sEmitCNPJ

            If gobjApp.iDebug = 1 Then MsgBox("10")

            gobjApp.sErro = "10"
            gobjApp.sMsg1 = "vai acessar as tabelas Empresas e AdmConfig"

            Dim enderEmit As TEnderEmi = New TEnderEmi
            a5.infNFe.emit.enderEmit = enderEmit


            If Len(objEndereco.sLogradouro) > 0 Then
                a5.infNFe.emit.enderEmit.xLgr = DesacentuaTexto(Left(IIf(Len(objEndereco.sTipoLogradouro) > 0, objEndereco.sTipoLogradouro & " ", "") & objEndereco.sLogradouro, 60))
                a5.infNFe.emit.enderEmit.nro = objEndereco.lNumero
                If Len(DesacentuaTexto(objEndereco.sComplemento)) > 0 Then a5.infNFe.emit.enderEmit.xCpl = DesacentuaTexto(objEndereco.sComplemento)
            Else
                a5.infNFe.emit.enderEmit.xLgr = DesacentuaTexto(objEndereco.sEndereco)
                a5.infNFe.emit.enderEmit.nro = "0"
            End If
            If Len(objEndereco.sBairro) = 0 Then
                a5.infNFe.emit.enderEmit.xBairro = "a"
            Else
                a5.infNFe.emit.enderEmit.xBairro = Trim(DesacentuaTexto(objEndereco.sBairro))
            End If

            a5.infNFe.emit.enderEmit.cMun = gobjVenda.objNFeInfo.cMun
            a5.infNFe.emit.enderEmit.xMun = DesacentuaTexto(objEndereco.sCidade)
            a5.infNFe.emit.enderEmit.UF = GetCode(Of TUf)(gobjApp.sSiglaEstado)

            If Len(objEndereco.sCEP) > 0 Then
                a5.infNFe.emit.enderEmit.CEP = objEndereco.sCEP
            End If
            If Len(objEndereco.sTelNumero1) > 0 Then
                Call Formata_String_Numero(IIf(Len(CStr(objEndereco.iTelDDD1)) > 0, CStr(objEndereco.iTelDDD1), "") + objEndereco.sTelNumero1, a5.infNFe.emit.enderEmit.fone)
            ElseIf Len(objEndereco.sTelefone1) > 0 Then
                Call Formata_String_Numero(objEndereco.sTelefone1, a5.infNFe.emit.enderEmit.fone)
                'Else
                '    a5.infNFe.emit.enderEmit.fone = "99999999"
            End If

            a5.infNFe.ide.cMunFG = a5.infNFe.emit.enderEmit.cMun


            a5.infNFe.emit.enderEmit.cPais = TEnderEmiCPais.Item1058
            a5.infNFe.emit.enderEmit.cPaisSpecified = True
            a5.infNFe.emit.enderEmit.xPais = TEnderEmiXPais.Brasil
            sIE = ""
            Call Formata_String_Numero(gobjVenda.objNFeInfo.sEmitIE, sIE)
            a5.infNFe.emit.IE = sIE
            sIM = ""
            Call Formata_String_Numero(gobjVenda.objNFeInfo.sEmitIM, sIM)
            If Len(sIM) > 0 Then
                a5.infNFe.emit.IM = sIM
                sCNAE = ""
                Call Formata_String_Numero(gobjVenda.objNFeInfo.sEmitCNAE, sCNAE)
                a5.infNFe.emit.CNAE = sCNAE
            End If

            Monta_NFiscal_Xml2 = SUCESSO

        Catch ex As Exception

            Monta_NFiscal_Xml2 = 1

            Call gobjApp.GravarLog(ex.Message)

        End Try

    End Function

    Private Function Monta_NFiscal_Xml3(ByVal a5 As TNFe) As Long
        Try

            Dim sIE As String

            If gobjApp.iDebug = 1 Then MsgBox("11")
            gobjApp.sErro = "11"
            gobjApp.sMsg1 = "vai acessar os dados do destinatario"

            If Len(gobjVenda.objCupomFiscal.sCPFCGC) > 0 Then

                Dim infNFeDest As TNFeInfNFeDest = New TNFeInfNFeDest
                a5.infNFe.dest = infNFeDest

                If gobjApp.iDebug = 1 Then MsgBox("12")
                gobjApp.sErro = "12"
                gobjApp.sMsg1 = "vai acessar a tabela Clientes"

                a5.infNFe.dest.xNome = DesacentuaTexto(Trim(gobjVenda.objCupomFiscal.sNomeCliente))
                If Len(Trim(a5.infNFe.dest.xNome)) = 0 Then a5.infNFe.dest.xNome = "NOME NAO INFORMADO"

                If Len(gobjVenda.objCupomFiscal.sCPFCGC) <= 11 Then
                    a5.infNFe.dest.ItemElementName = ItemChoiceType5.CPF
                Else
                    a5.infNFe.dest.ItemElementName = ItemChoiceType5.CNPJ
                End If

                a5.infNFe.dest.Item = gobjVenda.objCupomFiscal.sCPFCGC
                scDest = gobjVenda.objCupomFiscal.sCPFCGC

                If sModelo = "NFCe" Then
                    a5.infNFe.dest.indIEDest = TNFeInfNFeDestIndIEDest.Item9
                Else
                    sIE = ""
                    Call Formata_String_Numero(gobjVenda.objCupomFiscal.sInscricaoEstadual, sIE)

                    '???? nao trata qdo tem inscricao mas nao é contribuinte
                    If gobjVenda.objCupomFiscal.iIEIsento = 1 Or Len(Trim(sIE)) = 0 Then
                        a5.infNFe.dest.indIEDest = TNFeInfNFeDestIndIEDest.Item2
                    Else
                        a5.infNFe.dest.IE = sIE
                    End If
                End If

                If gobjApp.iDebug = 1 Then MsgBox("15")
                gobjApp.sErro = "15"
                gobjApp.sMsg1 = "vai acessar a tabela Endereco do Destinatario, Estado, Cidades e Paises"

                Dim objEndDest As New GlobaisAdm.ClassEndereco

                objEndDest = gobjVenda.objCupomFiscal.objEndDest

                If Len(objEndDest.sLogradouro) > 0 Then

                    Dim enderDest As TEndereco = New TEndereco
                    a5.infNFe.dest.enderDest = enderDest

                    a5.infNFe.dest.enderDest.cPais = 1058
                    'a5.infNFe.dest.enderDest.cPaisSpecified = True

                    a5.infNFe.dest.enderDest.xLgr = DesacentuaTexto(Left(IIf(Len(objEndDest.sTipoLogradouro) > 0, objEndDest.sTipoLogradouro & " ", "") & objEndDest.sLogradouro, 60))
                    a5.infNFe.dest.enderDest.nro = objEndDest.lNumero
                    If Len(DesacentuaTexto(objEndDest.sComplemento)) > 0 Then a5.infNFe.dest.enderDest.xCpl = DesacentuaTexto(objEndDest.sComplemento)

                    If Len(objEndDest.sTelNumero1) > 0 Then
                        Call Formata_String_Numero(IIf(Len(CStr(objEndDest.iTelDDD1)) > 0, CStr(objEndDest.iTelDDD1), "") + objEndDest.sTelNumero1, a5.infNFe.dest.enderDest.fone)
                    ElseIf Len(objEndDest.sTelefone1) > 0 Then
                        Call Formata_String_Numero(objEndDest.sTelefone1, a5.infNFe.dest.enderDest.fone)
                    End If

                    If Len(objEndDest.sBairro) = 0 Then
                        a5.infNFe.dest.enderDest.xBairro = "a"
                    Else
                        a5.infNFe.dest.enderDest.xBairro = Trim(DesacentuaTexto(objEndDest.sBairro))
                    End If

                    'v2.00
                    If Len(objEndDest.sEmail) <> 0 Then
                        a5.infNFe.dest.email = Trim(DesacentuaTexto(objEndDest.sEmail))
                    End If

                    a5.infNFe.dest.enderDest.cMun = gobjVenda.objCupomFiscal.lEndEntIBGECidade
                    a5.infNFe.dest.enderDest.xMun = DesacentuaTexto(objEndDest.sCidade)
                    If Len(Trim(objEndDest.sSiglaEstado)) Then a5.infNFe.dest.enderDest.UF = GetCode(Of TUf)(objEndDest.sSiglaEstado)

                    If Len(objEndDest.sCEP) > 0 Then
                        a5.infNFe.dest.enderDest.CEP = objEndDest.sCEP
                    End If

                End If

                If gobjApp.iNFeAmbiente = NFE_AMBIENTE_HOMOLOGACAO Then
                    a5.infNFe.dest.xNome = "NF-E EMITIDA EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL"
                    '                    a5.infNFe.dest.ItemElementName = ItemChoiceType5.CNPJ
                    '                   a5.infNFe.dest.Item = "99999999000191"
                End If

            End If

            'precisa dos dados do transportador
            '        If Len(objEndDest.sLogradouro) > 0 And sModelo = "NFCe" Then
            ' a5.infNFe.ide.indPres = TNFeInfNFeIdeIndPres.Item4
            ' Else
            a5.infNFe.ide.indPres = TNFeInfNFeIdeIndPres.Item1
            'End If

            Monta_NFiscal_Xml3 = SUCESSO

        Catch ex As Exception

            Monta_NFiscal_Xml3 = 1

            Call gobjApp.GravarLog(ex.Message)

        End Try

    End Function

    Private Function Monta_NFiscal_Xml4h(ByVal a5 As TNFe, ByVal dTotalDescontoItem As Double, ByVal dvProdICMS As Double, ByVal dValorPIS As Double, ByVal dValorServPIS As Double, ByVal dValorCOFINS As Double, ByVal dValorServCOFINS As Double, ByVal dServNTribICMS As Double) As Long

        Try

            Dim objTributacaoDoc As GlobaisAdm.ClassTributacaoDoc

            If gobjApp.iDebug = 1 Then MsgBox("28")
            gobjApp.sErro = "28"
            gobjApp.sMsg1 = "vai iniciar o tratamento dos totais da nota"

            objTributacaoDoc = gobjVenda.objCupomFiscal.objTributacaoDoc

            '***********  total ****************************

            Dim infNFeTotal As TNFeInfNFeTotal = New TNFeInfNFeTotal
            a5.infNFe.total = infNFeTotal


            '***********  icms total ****************************

            Dim infNFeTotalICMSTot As TNFeInfNFeTotalICMSTot = New TNFeInfNFeTotalICMSTot
            a5.infNFe.total.ICMSTot = infNFeTotalICMSTot

            a5.infNFe.total.ICMSTot.vBC = Replace(Format(objTributacaoDoc.dICMSBase, "fixed"), ",", ".")
            a5.infNFe.total.ICMSTot.vICMS = Replace(Format((objTributacaoDoc.dICMSValor - objTributacaoDoc.dICMSVlrFCP), "fixed"), ",", ".")
            svICMS = a5.infNFe.total.ICMSTot.vICMS
            a5.infNFe.total.ICMSTot.vICMSDeson = Replace(Format(objTributacaoDoc.dICMSValorIsento, "fixed"), ",", ".")
            a5.infNFe.total.ICMSTot.vBCST = Replace(Format(objTributacaoDoc.dICMSSubstBase, "fixed"), ",", ".")
            a5.infNFe.total.ICMSTot.vST = Replace(Format((objTributacaoDoc.dICMSSubstValor - objTributacaoDoc.dICMSVlrFCPST), "fixed"), ",", ".")
            a5.infNFe.total.ICMSTot.vProd = Replace(Format(dvProdICMS + gobjVenda.objCupomFiscal.dValorDesconto, "fixed"), ",", ".")
            a5.infNFe.total.ICMSTot.vFrete = Replace(Format(objTributacaoDoc.dValorFrete, "fixed"), ",", ".")
            a5.infNFe.total.ICMSTot.vSeg = Replace(Format(objTributacaoDoc.dValorSeguro, "fixed"), ",", ".")
            a5.infNFe.total.ICMSTot.vDesc = Replace(Format(dTotalDescontoItem, "fixed"), ",", ".")
            a5.infNFe.total.ICMSTot.vIPI = Replace(Format(objTributacaoDoc.dIPIValor, "fixed"), ",", ".")
            a5.infNFe.total.ICMSTot.vPIS = Replace(Format(dValorPIS - dValorServPIS, "fixed"), ",", ".")
            a5.infNFe.total.ICMSTot.vCOFINS = Replace(Format(dValorCOFINS - dValorServCOFINS, "fixed"), ",", ".")
            a5.infNFe.total.ICMSTot.vOutro = Replace(Format(objTributacaoDoc.dValorOutrasDespesas, "fixed"), ",", ".")
            a5.infNFe.total.ICMSTot.vNF = Replace(Format(gobjVenda.objCupomFiscal.dValorTotal + gobjVenda.objCupomFiscal.dValorDesconto, "fixed"), ",", ".")

            infNFeTotal.ICMSTot.vFCP = Replace(Format(objTributacaoDoc.dICMSVlrFCP, "fixed"), ",", ".")
            infNFeTotal.ICMSTot.vFCPST = Replace(Format(objTributacaoDoc.dICMSVlrFCPST, "fixed"), ",", ".")
            infNFeTotal.ICMSTot.vFCPSTRet = Replace(Format(objTributacaoDoc.dICMSVlrFCPSTRet, "fixed"), ",", ".")
            infNFeTotal.ICMSTot.vIPIDevol = Replace(Format(objTributacaoDoc.dIPIVlrDevolvido, "fixed"), ",", ".")

            svNF = a5.infNFe.total.ICMSTot.vNF

            a5.infNFe.total.ICMSTot.vII = Replace(Format(0, "fixed"), ",", ".")

            If objTributacaoDoc.dTotTrib <> 0 Then
                a5.infNFe.total.ICMSTot.vTotTrib = Replace(Format(objTributacaoDoc.dTotTrib, "fixed"), ",", ".")
            End If

            ' ************ ISSQN total ***********************
            If bComISSQN Then

                Dim infNFeTotalISSQNtot As TNFeInfNFeTotalISSQNtot = New TNFeInfNFeTotalISSQNtot
                a5.infNFe.total.ISSQNtot = infNFeTotalISSQNtot

                If objTributacaoDoc.dISSBase > 0 Then
                    a5.infNFe.total.ISSQNtot.vBC = Replace(Format(objTributacaoDoc.dISSBase, "fixed"), ",", ".")
                End If

                If dServNTribICMS > 0 Then
                    a5.infNFe.total.ISSQNtot.vServ = Replace(Format(dServNTribICMS, "fixed"), ",", ".")
                End If

                If objTributacaoDoc.dISSValor > 0 Then
                    a5.infNFe.total.ISSQNtot.vISS = Replace(Format(objTributacaoDoc.dISSValor, "fixed"), ",", ".")
                End If

                If dValorServPIS > 0 Then
                    a5.infNFe.total.ISSQNtot.vPIS = Replace(Format(dValorServPIS, "fixed"), ",", ".")
                End If

                If dValorServCOFINS > 0 Then
                    a5.infNFe.total.ISSQNtot.vCOFINS = Replace(Format(dValorServCOFINS, "fixed"), ",", ".")
                End If

                If gobjApp.iDebug = 1 Then MsgBox("29")
                gobjApp.sErro = "29"
                gobjApp.sMsg1 = "vai iniciar o tratamento das retencoes de impostos da nota"

                '***********  retencao total ****************************

                Dim infNFeTotalRetTrib As TNFeInfNFeTotalRetTrib = New TNFeInfNFeTotalRetTrib
                a5.infNFe.total.retTrib = infNFeTotalRetTrib

                If objTributacaoDoc.dPISRetido > 0.0 Or objTributacaoDoc.dCOFINSRetido > 0 Or objTributacaoDoc.dCSLLRetido > 0 Or objTributacaoDoc.dIRRFBase > 0 Or objTributacaoDoc.dIRRFValor > 0 Or objTributacaoDoc.dINSSValorBase > 0 Or objTributacaoDoc.iINSSRetido <> 0 Then

                    If objTributacaoDoc.dPISRetido > 0.0 Then
                        a5.infNFe.total.retTrib.vRetPIS = Replace(Format(objTributacaoDoc.dPISRetido, "fixed"), ",", ".")
                    End If
                    If objTributacaoDoc.dCOFINSRetido > 0 Then
                        a5.infNFe.total.retTrib.vRetCOFINS = Replace(Format(objTributacaoDoc.dCOFINSRetido, "fixed"), ",", ".")
                    End If
                    If objTributacaoDoc.dCSLLRetido > 0 Then
                        a5.infNFe.total.retTrib.vRetCSLL = Replace(Format(objTributacaoDoc.dCSLLRetido, "fixed"), ",", ".")
                    End If
                    If objTributacaoDoc.dIRRFBase > 0 Then
                        a5.infNFe.total.retTrib.vBCIRRF = Replace(Format(objTributacaoDoc.dIRRFBase, "fixed"), ",", ".")
                    End If
                    If objTributacaoDoc.dIRRFValor > 0 Then
                        a5.infNFe.total.retTrib.vIRRF = Replace(Format(objTributacaoDoc.dIRRFValor, "fixed"), ",", ".")
                    End If

                    If objTributacaoDoc.iINSSRetido <> 0 Then

                        If objTributacaoDoc.dINSSValorBase > 0 Then
                            a5.infNFe.total.retTrib.vBCRetPrev = Replace(Format(objTributacaoDoc.dINSSValorBase, "fixed"), ",", ".")
                        End If
                        If objTributacaoDoc.dValorINSS > 0 Then
                            a5.infNFe.total.retTrib.vRetPrev = Replace(Format(objTributacaoDoc.dValorINSS, "fixed"), ",", ".")
                        End If

                    End If

                End If

            End If

            Monta_NFiscal_Xml4h = SUCESSO

        Catch ex As Exception

            Monta_NFiscal_Xml4h = 1

            Call gobjApp.GravarLog(ex.Message)

        End Try

    End Function

    Private Function Monta_NFiscal_Xml4(ByVal a5 As TNFe) As Long

        Try
            Dim lErro As Long
            Dim iNumItensNF As Integer, iIndice As Integer
            Dim dValorServPIS As Double, dValorServCOFINS As Double
            Dim dServNTribICMS As Double, dValorPIS As Double, dValorCOFINS As Double, dTotalDescontoItem As Double, dvProdICMS As Double
            Dim objItem As GlobaisLoja.ClassItemCupomFiscal

            If gobjApp.iDebug = 1 Then MsgBox("16")
            gobjApp.sErro = "16"
            gobjApp.sMsg1 = "vai acessar a tabela TributacaoDoc"

            iNumItensNF = gobjVenda.objCupomFiscal.colItens.Count

            Dim NFDet(iNumItensNF) As TNFeInfNFeDet
            a5.infNFe.det() = NFDet

            iIndice = -1

            dValorServPIS = 0
            dValorServCOFINS = 0
            dServNTribICMS = 0
            dValorPIS = 0
            dValorCOFINS = 0

            dTotalDescontoItem = 0
            dvProdICMS = 0

            For Each objItem In gobjVenda.objCupomFiscal.colItens

                'pular itens cancelados
                If objItem.iStatus <> 7 Then

                    iIndice = iIndice + 1

                    If iIndice = 0 Then
                        'v2.00
                        If objItem.objTributacaoDocItem.iRegimeTributario = 1 Then
                            a5.infNFe.emit.CRT = TNFeInfNFeEmitCRT.Item1
                        Else
                            a5.infNFe.emit.CRT = TNFeInfNFeEmitCRT.Item3
                        End If
                    End If

                    Dim infNFeDet As TNFeInfNFeDet = New TNFeInfNFeDet
                    a5.infNFe.det(iIndice) = infNFeDet
                    infNFeDet.nItem = iIndice + 1

                    lErro = Monta_NFiscal_Xml4b(objItem, infNFeDet, dTotalDescontoItem)
                    If lErro <> SUCESSO Then Throw New System.Exception("")

                    Dim infNFeDetImposto As TNFeInfNFeDetImposto = New TNFeInfNFeDetImposto
                    infNFeDet.imposto = infNFeDetImposto

                    If objItem.objTributacaoDocItem.dTotTrib <> 0 Then
                        infNFeDet.imposto.vTotTrib = Replace(Format(objItem.objTributacaoDocItem.dTotTrib, "fixed"), ",", ".")
                    End If

                    Dim objObjeto(2) As Object

                    infNFeDetImposto.Items = objObjeto

                    If Len(Trim(objItem.objTributacaoDocItem.sISSQN)) = 0 Then

                        bComICMS = True

                        If gobjApp.iDebug = 1 Then MsgBox("20")
                        gobjApp.sErro = "20"
                        gobjApp.sMsg1 = "vai iniciar a tributacao de ICMS"

                        If a5.infNFe.det(iIndice).prod.indTot = TNFeInfNFeDetProdIndTot.Item1 Then
                            dvProdICMS = Arredonda_Moeda(dvProdICMS + (objItem.dPrecoUnitario * IIf(objItem.dQuantidade = 0, 1, objItem.dQuantidade)))
                        End If

                        '*************** ICMS ***************************************


                        Dim infNFeDetImpostoICMS As TNFeInfNFeDetImpostoICMS = New TNFeInfNFeDetImpostoICMS
                        '            a5.infNFe.det(iIndice).imposto.ICMS = infNFeDetImpostoICMS

                        objObjeto(0) = infNFeDetImpostoICMS

                        'v2.0 - se o regime tributario for normal
                        If objItem.objTributacaoDocItem.iRegimeTributario = 3 Then

                            lErro = Monta_NFiscal_Xml4c(infNFeDetImpostoICMS, objItem.objTributacaoDocItem)
                            If lErro <> SUCESSO Then Throw New System.Exception("")

                        Else
                            lErro = Monta_NFiscal_Xml4d(infNFeDetImpostoICMS, objItem.objTributacaoDocItem)
                            If lErro <> SUCESSO Then Throw New System.Exception("")

                        End If

                    Else

                        bComISSQN = True

                        dServNTribICMS = dServNTribICMS + (objItem.dPrecoUnitario * IIf(objItem.dQuantidade = 0, 1, objItem.dQuantidade))

                        lErro = Monta_NFiscal_Xml4g(infNFeDetImposto, objItem.objTributacaoDocItem, dValorServPIS, dValorServCOFINS)
                        If lErro <> SUCESSO Then Throw New System.Exception("")

                    End If

                    If sModelo <> "NFCe" Then

                        lErro = Monta_NFiscal_Xml4e(infNFeDetImposto, objItem.objTributacaoDocItem, dValorPIS)
                        If lErro <> SUCESSO Then Throw New System.Exception("")

                        lErro = Monta_NFiscal_Xml4f(infNFeDetImposto, objItem.objTributacaoDocItem, dValorCOFINS)
                        If lErro <> SUCESSO Then Throw New System.Exception("")

                    End If

                End If

            Next

            lErro = Monta_NFiscal_Xml4h(a5, dTotalDescontoItem, dvProdICMS, dValorPIS, dValorServPIS, dValorCOFINS, dValorServCOFINS, dServNTribICMS)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            Monta_NFiscal_Xml4 = SUCESSO

        Catch ex As Exception

            Monta_NFiscal_Xml4 = 1

            Call gobjApp.GravarLog(ex.Message)

        End Try

    End Function

    Private Function Monta_NFiscal_Xml4b(ByVal objItem As GlobaisLoja.ClassItemCupomFiscal, ByVal infNFeDet As TNFeInfNFeDet, ByRef dTotalDescontoItem As Double) As Long

        Dim lErro As Long
        Try

            Dim sProduto As String, objTribDocItem As GlobaisAdm.ClassTributacaoDocItem

            Dim infNFeDetProd As TNFeInfNFeDetProd = New TNFeInfNFeDetProd
            infNFeDet.prod = infNFeDetProd

            sProduto = ""
            Call Formata_Sem_Espaco(Trim(objItem.sProduto), sProduto)
            infNFeDetProd.cProd = sProduto

            If gobjApp.iDebug = 1 Then MsgBox("17.1")
            gobjApp.sErro = "17.1"
            gobjApp.sMsg1 = "vai tratar os dados do produto"

            If sModelo = "NFCe" And gobjApp.iNFeAmbiente = NFE_AMBIENTE_HOMOLOGACAO And infNFeDet.nItem = "1" Then
                infNFeDetProd.xProd = "NOTA FISCAL EMITIDA EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL"
            Else
                infNFeDetProd.xProd = Mid(DesacentuaTexto(Trim(IIf(gobjVenda.objNFeInfo.iNFDescricaoProd = 0, objItem.sProdutoNomeRed, objItem.sProdutoDescricao))), 1, 120)                'se nao for servico
            End If


            If gobjApp.iDebug = 1 Then MsgBox("18")
            gobjApp.sErro = "18"
            gobjApp.sMsg1 = "vai acessar a tabela TributacaoDocItem"

            objTribDocItem = objItem.objTributacaoDocItem

            lErro = Produto_Trata_EAN(objTribDocItem)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            infNFeDetProd.cEAN = objTribDocItem.sEAN
            If Len(Trim(objTribDocItem.sEANTrib)) > 0 Then
                infNFeDetProd.cEANTrib = objTribDocItem.sEANTrib
            Else
                infNFeDetProd.cEANTrib = infNFeDetProd.cEAN
            End If

            If gobjApp.iDebug = 1 Then MsgBox("18.1")
            gobjApp.sErro = "18.1"
            gobjApp.sMsg1 = "vai acessar a tabela NaturezaOp"

            'infNFeDetProd.CFOP = GetCode(Of TCfop)(objTribDocItem.sNaturezaOp)
            infNFeDetProd.CFOP = objTribDocItem.sNaturezaOp
            If Len(objTribDocItem.sFCI) = 36 Then
                infNFeDetProd.nFCI = objTribDocItem.sFCI
            End If

            If gobjApp.iDebug = 1 Then MsgBox("19")

            gobjApp.sErro = "19"
            gobjApp.sMsg1 = "carrega os dados do Produto"


            infNFeDetProd.uCom = objItem.sUnidadeMed
            infNFeDetProd.qCom = Replace(Format(objTribDocItem.dQuantidade, "######0.0000"), ",", ".")
            infNFeDetProd.vUnCom = Replace(Format(objTribDocItem.dPrecoUnitario, "#########0.0000######"), ",", ".")
            infNFeDetProd.vProd = Replace(Format((objTribDocItem.dPrecoUnitario * IIf(objTribDocItem.dQuantidade = 0, 1, objTribDocItem.dQuantidade)), "fixed"), ",", ".")

            If Len(Trim(objTribDocItem.sIPICodProduto)) <> 0 Then
                infNFeDetProd.NCM = Right(Trim(objTribDocItem.sIPICodProduto), 8)
            Else
                If Len(Trim(objTribDocItem.sISSQN)) = 0 Then
                    infNFeDetProd.NCM = "00000000"
                Else
                    infNFeDetProd.NCM = "00"
                End If
            End If

            If gobjVenda.objCupomFiscal.dtDataEmissao >= #4/1/2016# And Len(Trim(objTribDocItem.sCEST)) <> 0 Then
                infNFeDetProd.CEST = Trim(objTribDocItem.sCEST)
            End If

            If objTribDocItem.sCNPJFab <> "" Then
                infNFeDetProd.CNPJFab = objTribDocItem.sCNPJFab
            End If

            If objTribDocItem.sindEscala = "" Then
                infNFeDetProd.indEscalaSpecified = False
            Else
                If objTribDocItem.sindEscala = "S" Then
                    infNFeDetProd.indEscala = TNFeInfNFeDetProdIndEscala.S
                Else
                    infNFeDetProd.indEscala = TNFeInfNFeDetProdIndEscala.N
                End If
                infNFeDetProd.indEscalaSpecified = True
            End If

            If objTribDocItem.scBenef <> "" Then
                infNFeDetProd.cBenef = objTribDocItem.scBenef
            End If

            If objTribDocItem.sUMTrib = "" Then objTribDocItem.sUMTrib = objItem.sUnidadeMed '????????????
            infNFeDetProd.uTrib = objTribDocItem.sUMTrib
            infNFeDetProd.qTrib = Replace(Format(objTribDocItem.dQtdTrib, "######0.0000"), ",", ".")
            infNFeDetProd.vUnTrib = Replace(Format(objTribDocItem.dValorUnitTrib, "#########0.0000######"), ",", ".")
            'If objItemNF.ValorDesconto > 0 Then
            '    infNFeDetProd.vDesc = Replace(Format(objItemNF.ValorDesconto, "fixed"), ",", ".")
            'End If

            If objTribDocItem.dDescontoGrid > 0 Then
                infNFeDetProd.vDesc = Replace(Format(objTribDocItem.dDescontoGrid, "fixed"), ",", ".")
                dTotalDescontoItem = dTotalDescontoItem + CDbl(Format(objTribDocItem.dDescontoGrid, "fixed"))
            End If


            If objTribDocItem.dValorFreteItem > 0 Then
                infNFeDetProd.vFrete = Replace(Format(objTribDocItem.dValorFreteItem, "fixed"), ",", ".")
            End If

            If objTribDocItem.dValorSeguroItem > 0 Then
                infNFeDetProd.vSeg = Replace(Format(objTribDocItem.dValorSeguroItem, "fixed"), ",", ".")
            End If

            'v2.00
            If objTribDocItem.dValorOutrasDespesasItem > 0 Then
                infNFeDetProd.vOutro = Replace(Format(objTribDocItem.dValorOutrasDespesasItem, "fixed"), ",", ".")
            End If

            'v2.00
            infNFeDetProd.indTot = TNFeInfNFeDetProdIndTot.Item1

            Monta_NFiscal_Xml4b = SUCESSO

        Catch ex As Exception

            Monta_NFiscal_Xml4b = 1

            Call gobjApp.GravarLog(ex.Message)

        End Try

    End Function

    Private Function Monta_NFiscal_Xml4c(ByVal infNFeDetImpostoICMS As TNFeInfNFeDetImpostoICMS, ByVal objTribDocItem As GlobaisAdm.ClassTributacaoDocItem) As Long

        Try

            Select Case objTribDocItem.iICMSTipo

                'tributacao integral
                Case 1
                    Dim ICMS00 As New TNFeInfNFeDetImpostoICMSICMS00
                    infNFeDetImpostoICMS.Item = ICMS00
                    '                                ICMS00.orig = objProduto.OrigemMercadoria
                    ICMS00.orig = objTribDocItem.iOrigemMercadoria
                    ICMS00.CST = TNFeInfNFeDetImpostoICMSICMS00CST.Item00
                    ICMS00.modBC = TNFeInfNFeDetImpostoICMSICMS00ModBC.Item3
                    ICMS00.vBC = Replace(Format(objTribDocItem.dICMSBase, "fixed"), ",", ".")
                    ICMS00.pICMS = Replace(Format((objTribDocItem.dICMSAliquota - objTribDocItem.dICMSpFCP) * 100, "##0.00"), ",", ".")
                    ICMS00.vICMS = Replace(Format((objTribDocItem.dICMSValor - objTribDocItem.dICMSvFCP), "fixed"), ",", ".")

                    If objTribDocItem.dICMSvFCP <> 0 Then
                        ICMS00.pFCP = Replace(Format(objTribDocItem.dICMSpFCP * 100, "##0.00"), ",", ".")
                        ICMS00.vFCP = Replace(Format(objTribDocItem.dICMSvFCP, "fixed"), ",", ".")
                    End If

                    'Tributado com substituição
                Case 6
                    Dim ICMS10 As New TNFeInfNFeDetImpostoICMSICMS10
                    infNFeDetImpostoICMS.Item = ICMS10
                    '                                ICMS10.orig = objProduto.OrigemMercadoria
                    ICMS10.orig = objTribDocItem.iOrigemMercadoria
                    ICMS10.CST = TNFeInfNFeDetImpostoICMSICMS10CST.Item10
                    ICMS10.modBC = TNFeInfNFeDetImpostoICMSICMS00ModBC.Item3
                    ICMS10.vBC = Replace(Format(objTribDocItem.dICMSBase, "fixed"), ",", ".")
                    ICMS10.pICMS = Replace(Format((objTribDocItem.dICMSAliquota - objTribDocItem.dICMSpFCP) * 100, "##0.00"), ",", ".")
                    ICMS10.vICMS = Replace(Format((objTribDocItem.dICMSValor - objTribDocItem.dICMSvFCP), "fixed"), ",", ".")
                    ICMS10.modBCST = TNFeInfNFeDetImpostoICMSICMS10ModBCST.Item4
                    If objTribDocItem.dICMSSubstPercMVA > 0 Then
                        ICMS10.pMVAST = Replace(Format(objTribDocItem.dICMSSubstPercMVA * 100, "##0.00"), ",", ".")
                    End If
                    ICMS10.vBCST = Replace(Format(objTribDocItem.dICMSSubstBase, "fixed"), ",", ".")
                    ICMS10.pICMSST = Replace(Format((objTribDocItem.dICMSSubstAliquota - objTribDocItem.dICMSpFCPST) * 100, "##0.00"), ",", ".")
                    ICMS10.vICMSST = Replace(Format((objTribDocItem.dICMSSubstValor - objTribDocItem.dICMSvFCPST), "fixed"), ",", ".")
                    If objTribDocItem.dICMSSubstPercRedBase > 0 Then
                        ICMS10.pRedBCST = Replace(Format(objTribDocItem.dICMSSubstPercRedBase * 100, "##0.00"), ",", ".")
                    End If

                    If objTribDocItem.dICMSvFCP <> 0 Then
                        ICMS10.vBCFCP = Replace(Format(objTribDocItem.dICMSvBCFCP, "fixed"), ",", ".")
                        ICMS10.pFCP = Replace(Format(objTribDocItem.dICMSpFCP * 100, "##0.00"), ",", ".")
                        ICMS10.vFCP = Replace(Format(objTribDocItem.dICMSvFCP, "fixed"), ",", ".")
                    End If

                    If objTribDocItem.dICMSvFCPST <> 0 Then
                        ICMS10.vBCFCPST = Replace(Format(objTribDocItem.dICMSvBCFCPST, "fixed"), ",", ".")
                        ICMS10.pFCPST = Replace(Format(objTribDocItem.dICMSpFCPST * 100, "##0.00"), ",", ".")
                        ICMS10.vFCPST = Replace(Format(objTribDocItem.dICMSvFCPST, "fixed"), ",", ".")
                    End If

                    'Com redução da base de calc.
                Case 7
                    Dim ICMS20 As New TNFeInfNFeDetImpostoICMSICMS20
                    infNFeDetImpostoICMS.Item = ICMS20
                    'ICMS20.orig = objProduto.OrigemMercadoria
                    ICMS20.orig = objTribDocItem.iOrigemMercadoria
                    ICMS20.CST = TNFeInfNFeDetImpostoICMSICMS20CST.Item20
                    ICMS20.modBC = TNFeInfNFeDetImpostoICMSICMS00ModBC.Item3
                    If objTribDocItem.dICMSPercRedBase > 0 Then
                        ICMS20.pRedBC = Replace(Format(objTribDocItem.dICMSPercRedBase * 100, "##0.00"), ",", ".")
                    Else
                        ICMS20.pRedBC = 0
                    End If
                    ICMS20.vBC = Replace(Format(objTribDocItem.dICMSBase * (1 - objTribDocItem.dICMSPercRedBase), "fixed"), ",", ".")
                    ICMS20.pICMS = Replace(Format((objTribDocItem.dICMSAliquota - objTribDocItem.dICMSpFCP) * 100, "##0.00"), ",", ".")
                    ICMS20.vICMS = Replace(Format((objTribDocItem.dICMSValor - objTribDocItem.dICMSvFCP), "fixed"), ",", ".")

                    If objTribDocItem.dICMSvFCP <> 0 Then
                        ICMS20.vBCFCP = Replace(Format(objTribDocItem.dICMSvBCFCP, "fixed"), ",", ".")
                        ICMS20.pFCP = Replace(Format(objTribDocItem.dICMSpFCP * 100, "##0.00"), ",", ".")
                        ICMS20.vFCP = Replace(Format(objTribDocItem.dICMSvFCP, "fixed"), ",", ".")
                    End If

                    'Isento com cobrança por subst.
                    'Não trib com cobrança por subst.
                Case 9, 10
                    Dim ICMS30 As New TNFeInfNFeDetImpostoICMSICMS30
                    infNFeDetImpostoICMS.Item = ICMS30
                    'ICMS30.orig = objProduto.OrigemMercadoria
                    ICMS30.orig = objTribDocItem.iOrigemMercadoria
                    ICMS30.CST = TNFeInfNFeDetImpostoICMSICMS30CST.Item30
                    ICMS30.modBCST = TNFeInfNFeDetImpostoICMSICMS10ModBCST.Item4
                    If objTribDocItem.dICMSSubstPercMVA > 0 Then
                        ICMS30.pMVAST = Replace(Format(objTribDocItem.dICMSSubstPercMVA * 100, "##0.00"), ",", ".")
                    End If
                    ICMS30.vBCST = Replace(Format(objTribDocItem.dICMSSubstBase, "fixed"), ",", ".")
                    ICMS30.pICMSST = Replace(Format((objTribDocItem.dICMSSubstAliquota - objTribDocItem.dICMSpFCPST) * 100, "##0.00"), ",", ".")
                    ICMS30.vICMSST = Replace(Format((objTribDocItem.dICMSSubstValor - objTribDocItem.dICMSvFCPST), "fixed"), ",", ".")
                    If objTribDocItem.dICMSSubstPercRedBase > 0 Then
                        ICMS30.pRedBCST = Replace(Format(objTribDocItem.dICMSSubstPercRedBase * 100, "##0.00"), ",", ".")
                    End If

                    If objTribDocItem.dICMSvFCPST <> 0 Then
                        ICMS30.vBCFCPST = Replace(Format(objTribDocItem.dICMSvBCFCPST, "fixed"), ",", ".")
                        ICMS30.pFCPST = Replace(Format(objTribDocItem.dICMSpFCPST * 100, "##0.00"), ",", ".")
                        ICMS30.vFCPST = Replace(Format(objTribDocItem.dICMSvFCPST, "fixed"), ",", ".")
                    End If

                    'Isenta
                Case 2
                    Dim ICMS40 As New TNFeInfNFeDetImpostoICMSICMS40
                    infNFeDetImpostoICMS.Item = ICMS40
                    ICMS40.orig = objTribDocItem.iOrigemMercadoria
                    ICMS40.CST = TNFeInfNFeDetImpostoICMSICMS40CST.Item40
                    'ICMS40.motDesICMSSpecified = False
                    If objTribDocItem.iICMSMotivo <> 0 Then
                        ICMS40.vICMSDeson = Replace(Format(objTribDocItem.dICMSValorIsento, "fixed"), ",", ".")
                        ICMS40.motDesICMS = objTribDocItem.iICMSMotivo
                        'ICMS40.motDesICMSSpecified = True
                    End If

                    'Não Tributado
                Case 0

                    Dim ICMS40 As New TNFeInfNFeDetImpostoICMSICMS40
                    infNFeDetImpostoICMS.Item = ICMS40
                    ICMS40.orig = objTribDocItem.iOrigemMercadoria
                    ICMS40.CST = TNFeInfNFeDetImpostoICMSICMS40CST.Item41
                    ICMS40.vICMSDeson = Nothing
                    'ICMS40.motDesICMSSpecified = False
                    If objTribDocItem.iICMSMotivo <> 0 Then
                        ICMS40.vICMSDeson = Replace(Format(objTribDocItem.dICMSValorIsento, "fixed"), ",", ".")
                        ICMS40.motDesICMS = objTribDocItem.iICMSMotivo
                        'ICMS40.motDesICMSSpecified = True
                    End If

                    'Com suspensão
                Case 3

                    Dim ICMS40 As New TNFeInfNFeDetImpostoICMSICMS40
                    infNFeDetImpostoICMS.Item = ICMS40
                    '                                ICMS40.orig = objProduto.OrigemMercadoria
                    ICMS40.orig = objTribDocItem.iOrigemMercadoria
                    ICMS40.CST = TNFeInfNFeDetImpostoICMSICMS40CST.Item50
                    'ICMS40.motDesICMSSpecified = False
                    If objTribDocItem.iICMSMotivo <> 0 Then
                        ICMS40.vICMSDeson = Replace(Format(objTribDocItem.dICMSValorIsento, "fixed"), ",", ".")
                        ICMS40.motDesICMS = objTribDocItem.iICMSMotivo
                        'ICMS40.motDesICMSSpecified = True
                    End If

                    'Com diferimento
                Case 5
                    Dim ICMS51 As New TNFeInfNFeDetImpostoICMSICMS51
                    infNFeDetImpostoICMS.Item = ICMS51
                    'ICMS51.orig = objProduto.OrigemMercadoria
                    ICMS51.orig = objTribDocItem.iOrigemMercadoria
                    ICMS51.CST = TNFeInfNFeDetImpostoICMSICMS51CST.Item51
                    ICMS51.modBC = TNFeInfNFeDetImpostoICMSICMS00ModBC.Item3
                    If objTribDocItem.dICMSPercRedBase > 0 Then
                        ICMS51.pRedBC = Replace(Format(objTribDocItem.dICMSPercRedBase * 100, "##0.00"), ",", ".")
                    End If
                    ICMS51.vBC = Replace(Format(objTribDocItem.dICMSBase * (1 - objTribDocItem.dICMSPercRedBase), "fixed"), ",", ".")
                    ICMS51.pICMS = Replace(Format((objTribDocItem.dICMSAliquota - objTribDocItem.dICMSpFCP) * 100, "##0.00"), ",", ".")
                    ICMS51.vICMS = Replace(Format((objTribDocItem.dICMSValor - objTribDocItem.dICMSvFCP), "fixed"), ",", ".")

                    'nfe 3.10
                    'diferimento parcial
                    If objTribDocItem.dICMSPercDifer <> 0 And objTribDocItem.dICMSPercDifer <> 1 Then
                        ICMS51.vICMSOp = Replace(Format(objTribDocItem.dICMS51ValorOp, "fixed"), ",", ".")
                        ICMS51.pDif = Replace(Format(objTribDocItem.dICMSPercDifer * 100, "##0.00"), ",", ".")
                        ICMS51.vICMSDif = Replace(Format(objTribDocItem.dICMSValorDif, "fixed"), ",", ".")
                    End If
                    'fim nfe 3.10

                    If objTribDocItem.dICMSvFCP <> 0 Then
                        ICMS51.vBCFCP = Replace(Format(objTribDocItem.dICMSvBCFCP, "fixed"), ",", ".")
                        ICMS51.pFCP = Replace(Format(objTribDocItem.dICMSpFCP * 100, "##0.00"), ",", ".")
                        ICMS51.vFCP = Replace(Format(objTribDocItem.dICMSvFCP, "fixed"), ",", ".")
                    End If

                    'Cobrado anteriormente por subst.
                Case 8
                    Dim ICMS60 As New TNFeInfNFeDetImpostoICMSICMS60
                    infNFeDetImpostoICMS.Item = ICMS60
                    'ICMS60.orig = objProduto.OrigemMercadoria
                    ICMS60.orig = objTribDocItem.iOrigemMercadoria
                    ICMS60.CST = TNFeInfNFeDetImpostoICMSICMS60CST.Item60
                    ICMS60.vBCSTRet = Replace(Format(objTribDocItem.dICMSSTCobrAntBase, "fixed"), ",", ".")
                    ICMS60.vICMSSTRet = Replace(Format((objTribDocItem.dICMSSTCobrAntValor - objTribDocItem.dICMSvFCPSTRet), "fixed"), ",", ".")

                    If objTribDocItem.dICMSvFCPSTRet <> 0 Then
                        ICMS60.vBCFCPSTRet = Replace(Format(objTribDocItem.dICMSvBCFCPSTRet, "fixed"), ",", ".")
                        ICMS60.pFCPSTRet = Replace(Format(objTribDocItem.dICMSpFCPSTRet * 100, "##0.00"), ",", ".")
                        ICMS60.vFCPSTRet = Replace(Format(objTribDocItem.dICMSvFCPSTRet, "fixed"), ",", ".")
                    End If

                    ICMS60.pST = Replace(Format((objTribDocItem.dICMSSTCobrAntAliquota - objTribDocItem.dICMSpFCPSTRet) * 100, "##0.00"), ",", ".")

                    'Com redução da base e cobr. Por subst.
                Case 4
                    Dim ICMS70 As New TNFeInfNFeDetImpostoICMSICMS70
                    infNFeDetImpostoICMS.Item = ICMS70

                    '                                ICMS70.orig = objProduto.OrigemMercadoria
                    ICMS70.orig = objTribDocItem.iOrigemMercadoria
                    ICMS70.CST = TNFeInfNFeDetImpostoICMSICMS70CST.Item70
                    ICMS70.modBC = TNFeInfNFeDetImpostoICMSICMS00ModBC.Item3
                    ICMS70.pRedBC = Replace(Format(objTribDocItem.dICMSPercRedBase * 100, "##0.00"), ",", ".")
                    ICMS70.vBC = Replace(Format(objTribDocItem.dICMSBase * (1 - objTribDocItem.dICMSPercRedBase), "fixed"), ",", ".")
                    ICMS70.pICMS = Replace(Format((objTribDocItem.dICMSAliquota - objTribDocItem.dICMSpFCP) * 100, "##0.00"), ",", ".")
                    ICMS70.vICMS = Replace(Format((objTribDocItem.dICMSValor - objTribDocItem.dICMSvFCP), "fixed"), ",", ".")
                    ICMS70.modBCST = TNFeInfNFeDetImpostoICMSICMS10ModBCST.Item4
                    If objTribDocItem.dICMSSubstPercMVA > 0 Then
                        ICMS70.pMVAST = Replace(Format(objTribDocItem.dICMSSubstPercMVA * 100, "##0.00"), ",", ".")
                    End If
                    ICMS70.vBCST = Replace(Format(objTribDocItem.dICMSSubstBase, "fixed"), ",", ".")
                    ICMS70.pICMSST = Replace(Format((objTribDocItem.dICMSSubstAliquota - objTribDocItem.dICMSpFCPST) * 100, "##0.00"), ",", ".")
                    ICMS70.vICMSST = Replace(Format((objTribDocItem.dICMSSubstValor - objTribDocItem.dICMSvFCPST), "fixed"), ",", ".")
                    If objTribDocItem.dICMSSubstPercRedBase > 0 Then
                        ICMS70.pRedBCST = Replace(Format(objTribDocItem.dICMSSubstPercRedBase * 100, "##0.00"), ",", ".")
                    End If

                    If objTribDocItem.dICMSvFCP <> 0 Then
                        ICMS70.vBCFCP = Replace(Format(objTribDocItem.dICMSvBCFCP, "fixed"), ",", ".")
                        ICMS70.pFCP = Replace(Format(objTribDocItem.dICMSpFCP * 100, "##0.00"), ",", ".")
                        ICMS70.vFCP = Replace(Format(objTribDocItem.dICMSvFCP, "fixed"), ",", ".")
                    End If

                    If objTribDocItem.dICMSvFCPST <> 0 Then
                        ICMS70.vBCFCPST = Replace(Format(objTribDocItem.dICMSvBCFCPST, "fixed"), ",", ".")
                        ICMS70.pFCPST = Replace(Format(objTribDocItem.dICMSpFCPST * 100, "##0.00"), ",", ".")
                        ICMS70.vFCPST = Replace(Format(objTribDocItem.dICMSvFCPST, "fixed"), ",", ".")
                    End If

                    'Outras
                Case 99
                    Dim ICMS90 As New TNFeInfNFeDetImpostoICMSICMS90
                    infNFeDetImpostoICMS.Item = ICMS90
                    '                                ICMS90.orig = objProduto.OrigemMercadoria
                    ICMS90.orig = objTribDocItem.iOrigemMercadoria
                    ICMS90.CST = TNFeInfNFeDetImpostoICMSICMS90CST.Item90
                    ICMS90.modBC = TNFeInfNFeDetImpostoICMSICMS00ModBC.Item3
                    If objTribDocItem.dICMSPercRedBase > 0 Then
                        ICMS90.pRedBC = Replace(Format(objTribDocItem.dICMSPercRedBase * 100, "##0.00"), ",", ".")
                    End If
                    ICMS90.vBC = Replace(Format(objTribDocItem.dICMSBase * (1 - objTribDocItem.dICMSPercRedBase), "fixed"), ",", ".")
                    ICMS90.pICMS = Replace(Format((objTribDocItem.dICMSAliquota - objTribDocItem.dICMSpFCP) * 100, "##0.00"), ",", ".")
                    ICMS90.vICMS = Replace(Format((objTribDocItem.dICMSValor - objTribDocItem.dICMSvFCP), "fixed"), ",", ".")
                    ICMS90.modBCST = TNFeInfNFeDetImpostoICMSICMS10ModBCST.Item4
                    If objTribDocItem.dICMSSubstPercMVA > 0 Then
                        ICMS90.pMVAST = Replace(Format(objTribDocItem.dICMSSubstPercMVA * 100, "##0.00"), ",", ".")
                    End If
                    ICMS90.vBCST = Replace(Format(objTribDocItem.dICMSSubstBase, "fixed"), ",", ".")
                    ICMS90.pICMSST = Replace(Format((objTribDocItem.dICMSSubstAliquota - objTribDocItem.dICMSpFCPST) * 100, "##0.00"), ",", ".")
                    ICMS90.vICMSST = Replace(Format((objTribDocItem.dICMSSubstValor - objTribDocItem.dICMSvFCPST), "fixed"), ",", ".")
                    If objTribDocItem.dICMSSubstPercRedBase > 0 Then
                        ICMS90.pRedBCST = Replace(Format(objTribDocItem.dICMSSubstPercRedBase * 100, "##0.00"), ",", ".")
                    End If

                    If objTribDocItem.dICMSvFCP <> 0 Then
                        ICMS90.vBCFCP = Replace(Format(objTribDocItem.dICMSvBCFCP, "fixed"), ",", ".")
                        ICMS90.pFCP = Replace(Format(objTribDocItem.dICMSpFCP * 100, "##0.00"), ",", ".")
                        ICMS90.vFCP = Replace(Format(objTribDocItem.dICMSvFCP, "fixed"), ",", ".")
                    End If

                    If objTribDocItem.dICMSvFCPST <> 0 Then
                        ICMS90.vBCFCPST = Replace(Format(objTribDocItem.dICMSvBCFCPST, "fixed"), ",", ".")
                        ICMS90.pFCPST = Replace(Format(objTribDocItem.dICMSpFCPST * 100, "##0.00"), ",", ".")
                        ICMS90.vFCPST = Replace(Format(objTribDocItem.dICMSvFCPST, "fixed"), ",", ".")
                    End If

                    'Partilha do ICMS - Não trib com cobrança por subst
                Case 11
                    Dim ICMSPart As New TNFeInfNFeDetImpostoICMSICMSPart
                    infNFeDetImpostoICMS.Item = ICMSPart
                    ICMSPart.orig = objTribDocItem.iOrigemMercadoria
                    ICMSPart.CST = TNFeInfNFeDetImpostoICMSICMSPartCST.Item10
                    ICMSPart.modBC = TNFeInfNFeDetImpostoICMSICMSPartModBC.Item3
                    ICMSPart.vBC = Replace(Format(objTribDocItem.dICMSBase, "fixed"), ",", ".")
                    ICMSPart.pICMS = Replace(Format(objTribDocItem.dICMSAliquota * 100, "##0.00"), ",", ".")
                    ICMSPart.vICMS = Replace(Format(objTribDocItem.dICMSValor, "fixed"), ",", ".")
                    ICMSPart.modBCST = TNFeInfNFeDetImpostoICMSICMSPartModBCST.Item4
                    If objTribDocItem.dICMSSubstPercMVA > 0 Then
                        ICMSPart.pMVAST = Replace(Format(objTribDocItem.dICMSSubstPercMVA * 100, "##0.00"), ",", ".")
                    End If
                    ICMSPart.vBCST = Replace(Format(objTribDocItem.dICMSSubstBase, "fixed"), ",", ".")
                    ICMSPart.pICMSST = Replace(Format(objTribDocItem.dICMSSubstAliquota * 100, "##0.00"), ",", ".")
                    ICMSPart.vICMSST = Replace(Format(objTribDocItem.dICMSSubstValor, "fixed"), ",", ".")
                    If objTribDocItem.dICMSSubstPercRedBase > 0 Then
                        ICMSPart.pRedBCST = Replace(Format(objTribDocItem.dICMSSubstPercRedBase * 100, "##0.00"), ",", ".")
                    End If
                    ICMSPart.pBCOp = Replace(Format(objTribDocItem.dICMSpercBaseOperacaoPropria * 100, "##0.00"), ",", ".")
                    ICMSPart.UFST = GetCode(Of TUf)(objTribDocItem.sICMSUFDevidoST)

                    'Partilha do ICMS - Outras
                Case 13
                    Dim ICMSPart As New TNFeInfNFeDetImpostoICMSICMSPart
                    infNFeDetImpostoICMS.Item = ICMSPart
                    '                                ICMS10.orig = objProduto.OrigemMercadoria
                    ICMSPart.orig = objTribDocItem.iOrigemMercadoria
                    ICMSPart.CST = TNFeInfNFeDetImpostoICMSICMSPartCST.Item90
                    ICMSPart.modBC = TNFeInfNFeDetImpostoICMSICMSPartModBC.Item3
                    ICMSPart.vBC = Replace(Format(objTribDocItem.dICMSBase, "fixed"), ",", ".")
                    If objTribDocItem.dICMSPercRedBase > 0 Then
                        ICMSPart.pRedBC = Replace(Format(objTribDocItem.dICMSPercRedBase * 100, "##0.00"), ",", ".")
                    End If
                    ICMSPart.pICMS = Replace(Format(objTribDocItem.dICMSAliquota * 100, "##0.00"), ",", ".")
                    ICMSPart.vICMS = Replace(Format(objTribDocItem.dICMSValor, "fixed"), ",", ".")
                    ICMSPart.modBCST = TNFeInfNFeDetImpostoICMSICMSPartModBCST.Item4
                    If objTribDocItem.dICMSSubstPercMVA > 0 Then
                        ICMSPart.pMVAST = Replace(Format(objTribDocItem.dICMSSubstPercMVA * 100, "##0.00"), ",", ".")
                    End If
                    ICMSPart.vBCST = Replace(Format(objTribDocItem.dICMSSubstBase, "fixed"), ",", ".")
                    ICMSPart.pICMSST = Replace(Format(objTribDocItem.dICMSSubstAliquota * 100, "##0.00"), ",", ".")
                    ICMSPart.vICMSST = Replace(Format(objTribDocItem.dICMSSubstValor, "fixed"), ",", ".")
                    If objTribDocItem.dICMSSubstPercRedBase > 0 Then
                        ICMSPart.pRedBCST = Replace(Format(objTribDocItem.dICMSSubstPercRedBase * 100, "##0.00"), ",", ".")
                    End If

                    ICMSPart.pBCOp = Replace(Format(objTribDocItem.dICMSpercBaseOperacaoPropria * 100, "##0.00"), ",", ".")
                    ICMSPart.UFST = GetCode(Of TUf)(objTribDocItem.sICMSUFDevidoST)

                    'repasse de ICMSST retido ant. - Não tributado
                Case 12
                    Dim ICMSST As New TNFeInfNFeDetImpostoICMSICMSST
                    infNFeDetImpostoICMS.Item = ICMSST
                    ICMSST.orig = objTribDocItem.iOrigemMercadoria
                    ICMSST.CST = TNFeInfNFeDetImpostoICMSICMSSTCST.Item41
                    ICMSST.vBCSTRet = Replace(Format(objTribDocItem.dICMSvBCSTRet, "fixed"), ",", ".")
                    ICMSST.vICMSSTRet = Replace(Format(objTribDocItem.dICMSvICMSSTRet, "fixed"), ",", ".")
                    ICMSST.vBCSTDest = Replace(Format(objTribDocItem.dICMSvBCSTDest, "fixed"), ",", ".")
                    ICMSST.vICMSSTDest = Replace(Format(objTribDocItem.dICMSvICMSSTDest, "fixed"), ",", ".")

            End Select

            Monta_NFiscal_Xml4c = SUCESSO

        Catch ex As Exception

            Monta_NFiscal_Xml4c = 1

            Call gobjApp.GravarLog(ex.Message)

        End Try

    End Function

    Private Function Monta_NFiscal_Xml4d(ByVal infNFeDetImpostoICMS As TNFeInfNFeDetImpostoICMS, ByVal objTribDocItem As GlobaisAdm.ClassTributacaoDocItem) As Long

        Try

            'v2.0 - se for regime tributario simples
            If objTribDocItem.iRegimeTributario = 1 Or objTribDocItem.iRegimeTributario = 2 Then

                Select Case objTribDocItem.iICMSSimplesTipo

                    'Trib. pelo Simples permissão de crédito
                    Case 1
                        Dim ICMSSN101 As New TNFeInfNFeDetImpostoICMSICMSSN101
                        infNFeDetImpostoICMS.Item = ICMSSN101
                        ICMSSN101.orig = objTribDocItem.iOrigemMercadoria
                        ICMSSN101.CSOSN = TNFeInfNFeDetImpostoICMSICMSSN101CSOSN.Item101
                        ICMSSN101.pCredSN = Replace(Format(objTribDocItem.dICMSpCredSN * 100, "##0.00"), ",", ".")
                        ICMSSN101.vCredICMSSN = Replace(Format(objTribDocItem.dICMSvCredSN, "fixed"), ",", ".")

                        'Trib. pelo Simples s/permissão de crédito
                    Case 2
                        Dim ICMSSN102 As New TNFeInfNFeDetImpostoICMSICMSSN102
                        infNFeDetImpostoICMS.Item = ICMSSN102
                        ICMSSN102.orig = objTribDocItem.iOrigemMercadoria
                        ICMSSN102.CSOSN = TNFeInfNFeDetImpostoICMSICMSSN102CSOSN.Item102

                        'Isenção do ICMS no Simples Nacional
                    Case 3
                        Dim ICMSSN102 As New TNFeInfNFeDetImpostoICMSICMSSN102
                        infNFeDetImpostoICMS.Item = ICMSSN102
                        ICMSSN102.orig = objTribDocItem.iOrigemMercadoria
                        ICMSSN102.CSOSN = TNFeInfNFeDetImpostoICMSICMSSN102CSOSN.Item103

                        'Simples Nacional - Imune
                    Case 7
                        Dim ICMSSN102 As New TNFeInfNFeDetImpostoICMSICMSSN102
                        infNFeDetImpostoICMS.Item = ICMSSN102
                        ICMSSN102.orig = objTribDocItem.iOrigemMercadoria
                        ICMSSN102.CSOSN = TNFeInfNFeDetImpostoICMSICMSSN102CSOSN.Item300

                        'Não tributada pelo Simples Nacional
                    Case 8
                        Dim ICMSSN102 As New TNFeInfNFeDetImpostoICMSICMSSN102
                        infNFeDetImpostoICMS.Item = ICMSSN102
                        ICMSSN102.orig = objTribDocItem.iOrigemMercadoria
                        ICMSSN102.CSOSN = TNFeInfNFeDetImpostoICMSICMSSN102CSOSN.Item400

                        'Simples c/permissão cred. e cobr. ICMS ST
                    Case 4
                        Dim ICMSSN201 As New TNFeInfNFeDetImpostoICMSICMSSN201
                        infNFeDetImpostoICMS.Item = ICMSSN201
                        '                                ICMS10.orig = objProduto.OrigemMercadoria
                        ICMSSN201.orig = objTribDocItem.iOrigemMercadoria
                        ICMSSN201.CSOSN = TNFeInfNFeDetImpostoICMSICMSSN201CSOSN.Item201
                        ICMSSN201.modBCST = TNFeInfNFeDetImpostoICMSICMSSN201ModBCST.Item4
                        If objTribDocItem.dICMSSubstPercMVA > 0 Then
                            ICMSSN201.pMVAST = Replace(Format(objTribDocItem.dICMSSubstPercMVA * 100, "##0.00"), ",", ".")
                        End If
                        If objTribDocItem.dICMSSubstPercRedBase > 0 Then
                            ICMSSN201.pRedBCST = Replace(Format(objTribDocItem.dICMSSubstPercRedBase * 100, "##0.00"), ",", ".")
                        End If
                        ICMSSN201.vBCST = Replace(Format(objTribDocItem.dICMSSubstBase, "fixed"), ",", ".")
                        ICMSSN201.pICMSST = Replace(Format((objTribDocItem.dICMSSubstAliquota - objTribDocItem.dICMSpFCPST) * 100, "##0.00"), ",", ".")
                        ICMSSN201.vICMSST = Replace(Format((objTribDocItem.dICMSSubstValor - objTribDocItem.dICMSvFCPST), "fixed"), ",", ".")
                        ICMSSN201.pCredSN = Replace(Format(objTribDocItem.dICMSpCredSN * 100, "##0.00"), ",", ".")
                        ICMSSN201.vCredICMSSN = Replace(Format(objTribDocItem.dICMSvCredSN, "fixed"), ",", ".")

                        If objTribDocItem.dICMSvFCPST <> 0 Then
                            ICMSSN201.vBCFCPST = Replace(Format(objTribDocItem.dICMSvBCFCPST, "fixed"), ",", ".")
                            ICMSSN201.pFCPST = Replace(Format(objTribDocItem.dICMSpFCPST * 100, "##0.00"), ",", ".")
                            ICMSSN201.vFCPST = Replace(Format(objTribDocItem.dICMSvFCPST, "fixed"), ",", ".")
                        End If

                        'Simples s/permissão cred. e cobr. ICMS ST
                    Case 5
                        Dim ICMSSN202 As New TNFeInfNFeDetImpostoICMSICMSSN202
                        infNFeDetImpostoICMS.Item = ICMSSN202
                        '                                ICMS10.orig = objProduto.OrigemMercadoria
                        ICMSSN202.orig = objTribDocItem.iOrigemMercadoria
                        ICMSSN202.CSOSN = TNFeInfNFeDetImpostoICMSICMSSN202CSOSN.Item202
                        ICMSSN202.modBCST = TNFeInfNFeDetImpostoICMSICMSSN202ModBCST.Item4
                        If objTribDocItem.dICMSSubstPercMVA > 0 Then
                            ICMSSN202.pMVAST = Replace(Format(objTribDocItem.dICMSSubstPercMVA * 100, "##0.00"), ",", ".")
                        End If
                        If objTribDocItem.dICMSSubstPercRedBase > 0 Then
                            ICMSSN202.pRedBCST = Replace(Format(objTribDocItem.dICMSSubstPercRedBase * 100, "##0.00"), ",", ".")
                        End If
                        ICMSSN202.vBCST = Replace(Format(objTribDocItem.dICMSSubstBase, "fixed"), ",", ".")
                        ICMSSN202.pICMSST = Replace(Format((objTribDocItem.dICMSSubstAliquota - objTribDocItem.dICMSpFCPST) * 100, "##0.00"), ",", ".")
                        ICMSSN202.vICMSST = Replace(Format((objTribDocItem.dICMSSubstValor - objTribDocItem.dICMSvFCPST), "fixed"), ",", ".")

                        If objTribDocItem.dICMSvFCPST <> 0 Then
                            ICMSSN202.vBCFCPST = Replace(Format(objTribDocItem.dICMSvBCFCPST, "fixed"), ",", ".")
                            ICMSSN202.pFCPST = Replace(Format(objTribDocItem.dICMSpFCPST * 100, "##0.00"), ",", ".")
                            ICMSSN202.vFCPST = Replace(Format(objTribDocItem.dICMSvFCPST, "fixed"), ",", ".")
                        End If

                        'Simples - Isenção ICMS e cobr. ICMS ST
                    Case 6
                        Dim ICMSSN202 As New TNFeInfNFeDetImpostoICMSICMSSN202
                        infNFeDetImpostoICMS.Item = ICMSSN202
                        '                                ICMS10.orig = objProduto.OrigemMercadoria
                        ICMSSN202.orig = objTribDocItem.iOrigemMercadoria
                        ICMSSN202.CSOSN = TNFeInfNFeDetImpostoICMSICMSSN202CSOSN.Item203
                        ICMSSN202.modBCST = TNFeInfNFeDetImpostoICMSICMSSN202ModBCST.Item4
                        If objTribDocItem.dICMSSubstPercMVA > 0 Then
                            ICMSSN202.pMVAST = Replace(Format(objTribDocItem.dICMSSubstPercMVA * 100, "##0.00"), ",", ".")
                        End If
                        If objTribDocItem.dICMSSubstPercRedBase > 0 Then
                            ICMSSN202.pRedBCST = Replace(Format(objTribDocItem.dICMSSubstPercRedBase * 100, "##0.00"), ",", ".")
                        End If
                        ICMSSN202.vBCST = Replace(Format(objTribDocItem.dICMSSubstBase, "fixed"), ",", ".")
                        ICMSSN202.pICMSST = Replace(Format((objTribDocItem.dICMSSubstAliquota - objTribDocItem.dICMSpFCPST) * 100, "##0.00"), ",", ".")
                        ICMSSN202.vICMSST = Replace(Format((objTribDocItem.dICMSSubstValor - objTribDocItem.dICMSvFCPST), "fixed"), ",", ".")

                        If objTribDocItem.dICMSvFCPST <> 0 Then
                            ICMSSN202.vBCFCPST = Replace(Format(objTribDocItem.dICMSvBCFCPST, "fixed"), ",", ".")
                            ICMSSN202.pFCPST = Replace(Format(objTribDocItem.dICMSpFCPST * 100, "##0.00"), ",", ".")
                            ICMSSN202.vFCPST = Replace(Format(objTribDocItem.dICMSvFCPST, "fixed"), ",", ".")
                        End If

                    Case 9
                        Dim ICMSSN500 As New TNFeInfNFeDetImpostoICMSICMSSN500
                        infNFeDetImpostoICMS.Item = ICMSSN500
                        ICMSSN500.orig = objTribDocItem.iOrigemMercadoria
                        ICMSSN500.CSOSN = TNFeInfNFeDetImpostoICMSICMSSN500CSOSN.Item500
                        ICMSSN500.vBCSTRet = Replace(Format(objTribDocItem.dICMSSTCobrAntBase, "fixed"), ",", ".")
                        ICMSSN500.vICMSSTRet = Replace(Format((objTribDocItem.dICMSSTCobrAntValor - objTribDocItem.dICMSvFCPSTRet), "fixed"), ",", ".")

                        ICMSSN500.pST = Replace(Format((objTribDocItem.dICMSSTCobrAntAliquota - objTribDocItem.dICMSpFCPSTRet) * 100, "##0.00"), ",", ".")

                        If objTribDocItem.dICMSvFCPSTRet <> 0 Then
                            ICMSSN500.vBCFCPSTRet = Replace(Format(objTribDocItem.dICMSvBCFCPSTRet, "fixed"), ",", ".")
                            ICMSSN500.pFCPSTRet = Replace(Format(objTribDocItem.dICMSpFCPSTRet * 100, "##0.00"), ",", ".")
                            ICMSSN500.vFCPSTRet = Replace(Format(objTribDocItem.dICMSvFCPSTRet, "fixed"), ",", ".")
                        End If

                    Case 10
                        Dim ICMSSN900 As New TNFeInfNFeDetImpostoICMSICMSSN900
                        infNFeDetImpostoICMS.Item = ICMSSN900
                        '                                ICMS10.orig = objProduto.OrigemMercadoria
                        ICMSSN900.orig = objTribDocItem.iOrigemMercadoria
                        ICMSSN900.CSOSN = TNFeInfNFeDetImpostoICMSICMSSN900CSOSN.Item900
                        ICMSSN900.modBC = TNFeInfNFeDetImpostoICMSICMSSN900ModBC.Item3
                        ICMSSN900.vBC = Replace(Format(objTribDocItem.dICMSBase, "fixed"), ",", ".")
                        If objTribDocItem.dICMSPercRedBase > 0 Then
                            ICMSSN900.pRedBC = Replace(Format(objTribDocItem.dICMSPercRedBase * 100, "##0.00"), ",", ".")
                        End If
                        ICMSSN900.pICMS = Replace(Format(objTribDocItem.dICMSAliquota * 100, "##0.00"), ",", ".")
                        ICMSSN900.vICMS = Replace(Format(objTribDocItem.dICMSValor, "fixed"), ",", ".")
                        ICMSSN900.modBCST = TNFeInfNFeDetImpostoICMSICMSSN900ModBCST.Item4
                        If objTribDocItem.dICMSSubstPercMVA > 0 Then
                            ICMSSN900.pMVAST = Replace(Format(objTribDocItem.dICMSSubstPercMVA * 100, "##0.00"), ",", ".")
                        End If
                        If objTribDocItem.dICMSSubstPercRedBase > 0 Then
                            ICMSSN900.pRedBCST = Replace(Format(objTribDocItem.dICMSSubstPercRedBase * 100, "##0.00"), ",", ".")
                        End If
                        ICMSSN900.vBCST = Replace(Format(objTribDocItem.dICMSSubstBase, "fixed"), ",", ".")
                        ICMSSN900.pICMSST = Replace(Format((objTribDocItem.dICMSSubstAliquota - objTribDocItem.dICMSpFCPST) * 100, "##0.00"), ",", ".")
                        ICMSSN900.vICMSST = Replace(Format((objTribDocItem.dICMSSubstValor - objTribDocItem.dICMSvFCPST), "fixed"), ",", ".")
                        ICMSSN900.pCredSN = Replace(Format(objTribDocItem.dICMSpCredSN * 100, "##0.00"), ",", ".")
                        ICMSSN900.vCredICMSSN = Replace(Format(objTribDocItem.dICMSvCredSN, "fixed"), ",", ".")

                        If objTribDocItem.dICMSvFCPST <> 0 Then
                            ICMSSN900.vBCFCPST = Replace(Format(objTribDocItem.dICMSvBCFCPST, "fixed"), ",", ".")
                            ICMSSN900.pFCPST = Replace(Format(objTribDocItem.dICMSpFCPST * 100, "##0.00"), ",", ".")
                            ICMSSN900.vFCPST = Replace(Format(objTribDocItem.dICMSvFCPST, "fixed"), ",", ".")
                        End If

                End Select
            End If

            Monta_NFiscal_Xml4d = SUCESSO

        Catch ex As Exception

            Monta_NFiscal_Xml4d = 1

            Call gobjApp.GravarLog(ex.Message)

        End Try

    End Function

    Private Function Monta_NFiscal_Xml4e(ByVal infNFeDetImposto As TNFeInfNFeDetImposto, ByVal objTribDocItem As GlobaisAdm.ClassTributacaoDocItem, ByRef dValorPIS As Double) As Long

        Try

            '***********  PIS ****************************

            If gobjApp.iDebug = 1 Then MsgBox("24")
            gobjApp.sErro = "24"
            gobjApp.sMsg1 = "vai iniciar a tributacao de PIS"


            Dim infNFeDetImpostoPIS As TNFeInfNFeDetImpostoPIS = New TNFeInfNFeDetImpostoPIS
            infNFeDetImposto.PIS = infNFeDetImpostoPIS

            dValorPIS = dValorPIS + objTribDocItem.dPISValor

            'lErro = PIS_CST(iCST, objTribDocItem)
            'If lErro <> SUCESSO Then Error 10000

            Select Case objTribDocItem.iPISTipo

                Case 1, 2
                    Dim PISAliq As New TNFeInfNFeDetImpostoPISPISAliq


                    infNFeDetImpostoPIS.Item = PISAliq

                    If objTribDocItem.iPISTipo = 1 Then
                        PISAliq.CST = TNFeInfNFeDetImpostoPISPISAliqCST.Item01
                    Else
                        PISAliq.CST = TNFeInfNFeDetImpostoPISPISAliqCST.Item02
                    End If

                    'lErro = PIS_Aliquota(dPISAliquota, objFilialEmpresa)
                    'If lErro <> SUCESSO Then Error 10001

                    PISAliq.pPIS = Replace(Format(objTribDocItem.dPISAliquota * 100, "##0.00"), ",", ".")
                    PISAliq.vPIS = Replace(Format(objTribDocItem.dPISValor, "fixed"), ",", ".")
                    PISAliq.vBC = Replace(Format(objTribDocItem.dPISBase, "fixed"), ",", ".")

                Case 3
                    Dim PISQtde As New TNFeInfNFeDetImpostoPISPISQtde
                    infNFeDetImpostoPIS.Item = PISQtde

                    PISQtde.CST = TNFeInfNFeDetImpostoPISPISQtdeCST.Item03
                    PISQtde.qBCProd = Replace(Format(objTribDocItem.dPISQtde, "#########0.0000"), ",", ".")
                    PISQtde.vAliqProd = Replace(Format(objTribDocItem.dPISAliquotaValor, "#########0.0000"), ",", ".")
                    PISQtde.vPIS = Replace(Format(objTribDocItem.dPISValor, "fixed"), ",", ".")

                Case 4
                    Dim PISNT As New TNFeInfNFeDetImpostoPISPISNT
                    infNFeDetImpostoPIS.Item = PISNT

                    PISNT.CST = TNFeInfNFeDetImpostoPISPISNTCST.Item04

                Case 6
                    Dim PISNT As New TNFeInfNFeDetImpostoPISPISNT
                    infNFeDetImpostoPIS.Item = PISNT

                    PISNT.CST = TNFeInfNFeDetImpostoPISPISNTCST.Item06

                Case 7
                    Dim PISNT As New TNFeInfNFeDetImpostoPISPISNT
                    infNFeDetImpostoPIS.Item = PISNT

                    PISNT.CST = TNFeInfNFeDetImpostoPISPISNTCST.Item07

                Case 8
                    Dim PISNT As New TNFeInfNFeDetImpostoPISPISNT
                    infNFeDetImpostoPIS.Item = PISNT

                    PISNT.CST = TNFeInfNFeDetImpostoPISPISNTCST.Item08

                Case 9
                    Dim PISNT As New TNFeInfNFeDetImpostoPISPISNT
                    infNFeDetImpostoPIS.Item = PISNT

                    PISNT.CST = TNFeInfNFeDetImpostoPISPISNTCST.Item09

                Case 49 To 56, 60 To 67, 70 To 75, 98, 99
                    Dim PISOutr As New TNFeInfNFeDetImpostoPISPISOutr
                    infNFeDetImpostoPIS.Item = PISOutr

                    Dim ItemsElementName1(1) As ItemsChoiceType1
                    Dim ItemsString1(1) As String

                    PISOutr.ItemsElementName = ItemsElementName1
                    PISOutr.Items = ItemsString1

                    If objTribDocItem.iPISTipoCalculo = TRIB_TIPO_CALCULO_PERCENTUAL Then

                        PISOutr.ItemsElementName(0) = ItemsChoiceType1.vBC
                        PISOutr.Items(0) = Replace(Format(objTribDocItem.dPISBase, "fixed"), ",", ".")
                        PISOutr.ItemsElementName(1) = ItemsChoiceType1.pPIS
                        PISOutr.Items(1) = Replace(Format(objTribDocItem.dPISAliquota * 100, "##0.00"), ",", ".")

                    Else

                        PISOutr.ItemsElementName(0) = ItemsChoiceType1.qBCProd
                        PISOutr.Items(0) = Replace(Format(objTribDocItem.dPISQtde, "#########0.0000"), ",", ".")
                        PISOutr.ItemsElementName(1) = ItemsChoiceType1.vAliqProd
                        PISOutr.Items(1) = Replace(Format(objTribDocItem.dPISAliquotaValor, "#########0.0000"), ",", ".")
                    End If

                    PISOutr.vPIS = Replace(Format(objTribDocItem.dPISValor, "fixed"), ",", ".")
                    PISOutr.CST = GetCode(Of TNFeInfNFeDetImpostoPISPISOutrCST)(CStr(objTribDocItem.iPISTipo))

                    If gobjApp.iDebug = 1 Then MsgBox("25")
                    gobjApp.sErro = "25"
                    gobjApp.sMsg1 = "vai iniciar a tributacao de PIS ST"


                    '***********  PIS ST ****************************

                    If objTribDocItem.dPISSTValor <> 0 Then

                        Dim PISST As TNFeInfNFeDetImpostoPISST = New TNFeInfNFeDetImpostoPISST
                        infNFeDetImposto.PISST = PISST

                        Dim ItemsElementName2(1) As ItemsChoiceType2
                        Dim ItemsString2(1) As String

                        PISST.ItemsElementName = ItemsElementName2
                        PISST.Items = ItemsString2

                        If objTribDocItem.iPISSTTipoCalculo = TRIB_TIPO_CALCULO_PERCENTUAL Then

                            PISST.ItemsElementName(0) = ItemsChoiceType2.vBC
                            PISST.Items(0) = Replace(Format(objTribDocItem.dPISSTBase, "fixed"), ",", ".")
                            PISST.ItemsElementName(1) = ItemsChoiceType2.pPIS
                            PISST.Items(1) = Replace(Format(objTribDocItem.dPISSTAliquota * 100, "##0.00"), ",", ".")

                        Else

                            PISST.ItemsElementName(0) = ItemsChoiceType1.qBCProd
                            PISST.Items(0) = Replace(Format(objTribDocItem.dPISSTQtde, "#########0.0000"), ",", ".")
                            PISST.ItemsElementName(1) = ItemsChoiceType1.vAliqProd
                            PISST.Items(1) = Replace(Format(objTribDocItem.dPISSTAliquotaValor, "#########0.0000"), ",", ".")

                        End If

                        PISST.vPIS = Replace(Format(objTribDocItem.dPISSTValor, "fixed"), ",", ".")

                    End If

            End Select

            Monta_NFiscal_Xml4e = SUCESSO

        Catch ex As Exception

            Monta_NFiscal_Xml4e = 1

            Call gobjApp.GravarLog(ex.Message)

        End Try

    End Function

    Private Function Monta_NFiscal_Xml4f(ByVal infNFeDetImposto As TNFeInfNFeDetImposto, ByVal objTribDocItem As GlobaisAdm.ClassTributacaoDocItem, ByRef dValorCOFINS As Double) As Long

        Try

            '***********  COFINS ****************************

            If gobjApp.iDebug = 1 Then MsgBox("26")
            gobjApp.sErro = "26"
            gobjApp.sMsg1 = "vai iniciar a tributacao de COFINS"


            Dim infNFeDetImpostoCOFINS As TNFeInfNFeDetImpostoCOFINS = New TNFeInfNFeDetImpostoCOFINS
            infNFeDetImposto.COFINS = infNFeDetImpostoCOFINS

            'lErro = COFINS_CST(iCST, objTribDocItem)
            'If lErro <> SUCESSO Then Error 10002

            dValorCOFINS = dValorCOFINS + objTribDocItem.dCOFINSValor

            Select Case objTribDocItem.iCOFINSTipo

                Case 1, 2
                    Dim COFINSAliq As New TNFeInfNFeDetImpostoCOFINSCOFINSAliq

                    infNFeDetImpostoCOFINS.Item = COFINSAliq


                    If objTribDocItem.iCOFINSTipo = 1 Then
                        COFINSAliq.CST = TNFeInfNFeDetImpostoCOFINSCOFINSAliqCST.Item01
                    Else
                        COFINSAliq.CST = TNFeInfNFeDetImpostoCOFINSCOFINSAliqCST.Item02
                    End If


                    'lErro = COFINS_Aliquota(dCOFINSAliquota, objFilialEmpresa)
                    'If lErro <> SUCESSO Then Error 10003

                    COFINSAliq.pCOFINS = Replace(Format(objTribDocItem.dCOFINSAliquota * 100, "##0.00"), ",", ".")
                    COFINSAliq.vCOFINS = Replace(Format(objTribDocItem.dCOFINSValor, "fixed"), ",", ".")
                    COFINSAliq.vBC = Replace(Format(objTribDocItem.dCOFINSBase, "fixed"), ",", ".")

                Case 3
                    Dim COFINSQtde As New TNFeInfNFeDetImpostoCOFINSCOFINSQtde
                    infNFeDetImpostoCOFINS.Item = COFINSQtde

                    COFINSQtde.CST = TNFeInfNFeDetImpostoCOFINSCOFINSQtdeCST.Item03
                    COFINSQtde.qBCProd = Replace(Format(objTribDocItem.dCOFINSQtde, "#########0.0000"), ",", ".")
                    COFINSQtde.vAliqProd = Replace(Format(objTribDocItem.dCOFINSAliquotaValor, "#########0.0000"), ",", ".")
                    COFINSQtde.vCOFINS = Replace(Format(objTribDocItem.dCOFINSValor, "fixed"), ",", ".")

                Case 4
                    Dim COFINSNT As New TNFeInfNFeDetImpostoCOFINSCOFINSNT
                    infNFeDetImpostoCOFINS.Item = COFINSNT

                    COFINSNT.CST = TNFeInfNFeDetImpostoCOFINSCOFINSNTCST.Item04

                Case 6
                    Dim COFINSNT As New TNFeInfNFeDetImpostoCOFINSCOFINSNT
                    infNFeDetImpostoCOFINS.Item = COFINSNT

                    COFINSNT.CST = TNFeInfNFeDetImpostoCOFINSCOFINSNTCST.Item06

                Case 7
                    Dim COFINSNT As New TNFeInfNFeDetImpostoCOFINSCOFINSNT
                    infNFeDetImpostoCOFINS.Item = COFINSNT

                    COFINSNT.CST = TNFeInfNFeDetImpostoCOFINSCOFINSNTCST.Item07

                Case 8
                    Dim COFINSNT As New TNFeInfNFeDetImpostoCOFINSCOFINSNT
                    infNFeDetImpostoCOFINS.Item = COFINSNT

                    COFINSNT.CST = TNFeInfNFeDetImpostoCOFINSCOFINSNTCST.Item08

                Case 9
                    Dim COFINSNT As New TNFeInfNFeDetImpostoCOFINSCOFINSNT
                    infNFeDetImpostoCOFINS.Item = COFINSNT

                    COFINSNT.CST = TNFeInfNFeDetImpostoCOFINSCOFINSNTCST.Item09

                Case 49 To 56, 60 To 67, 70 To 75, 98, 99
                    Dim COFINSOutr As New TNFeInfNFeDetImpostoCOFINSCOFINSOutr
                    infNFeDetImpostoCOFINS.Item = COFINSOutr

                    Dim ItemsElementName3(1) As ItemsChoiceType3
                    Dim ItemsString3(1) As String

                    COFINSOutr.ItemsElementName = ItemsElementName3
                    COFINSOutr.Items = ItemsString3

                    If objTribDocItem.iCOFINSTipoCalculo = TRIB_TIPO_CALCULO_PERCENTUAL Then

                        COFINSOutr.ItemsElementName(0) = ItemsChoiceType3.vBC
                        COFINSOutr.Items(0) = Replace(Format(objTribDocItem.dCOFINSBase, "fixed"), ",", ".")
                        COFINSOutr.ItemsElementName(1) = ItemsChoiceType3.pCOFINS
                        COFINSOutr.Items(1) = Replace(Format(objTribDocItem.dCOFINSAliquota * 100, "##0.00"), ",", ".")

                    Else

                        COFINSOutr.ItemsElementName(0) = ItemsChoiceType3.qBCProd
                        COFINSOutr.Items(0) = Replace(Format(objTribDocItem.dCOFINSQtde, "#########0.0000"), ",", ".")
                        COFINSOutr.ItemsElementName(1) = ItemsChoiceType3.vAliqProd
                        COFINSOutr.Items(1) = Replace(Format(objTribDocItem.dCOFINSAliquotaValor, "#########0.0000"), ",", ".")

                    End If

                    COFINSOutr.vCOFINS = Replace(Format(objTribDocItem.dCOFINSValor, "fixed"), ",", ".")
                    COFINSOutr.CST = GetCode(Of TNFeInfNFeDetImpostoCOFINSCOFINSOutrCST)(CStr(objTribDocItem.iCOFINSTipo))

                    If gobjApp.iDebug = 1 Then MsgBox("27")
                    gobjApp.sErro = "27"
                    gobjApp.sMsg1 = "vai iniciar a tributacao de COFINS ST"

                    '***********  COFINS ST****************************

                    If objTribDocItem.dCOFINSSTValor <> 0 Then
                        Dim COFINSST As TNFeInfNFeDetImpostoCOFINSST = New TNFeInfNFeDetImpostoCOFINSST
                        infNFeDetImposto.COFINSST = COFINSST

                        Dim ItemsElementName4(1) As ItemsChoiceType4
                        Dim ItemsString4(1) As String

                        COFINSST.ItemsElementName = ItemsElementName4
                        COFINSST.Items = ItemsString4


                        If objTribDocItem.iCOFINSSTTipoCalculo = TRIB_TIPO_CALCULO_PERCENTUAL Then

                            COFINSST.ItemsElementName(0) = ItemsChoiceType4.vBC
                            COFINSST.Items(0) = Replace(Format(objTribDocItem.dCOFINSSTBase, "fixed"), ",", ".")
                            COFINSST.ItemsElementName(1) = ItemsChoiceType4.pCOFINS
                            COFINSST.Items(1) = Replace(Format(objTribDocItem.dCOFINSSTAliquota * 100, "##0.00"), ",", ".")

                        Else


                            COFINSST.ItemsElementName(0) = ItemsChoiceType4.qBCProd
                            COFINSST.Items(0) = Replace(Format(objTribDocItem.dCOFINSSTQtde, "#########0.0000"), ",", ".")
                            COFINSST.ItemsElementName(1) = ItemsChoiceType4.vAliqProd
                            COFINSST.Items(1) = Replace(Format(objTribDocItem.dCOFINSSTAliquotaValor, "#########0.0000"), ",", ".")

                        End If

                        COFINSST.vCOFINS = Replace(Format(objTribDocItem.dCOFINSSTValor, "fixed"), ",", ".")

                    End If

            End Select

            Monta_NFiscal_Xml4f = SUCESSO

        Catch ex As Exception

            Monta_NFiscal_Xml4f = 1

            Call gobjApp.GravarLog(ex.Message)

        End Try

    End Function

    Private Function Monta_NFiscal_Xml4g(ByVal infNFeDetImposto As TNFeInfNFeDetImposto, ByVal objTribDocItem As GlobaisAdm.ClassTributacaoDocItem, ByRef dValorServPIS As Double, ByRef dValorServCOFINS As Double) As Long

        Try

            ''***************** ISS **************************************
            'resTiposDeTributacaoMovto = db1.ExecuteQuery(Of TiposDeTributacaoMovto) _
            '("SELECT * FROM TiposDeTributacaoMovto WHERE  Tipo = {0} ", objTribDocItem.TipoTributacao)

            'objTiposDeTributacaoMovto = resTiposDeTributacaoMovto(0)

            dValorServPIS = dValorServPIS + objTribDocItem.dPISValor
            dValorServCOFINS = dValorServCOFINS + objTribDocItem.dCOFINSValor

            Dim infNFeDetImpostoISSQN As TNFeInfNFeDetImpostoISSQN = New TNFeInfNFeDetImpostoISSQN
            infNFeDetImposto.Items(0) = infNFeDetImpostoISSQN

            infNFeDetImpostoISSQN.vBC = Replace(Format(objTribDocItem.dISSBase, "fixed"), ",", ".")
            infNFeDetImpostoISSQN.vAliq = Replace(Format(objTribDocItem.dISSAliquota * 100, "##0.00"), ",", ".")
            infNFeDetImpostoISSQN.vISSQN = Replace(Format(objTribDocItem.dISSValor, "fixed"), ",", ".")
            infNFeDetImpostoISSQN.cMunFG = gobjVenda.objNFeInfo.cMun

            ''classificacao do servico conforme tabela da lei complementar 116 de 2003 (LC 116/03)
            infNFeDetImpostoISSQN.cListServ = objTribDocItem.sCListServNFe

            Monta_NFiscal_Xml4g = SUCESSO

        Catch ex As Exception

            Monta_NFiscal_Xml4g = 1

            Call gobjApp.GravarLog(ex.Message)

        End Try

    End Function

    Private Function Monta_NFiscal_Xml(ByRef XMLStringNFes As String, ByVal a5 As TNFe) As Long

        Dim objEndereco As GlobaisAdm.ClassEndereco
        Dim lErro As Long
        Dim XMLStream As MemoryStream = New MemoryStream(10000)
        Dim XMLString As String
        Dim iPos As Integer
        Dim iPos2 As Integer
        Dim iPos3 As String
        Dim XMLStreamDados As MemoryStream = New MemoryStream(10000)
        Dim sArquivo As String

        Try

            gobjApp.sErro = "3.02"
            gobjApp.sMsg1 = "vai testar a serie para ver se é scan"

            objEndereco = gobjVenda.objNFeInfo.objEndereco

            If CInt(sSerie) >= 890 And CInt(sSerie) <= 899 Then
                Throw New System.Exception("Série 890-899 de uso exclusivo para emissão de NF-e avulsa pelo contribuinte com seu certificado digital, através do site do Fisco.")
            End If

            If gobjApp.iDebug = 1 Then MsgBox("3.1")

            gobjApp.sErro = "3.1"
            gobjApp.sMsg1 = "vai fazer INSERT NFeFedLoteLog"

            gobjApp.GravarLog("Iniciando o processamento da Nota Fiscal = " & CStr(lNumNotaFiscal) & " Série = " & sSerie)

            gobjVenda.objCupomFiscal.dtDataEmissao = Now.Date
            gobjVenda.objCupomFiscal.dHoraEmissao = TimeOfDay.ToOADate

            'System.Windows.Forms.Application.DoEvents()

            If gobjApp.iDebug = 1 Then MsgBox("4")

            lErro = Monta_NFiscal_Xml1(a5)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            lErro = Monta_NFiscal_Xml2(a5)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            lErro = Monta_NFiscal_Xml3(a5)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            'v2.00

            lErro = Monta_NFiscal_Xml4(a5)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            If gobjApp.iDebug = 1 Then MsgBox("30")
            gobjApp.sErro = "30"
            gobjApp.sMsg1 = "vai tratar os dados da transportadora"

            '***********  transportadora ****************************

            Dim infNFeTransp As TNFeInfNFeTransp = New TNFeInfNFeTransp
            a5.infNFe.transp = infNFeTransp

            'Dim infNFeTranspTransporta As TNFeInfNFeTranspTransporta = New TNFeInfNFeTranspTransporta
            'a5.infNFe.transp.transporta = infNFeTranspTransporta

            a5.infNFe.transp.modFrete = TNFeInfNFeTranspModFrete.Item9

            'a5.infNFe.ide.indPag = TNFeInfNFeIdeIndPag.Item0

            Dim dValorDesconto As Double, dValorTotalTitulo As Double

            dValorDesconto = CDbl(Format(gobjVenda.objCupomFiscal.dValorDesconto, "fixed"))
            dValorTotalTitulo = CDbl(Format(gobjVenda.objCupomFiscal.dValorTotal, "fixed"))

            If sModelo = "NFCe" Then

                Dim iIndice As Integer
                Dim apag(50) As TNFeInfNFePagDetPag
                a5.infNFe.pag = New TNFeInfNFePag
                a5.infNFe.pag.detPag = apag

                Dim dDinheiro As Double, dOutros As Double
                dDinheiro = 0
                dOutros = 0
                For Each objMovCaixa In gobjVenda.colMovimentosCaixa

                    Select Case objMovCaixa.iTipo

                        Case MOVIMENTOCAIXA_TROCO_DINHEIRO
                            dDinheiro = Arredonda_Moeda(dDinheiro - Math.Abs(objMovCaixa.dValor))

                        Case MOVIMENTOCAIXA_RECEB_DINHEIRO
                            dDinheiro = Arredonda_Moeda(dDinheiro + objMovCaixa.dValor)

                        Case MOVIMENTOCAIXA_TROCO_CONTRAVALE, MOVIMENTOCAIXA_TROCO_VALE
                            dOutros = Arredonda_Moeda(dOutros + objMovCaixa.dValor)

                        Case Else
                            If objMovCaixa.iCodigoCFe = 99 Then dOutros = Arredonda_Moeda(dOutros + objMovCaixa.dValor)

                    End Select

                Next

                iIndice = 0

                If dDinheiro > 0 Then

                    a5.infNFe.pag.detPag(iIndice) = New TNFeInfNFePagDetPag

                    a5.infNFe.pag.detPag(iIndice).tPag = TNFeInfNFePagDetPagTPag.Item01
                    a5.infNFe.pag.detPag(iIndice).vPag = Replace(Format(dDinheiro, "fixed"), ",", ".")
                    iIndice = iIndice + 1

                End If

                For Each objMovCaixa In gobjVenda.colMovimentosCaixa

                    Select Case objMovCaixa.iTipo

                        Case MOVIMENTOCAIXA_RECEB_DINHEIRO, MOVIMENTOCAIXA_TROCO_DINHEIRO, MOVIMENTOCAIXA_TROCO_CONTRAVALE, MOVIMENTOCAIXA_TROCO_VALE

                        Case Else

                            If objMovCaixa.iCodigoCFe <> 99 Then

                                a5.infNFe.pag.detPag(iIndice) = New TNFeInfNFePagDetPag

                                a5.infNFe.pag.detPag(iIndice).tPag = GetCode(Of TNFeInfNFePagDetPagTPag)(Format(objMovCaixa.iCodigoCFe, "00"))
                                a5.infNFe.pag.detPag(iIndice).vPag = Replace(Format(objMovCaixa.dValor, "fixed"), ",", ".")

                                '??? cAdmC

                                If objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAOCREDITO Or objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAODEBITO Then

                                    Dim objCard As New TNFeInfNFePagDetPagCard

                                    objCard.tpIntegra = GetCode(Of TNFeInfNFePagDetPagCardTpIntegra)(IIf(objMovCaixa.iTipoCartao = TIPO_TEF, "1", "2"))
                                    'objCard.tpIntegraSpecified = True

                                    Dim sAutorizacao As String
                                    If Len(Trim(objMovCaixa.sAutorizacao)) <> 0 Then

                                        sAutorizacao = Trim(objMovCaixa.sAutorizacao)

                                        While Left(sAutorizacao, 1) = "0"

                                            sAutorizacao = Mid(sAutorizacao, 2)

                                        End While

                                        If Len(Trim(sAutorizacao)) <> 0 Then objCard.cAut = sAutorizacao

                                    End If

                                    If objMovCaixa.iBandeiraCartao <> 0 Then

                                        objCard.tBand = GetCode(Of TNFeInfNFePagDetPagCardTBand)(Format(objMovCaixa.iBandeiraCartao, "00"))
                                        objCard.tBandSpecified = True

                                    End If

                                    a5.infNFe.pag.detPag(iIndice).card = objCard

                                End If

                                iIndice = iIndice + 1

                            End If

                    End Select

                Next

                If dOutros <> 0 Then

                    a5.infNFe.pag.detPag(iIndice) = New TNFeInfNFePagDetPag

                    a5.infNFe.pag.detPag(iIndice).tPag = GetCode(Of TNFeInfNFePagDetPagTPag)("99")
                    a5.infNFe.pag.detPag(iIndice).vPag = Replace(Format(dOutros, "fixed"), ",", ".")
                    iIndice = iIndice + 1

                End If

            Else

                If gobjApp.iDebug = 1 Then MsgBox("33")
                gobjApp.sErro = "33"
                gobjApp.sMsg1 = "vai tratar os dados de cobrança"

                '***********  cobranca ****************************
                Dim infNFeCobr As TNFeInfNFeCobr = New TNFeInfNFeCobr
                a5.infNFe.cobr = infNFeCobr

                Dim infNFeCobrFat As TNFeInfNFeCobrFat = New TNFeInfNFeCobrFat
                a5.infNFe.cobr.fat = infNFeCobrFat

                a5.infNFe.cobr.fat.nFat = lNumNotaFiscal
                a5.infNFe.cobr.fat.vOrig = Replace(Format(dValorTotalTitulo, "fixed"), ",", ".")
                If dValorDesconto <> 0 Then a5.infNFe.cobr.fat.vDesc = Replace(Format(dValorDesconto, "fixed"), ",", ".")
                a5.infNFe.cobr.fat.vLiq = Replace(Format(dValorTotalTitulo - dValorDesconto, "fixed"), ",", ".")

                Dim Dup(50) As TNFeInfNFeCobrDup

                a5.infNFe.cobr.dup = Dup

                Dim infNFeCobrDup As TNFeInfNFeCobrDup = New TNFeInfNFeCobrDup
                a5.infNFe.cobr.dup(0) = infNFeCobrDup

                a5.infNFe.cobr.dup(0).nDup = lNumNotaFiscal
                a5.infNFe.cobr.dup(0).dVenc = Format(gobjVenda.objCupomFiscal.dtDataEmissao, "yyyy-MM-dd")
                a5.infNFe.cobr.dup(0).vDup = Replace(Format(dValorTotalTitulo - dValorDesconto, "fixed"), ",", ".")

            End If

            If gobjApp.iDebug = 1 Then MsgBox("34")
            gobjApp.sErro = "34"
            gobjApp.sMsg1 = "vai montar a chave da nota"

            a5.infNFe.Id = gobjVenda.objNFeInfo.iUFCodIBGE & Format(gobjVenda.objCupomFiscal.dtDataEmissao, "yyMM") & a5.infNFe.emit.Item
            a5.infNFe.Id = a5.infNFe.Id & GetXmlAttrNameFromEnumValue(Of TMod)(a5.infNFe.ide.mod) & Format(CInt(a5.infNFe.ide.serie), "000")
            a5.infNFe.Id = a5.infNFe.Id & Format(CLng(a5.infNFe.ide.nNF), "000000000") & GetXmlAttrNameFromEnumValue(Of TNFeInfNFeIdeTpEmis)(a5.infNFe.ide.tpEmis)
            a5.infNFe.Id = a5.infNFe.Id & Format(CLng(a5.infNFe.ide.cNF), "00000000")

            If gobjApp.iDebug = 1 Then MsgBox("35")
            gobjApp.sErro = "35"
            gobjApp.sMsg1 = "vai calcular o DV da chave"

            Dim iDigito As Integer

            CalculaDV_Modulo11(a5.infNFe.Id, iDigito)

            If gobjApp.iDebug = 1 Then MsgBox("36")

            gobjApp.sErro = "36"
            gobjApp.sMsg1 = "vai serializar os dados da nota"


            a5.infNFe.Id = "NFe" & a5.infNFe.Id & iDigito
            a5.infNFe.ide.cDV = iDigito

            If sModelo = "NFe" Then

                a5.infNFe.infAdic = New TNFeInfNFeInfAdic
                a5.infNFe.infAdic.infAdFisco = "MD-5:" & gobjVenda.objNFeInfo.sMD5PAFECF

            Else

                If Len(Trim(gobjVenda.objCupomFiscal.sNFCeMensagem)) <> 0 Then

                    a5.infNFe.infAdic = New TNFeInfNFeInfAdic
                    a5.infNFe.infAdic.infCpl = gobjVenda.objCupomFiscal.sNFCeMensagem

                End If

                If gobjApp.iNFeAmbiente = NFE_AMBIENTE_HOMOLOGACAO Or gobjVenda.objCupomFiscal.dtDataEmissao >= #12/1/2015# Then

                    a5.infNFeSupl = New TNFeInfNFeSupl
                    a5.infNFeSupl.qrCode = QRCODE_PROVISORIO

                    a5.infNFeSupl.urlChave = url_chave2(Mid(a5.infNFe.Id, 4), a5.infNFe.ide.tpAmb)

                End If

            End If

            Dim AD As AssinaturaDigital = New AssinaturaDigital

            Dim mySerializer As New XmlSerializer(GetType(TNFe))

            XMLStream = New MemoryStream(10000)

            mySerializer.Serialize(XMLStream, a5)

            Dim xm As Byte()
            xm = XMLStream.ToArray

            XMLString = System.Text.Encoding.UTF8.GetString(xm)

            XMLString = Replace(XMLString, "<ICMS100>", "<ICMS40>")
            XMLString = Replace(XMLString, "</ICMS100>", "</ICMS40>")

            '***********************************



            iPos = InStr(XMLString, "xmlns:xsi")
            iPos2 = InStr(XMLString, """>")

            XMLString = Mid(XMLString, 1, iPos - 1) & Mid(XMLString, iPos2 + 1)

            'retirado em 31/03/2010 pois estava dando erro no xml
            iPos3 = InStr(XMLString, "<NFe >")

            XMLString = Mid(XMLString, 1, iPos3 + 4) & "xmlns = ""http://www.portalfiscal.inf.br/nfe""" & Mid(XMLString, iPos3 + 5)

            '****************************************


            iPos = InStr(XMLString, "<infNFe")

            If iPos <> 0 Then

                Dim iPos1 As Integer

                iPos1 = InStr(Mid(XMLString, iPos), "xmlns=""http://www.portalfiscal.inf.br/nfe""")

                If iPos1 <> 0 Then

                    XMLString = Mid(XMLString, 1, iPos + iPos1 - 2) & Mid(XMLString, iPos + iPos1 + 41)

                End If

            End If


            If gobjApp.iDebug = 1 Then MsgBox("37")
            gobjApp.sErro = "37"
            gobjApp.sMsg1 = "vai assinar a nota"


            lErro = AD.Assinar(XMLString, "infNFe", gobjApp.cert, gobjApp.iDebug)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            'Dim xMlD As XmlDocument

            'xMlD = AD.XMLDocAssinado()

            Dim xString As String
            xString = AD.XMLStringAssinado

            If sModelo = "NFCe" And InStr(xString, QRCODE_PROVISORIO) <> 0 Then

                'calcular o qrcode e substituir
                Dim sDigVal As String = "", sNFCeQRCode As String = ""
                Dim iPos1Aux As Integer, iPos2Aux As Integer
                iPos1Aux = InStr(xString, "<DigestValue>") + Len("<DigestValue>")
                iPos2Aux = InStr(xString, "</DigestValue>")
                If iPos2Aux > iPos1Aux And iPos1Aux <> 0 Then

                    sDigVal = Mid(xString, iPos1Aux, iPos2Aux - iPos1Aux)

                    If a5.infNFe.ide.dhCont Is Nothing Then

                        'TODO: Trocada chamada do QRCode Online 2
                        gobjApp.GravarLog("Monta_NFiscal_Xml NFCE_Gera_QRCode2_Online",, False)

                        'sNFCeQRCode = NFCE_Gera_QRCode(Mid(a5.infNFe.Id, 4), "100", GetXmlAttrNameFromEnumValue(Of TAmb)(a5.infNFe.ide.tpAmb), scDest, sdhEmi, svNF, svICMS, sDigVal, gobjVenda.objNFeInfo.sidNFCECSC, gobjVenda.objNFeInfo.sNFCECSC)
                        sNFCeQRCode = NFCE_Gera_QRCode2_Online(gobjApp, Mid(a5.infNFe.Id, 4), GetXmlAttrNameFromEnumValue(Of TAmb)(a5.infNFe.ide.tpAmb), gobjVenda.objNFeInfo.sidNFCECSC, gobjVenda.objNFeInfo.sNFCECSC)

                    Else

                        'TODO: Trocada chamada do QRCode Offline 3
                        gobjApp.GravarLog("Monta_NFiscal_Xml NFCE_Gera_QRCode2_Offline",, False)

                        'sNFCeQRCode = NFCE_Gera_QRCode(Mid(a5.infNFe.Id, 4), "100", GetXmlAttrNameFromEnumValue(Of TAmb)(a5.infNFe.ide.tpAmb), scDest, sdhEmi, svNF, svICMS, sDigVal, gobjVenda.objNFeInfo.sidNFCECSC, gobjVenda.objNFeInfo.sNFCECSC)
                        sNFCeQRCode = NFCE_Gera_QRCode2_Offline(gobjApp, Mid(a5.infNFe.Id, 4), GetXmlAttrNameFromEnumValue(Of TAmb)(a5.infNFe.ide.tpAmb), a5.infNFe.ide.dhEmi, a5.infNFe.total.ICMSTot.vNF, sDigVal, gobjVenda.objNFeInfo.sidNFCECSC, gobjVenda.objNFeInfo.sNFCECSC)

                    End If

                    sNFCeQRCode = "<![CDATA[" & sNFCeQRCode & "]]>"
                    xString = Replace(xString, QRCODE_PROVISORIO, sNFCeQRCode)

                End If

            End If

            XMLStringNFes = XMLStringNFes & Mid(xString, 22) & " "

            '            XMLStringNFes = Replace(XMLStringNFes, "TNFe", "NFe")

            '****************  salva o arquivo 

            XMLStreamDados = New MemoryStream(10000)

            Dim xDados1 As Byte()

            xDados1 = System.Text.Encoding.UTF8.GetBytes(Mid(xString, 22))

            XMLStreamDados.Write(xDados1, 0, xDados1.Length)

            Dim DocDados1 As XmlDocument = New XmlDocument

            XMLStreamDados.Position = 0
            DocDados1.Load(XMLStreamDados)



            sArquivo = gobjApp.sDirXml & Mid(a5.infNFe.Id, 4) & "-pre.xml"

            Dim writer As New XmlTextWriter(sArquivo, Nothing)

            writer.Formatting = Formatting.None
            DocDados1.WriteTo(writer)

            writer.Close()

            gobjVenda.objCupomFiscal.sNFeArqXmlPre = sArquivo

            Monta_NFiscal_Xml = SUCESSO

        Catch ex As Exception

            Monta_NFiscal_Xml = 1

            Dim sMsg2 As String

            If ex.InnerException Is Nothing Then
                sMsg2 = ""
            Else
                sMsg2 = " - " & ex.InnerException.Message
            End If

            gobjApp.GravarLog("ERRO - " & ex.Message & sMsg2 & IIf(lNumNotaFiscal <> 0, "Serie = " & sSerie & " Nota Fiscal = " & lNumNotaFiscal, ""))

            'System.Windows.Forms.Application.DoEvents()

        End Try

    End Function

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Public Function Consulta_NFe(ByVal sChaveNFe As String, ByRef bAutorizou As Boolean) As Long

        Dim NfeConsulta As New nfeconsulta2.NFeConsultaProtocolo4

        Dim XMLStream1 As MemoryStream = New MemoryStream(10000)
        Dim XMLStreamCabec As MemoryStream = New MemoryStream(10000)
        Dim XMLStreamRet As MemoryStream = New MemoryStream(10000)

        Dim XMLString1 As String

        Dim xRet As Byte()
        'Dim mySerializercabec As New XmlSerializer(GetType(nfeconsulta2.nfeCabecMsg))

        Dim colNumIntNFiscal As Collection = New Collection

        ' Dim objCabecMsg As nfeconsulta2.nfeCabecMsg = New nfeconsulta2.nfeCabecMsg
        'Dim XMLStringCabec As String

        Dim iPos As Integer
        Dim lErro As Long


        'Dim xmcabec As Byte()

        Dim colNFiscal As Collection = New Collection

        Dim xmlNode1 As XmlNode
        Dim iProcessado As Integer

        Try

            gobjApp.GravarLog("Iniciando a consulta da nota fiscal com chave " & sChaveNFe)

            gobjApp.sErro = "13"
            gobjApp.sMsg1 = "vai montar o cabecalho"

            'objCabecMsg.versaoDados = NFE_VERSAO_XML

            'mySerializercabec = New XmlSerializer(GetType(nfeconsulta2.nfeCabecMsg))

            'XMLStreamCabec = New MemoryStream(10000)

            'mySerializercabec.Serialize(XMLStreamCabec, objCabecMsg)

            'XMLStreamCabec.Position = 0

            'xmcabec = XMLStreamCabec.ToArray

            'XMLStringCabec = System.Text.Encoding.UTF8.GetString(xmcabec)

            'XMLStringCabec = Mid(XMLStringCabec, 1, 19) & " encoding=""UTF-8"" " & Mid(XMLStringCabec, 20)

            gobjApp.sErro = "19"
            gobjApp.sMsg1 = "vai montar a estrutura TConsSitNFe"

            Dim objconsSitNFe As TConsSitNFe = New TConsSitNFe

            objconsSitNFe.chNFe = sChaveNFe

            If gobjApp.iNFeAmbiente = NFE_AMBIENTE_HOMOLOGACAO Then
                objconsSitNFe.tpAmb = TAmb.Item2
            Else
                objconsSitNFe.tpAmb = TAmb.Item1
            End If

            objconsSitNFe.versao = TVerConsSitNFe.Item400
            objconsSitNFe.xServ = TConsSitNFeXServ.CONSULTAR

            gobjApp.sErro = "20"
            gobjApp.sMsg1 = "vai serializar TConsSitNFe"

            Dim mySerializerw As New XmlSerializer(GetType(TConsSitNFe))

            Dim XMLStream2 = New MemoryStream(10000)
            mySerializerw.Serialize(XMLStream2, objconsSitNFe)

            Dim xm2 As Byte()
            xm2 = XMLStream2.ToArray

            XMLString1 = System.Text.Encoding.UTF8.GetString(xm2)

            XMLString1 = Mid(XMLString1, 1, 19) & " encoding=""UTF-8"" " & Mid(XMLString1, 20)

            iPos = InStr(XMLString1, "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema""")

            If iPos <> 0 Then

                XMLString1 = Mid(XMLString1, 1, iPos - 1) & Mid(XMLString1, iPos + 99)

            End If

            Dim XMLStringRetConsultaNF As String

            Dim DocDados1 As New XmlDocument

            Call Salva_Arquivo(DocDados1, XMLString1)

            'Dim NfeCabec As New nfeconsulta2.nfeCabecMsg
            'NfeCabec.cUF = CStr(gobjApp.iUFCodIBGE)
            'NfeCabec.versaoDados = NFE_VERSAO_XML

            'NfeConsulta.nfeCabecMsgValue = NfeCabec

            Dim sURL As String
            sURL = ""
            Call WS_Obter_URL(sURL, gobjApp.iNFeAmbiente = NFE_AMBIENTE_HOMOLOGACAO, gobjApp.sSiglaEstado, "NfeConsultaProtocolo", sModelo)
            NfeConsulta.Url = sURL

            NfeConsulta.ClientCertificates.Add(gobjApp.cert)
            xmlNode1 = NfeConsulta.nfeConsultaNF(DocDados1)

            XMLStringRetConsultaNF = xmlNode1.OuterXml

            If gobjApp.iDebug = 1 Then MsgBox(XMLStringRetConsultaNF)

            xRet = System.Text.Encoding.UTF8.GetBytes(XMLStringRetConsultaNF)

            XMLStreamRet = New MemoryStream(10000)
            XMLStreamRet.Write(xRet, 0, xRet.Length)

            Dim mySerializerRetConsultaNF As New XmlSerializer(GetType(TRetConsSitNFe))

            Dim objRetConsSitNFe As TRetConsSitNFe = New TRetConsSitNFe

            XMLStreamRet.Position = 0

            objRetConsSitNFe = mySerializerRetConsultaNF.Deserialize(XMLStreamRet)

            gobjApp.sErro = "21"
            gobjApp.sMsg1 = "trata a nota fiscal consultada"

            iProcessado = 0

            If Not objRetConsSitNFe.protNFe Is Nothing Then

                lErro = protNFe_Processa(objRetConsSitNFe.protNFe, bAutorizou)
                If lErro <> SUCESSO Then Throw New System.Exception("")

            End If

            Consulta_NFe = SUCESSO

        Catch ex As Exception

            Consulta_NFe = 1

            Call gobjApp.GravarLog(ex.Message)

        Finally

        End Try

    End Function

    Private Function Produto_Trata_EAN(ByVal objTrib As GlobaisAdm.ClassTributacaoDocItem) As Long

        Dim lErro As Long
        Dim iValidacaoEAN As Integer
        'Dim resCRFatConfig As IEnumerable(Of CRFatConfig)
        'Dim objCRFatConfig As CRFatConfig
        Dim dtValidacaoEAN As Date

        Try

            'resCRFatConfig = gobjApp.dbDadosNfe.ExecuteQuery(Of CRFatConfig) _
            '("SELECT * FROM CRFatConfig WHERE Codigo = 'NFE_VALIDACAO_EAN'")

            'For Each objCRFatConfig In resCRFatConfig
            iValidacaoEAN = 1 'CInt(objCRFatConfig.Conteudo)
            '    Exit For
            'Next

            'resCRFatConfig = gobjApp.dbDadosNfe.ExecuteQuery(Of CRFatConfig) _
            '("SELECT * FROM CRFatConfig WHERE Codigo = 'NFE_VALIDACAO_EAN_A_PARTIR_DE'")

            'For Each objCRFatConfig In resCRFatConfig
            dtValidacaoEAN = #1/1/2018# 'CDate(objCRFatConfig.Conteudo)
            'Exit For
            'Next

            If dtValidacaoEAN <= Now.Date Then
                If Left(objTrib.sEAN, 1) = "2" Then objTrib.sEAN = "" 'Não enviar código de Barras de balança
                If Left(objTrib.sEANTrib, 1) = "2" Then objTrib.sEANTrib = "" 'Não enviar código de Barras de balança
            End If

            Select Case iValidacaoEAN

                Case 1 'Valida 

                    If dtValidacaoEAN <= Now.Date Then
                        lErro = Valida_EAN(objTrib.sEAN)
                        If lErro <> SUCESSO Then objTrib.sEAN = ""

                        lErro = Valida_EAN(objTrib.sEANTrib)
                        If lErro <> SUCESSO Then objTrib.sEANTrib = ""

                        'Para produtos que não possuem código de barras com GTIN , deve ser informado o literal “SEM GTIN”;
                        If objTrib.sEAN = "" Then objTrib.sEAN = "SEM GTIN"
                        If objTrib.sEANTrib = "" Then objTrib.sEANTrib = objTrib.sEAN

                    End If

                Case 2 'Não Envia nada
                    objTrib.sEAN = ""
                    objTrib.sEANTrib = ""

                Case Else
                    'Vai o que tiver

            End Select

            Produto_Trata_EAN = SUCESSO

        Catch ex As Exception
            Produto_Trata_EAN = 1
        End Try
    End Function

    Private Function Valida_EAN(ByVal sEAN As String) As Long
        Dim lErro As Long = 0
        Try
            Dim intTotalSoma As Integer
            Dim intDv As Integer
            Dim I As Integer
            Dim iNumChar As Integer
            Dim iMult As Integer

            iNumChar = Len(Trim(sEAN))
            sEAN = Trim(sEAN)
            iMult = 3

            If iNumChar <> 0 Then

                If Not (iNumChar = 8 Or iNumChar = 12 Or iNumChar = 13 Or iNumChar = 14) Or Not IsNumeric(sEAN) Then Error 6015

                '0,8,12,13, 14
                'Preencher com o código GTIN-8, GTIN-12, GTIN-13
                'ou GTIN-14 (antigos códigos EAN, UPC e DUN-14).
                'Para produtos que não possuem código de
                'barras com GTIN, deve ser informado o literal
                '“SEM GTIN”;
                'Nos demais casos, preencher com GTIN contido na
                'embalagem com código de barras

                intTotalSoma = 0
                intDv = 0

                For I = iNumChar - 1 To 1 Step -1
                    intTotalSoma = intTotalSoma + CInt(Mid(sEAN, I, 1)) * iMult
                    If iMult = 3 Then
                        iMult = 1
                    Else
                        iMult = 3
                    End If
                Next

                Do While intTotalSoma Mod 10 <> 0
                    intDv = intDv + 1
                    intTotalSoma = intTotalSoma + 1
                Loop

                If Right(sEAN, 1) <> CStr(intDv) Then Error 6016

            End If

            Valida_EAN = SUCESSO

        Catch ex As Exception
            Valida_EAN = Err.Number
        End Try

    End Function

End Class