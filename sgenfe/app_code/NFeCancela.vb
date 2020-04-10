Imports System
Imports System.Xml.Serialization
Imports System.IO
Imports System.Text
Imports System.Security.Cryptography.Xml
Imports System.Security.Cryptography.X509Certificates
Imports System.Xml
Imports System.Xml.XPath
Imports System.Xml.Schema
Imports Microsoft.Win32
Imports sgenfe4.NFeXsd
Imports sgenfe4.CancXsd
Imports sgenfe4.EventoXsd

<ComClass(NFeCancela.ClassId, NFeCancela.InterfaceId, NFeCancela.EventsId)> _
Public Class NFeCancela

    Private gobjApp As ClassGlobalApp
    Private gobjVenda As GlobaisLoja.ClassVenda
    Private sErro As String
    Private lNumNotaFiscal As Long
    Private sSerie As String
    Private lLote As Long
    Private sNFeChaveAcesso As String

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "6C53F528-48C6-4d53-A1BD-E12A5930F7E0"
    Public Const InterfaceId As String = "A0DFD0DC-1475-406a-856E-39404579E502"
    Public Const EventsId As String = "53FB280E-E32D-458a-B41A-058F9BC0D05B"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
        'MsgBox("oi, cancelando...")
    End Sub

    Private Function Email_Monta_Texto_CancVenda(ByRef sTexto As String) As Long

        Try

            sTexto = gobjVenda.objNFeInfo.sEmitRazaoSocial & " - " & gobjVenda.objNFeInfo.sEmitNomeReduzido & vbCrLf
            sTexto = sTexto & vbCrLf & vbCrLf
            sTexto = sTexto & "Nota Fiscal de Consumidor Eletronica - NFCe" & vbCrLf
            sTexto = sTexto & vbCrLf & vbCrLf
            sTexto = sTexto & "C A N C E L A M E N T O"
            sTexto = sTexto & vbCrLf & vbCrLf
            sTexto = sTexto & "Link para consulta pelo QRCode: " & gobjVenda.objCupomFiscal.sNFCeQRCode & vbCrLf
            sTexto = sTexto & vbCrLf & vbCrLf
            sTexto = sTexto & "Chave de Acesso: " & gobjVenda.objCupomFiscal.sNFeChaveAcesso & vbCrLf
            sTexto = sTexto & vbCrLf & vbCrLf
            sTexto = sTexto & "Emitida através do Sistema Corporator"

            Email_Monta_Texto_CancVenda = SUCESSO

        Catch ex As Exception

            Email_Monta_Texto_CancVenda = 1

        Finally

        End Try

    End Function

    Private Function Evento_Grava_Xml(ByVal objprocEventoNFe As CancXsd.TProcEvento, ByRef sArquivo As String) As Long

        Dim DocDados2 As XmlDocument = New XmlDocument
        Dim XMLStreamDados1 = New MemoryStream(10000)
        Dim XMLStreamDados2 = New MemoryStream(10000)
        Dim iPos As Integer
        Dim XMLStreamDados = New MemoryStream(10000)

        Try

            Dim mySerializerProcEvento As New XmlSerializer(GetType(CancXsd.TProcEvento))

            Dim XMLStream1 = New MemoryStream(10000)

            mySerializerProcEvento.Serialize(XMLStream1, objprocEventoNFe)

            Dim xmw1 As Byte()
            Dim XMLStringProc As String

            xmw1 = XMLStream1.ToArray

            XMLStringProc = System.Text.Encoding.UTF8.GetString(xmw1)

            XMLStringProc = Mid(XMLStringProc, 1, 19) & " encoding=""UTF-8"" " & Mid(XMLStringProc, 20)

            iPos = InStr(XMLStringProc, "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema""")

            If iPos <> 0 Then

                XMLStringProc = Mid(XMLStringProc, 1, iPos - 1) & Mid(XMLStringProc, iPos + 99)

            End If

            Dim xDadosCanc As Byte()

            xDadosCanc = System.Text.Encoding.UTF8.GetBytes(XMLStringProc)

            XMLStreamDados = New MemoryStream(10000)

            XMLStreamDados.Write(xDadosCanc, 0, xDadosCanc.Length)

            gobjApp.sErro = "25"
            gobjApp.sMsg1 = "vai gravar o xml"

            Dim DocDadosCanc As XmlDocument = New XmlDocument
            XMLStreamDados.Position = 0
            DocDadosCanc.Load(XMLStreamDados)
            sArquivo = gobjApp.sDirXml & objprocEventoNFe.retEvento.infEvento.tpEvento.ToString & "-" & objprocEventoNFe.retEvento.infEvento.chNFe & "-" & objprocEventoNFe.retEvento.infEvento.nSeqEvento & "-procEventoNfe.xml"

            Dim writer As New XmlTextWriter(sArquivo, Nothing)

            writer.Formatting = Formatting.None
            DocDadosCanc.WriteTo(writer)
            writer.Close()

            Evento_Grava_Xml = 0

        Catch ex As Exception

            Evento_Grava_Xml = 1

            Call gobjApp.GravarLog(ex.Message)

        End Try

    End Function

    Public Function NFCE_Cancela(ByVal objVenda As GlobaisLoja.ClassVenda) As Long

        Try

            Dim NfeRecEvento As New recepcaoevento.NFeRecepcaoEvento4
            'Dim NfeCabec As New recepcaoevento.nfeCabecMsg
            Dim objCancEvento As CancXsd.TEvento
            Dim objDescEvento As TEventoInfEventoDetEventoDescEvento
            'Dim objCabecMsg As recepcaoevento.nfeCabecMsg = New recepcaoevento.nfeCabecMsg
            Dim lErro As Long
            Dim XMLString1 As String
            Dim iPos As Integer
            Dim objEnvCancEvento As CancXsd.TEnvEvento = New CancXsd.TEnvEvento
            Dim AD As AssinaturaDigital = New AssinaturaDigital
            Dim XMLStreamDados As MemoryStream = New MemoryStream(10000)
            Dim objValidaXML As ClassValidaXML = New ClassValidaXML
            Dim XMLStringEvento As String
            Dim XMLString2 As String
            Dim XMLStream1 As MemoryStream = New MemoryStream(10000)
            'Dim XMLStreamCabec As MemoryStream = New MemoryStream(10000)
            Dim XMLStreamRet As MemoryStream = New MemoryStream(10000)
            Dim sArquivo As String
            Dim xRet As Byte()
            Dim xmlNode1 As XmlNode, objProcEvento As New CancXsd.TProcEvento

            gobjVenda = objVenda
            gobjApp = New ClassGlobalApp(objVenda, "NFCe")

            lErro = gobjApp.Obtem_Certificado()
            If lErro <> SUCESSO Then
            End If

            sNFeChaveAcesso = gobjVenda.objCupomFiscal.sNFeChaveAcesso

            objCancEvento = New CancXsd.TEvento

            objCancEvento.versao = "1.00"

            objCancEvento.infEvento = New CancXsd.TEventoInfEvento

            objCancEvento.infEvento.Id = "ID" & "110111" & sNFeChaveAcesso & "01"

            Dim sUF As String

            sUF = Left(sNFeChaveAcesso, 2)

            objCancEvento.infEvento.cOrgao = GetCode(Of CancXsd.TCOrgaoIBGE)(sUF)

            If gobjApp.iNFeAmbiente = NFE_AMBIENTE_HOMOLOGACAO Then
                objCancEvento.infEvento.tpAmb = CancXsd.TAmb.Item2
            Else
                objCancEvento.infEvento.tpAmb = CancXsd.TAmb.Item1
            End If

            objCancEvento.infEvento.ItemElementName = CancXsd.ItemChoiceType.CNPJ
            objCancEvento.infEvento.Item = Mid(sNFeChaveAcesso, 7, 14)

            objCancEvento.infEvento.chNFe = sNFeChaveAcesso

            objCancEvento.infEvento.dhEvento = Format(Now.Date, "yyyy-MM-dd") & "T" & Format(TimeOfDay, "HH:mm:ss") & Format(TimeOfDay, "zzz")

            objCancEvento.infEvento.tpEvento = TEventoInfEventoTpEvento.Item110111

            objCancEvento.infEvento.nSeqEvento = "1"

            objCancEvento.infEvento.verEvento = TEventoInfEventoVerEvento.Item100

            objCancEvento.infEvento.detEvento = New CancXsd.TEventoInfEventoDetEvento

            objCancEvento.infEvento.detEvento.versao = TEventoInfEventoDetEventoVersao.Item100

            objDescEvento = New TEventoInfEventoDetEventoDescEvento
            objCancEvento.infEvento.detEvento.descEvento = objDescEvento

            objCancEvento.infEvento.detEvento.xJust = Trim(DesacentuaTexto("cancelamento de nfce pelo usuario."))

            objCancEvento.infEvento.detEvento.nProt = gobjVenda.objCupomFiscal.sNFenProt

            Dim mySerializer As New XmlSerializer(GetType(CancXsd.TEvento))

            Dim XMLStream2 = New MemoryStream(10000)
            mySerializer.Serialize(XMLStream2, objCancEvento)

            Dim xm2 As Byte()
            xm2 = XMLStream2.ToArray

            XMLString1 = System.Text.Encoding.UTF8.GetString(xm2)

            '                    XMLString1 = Mid(XMLString1, 1, 19) & " encoding=""UTF-8"" " & Mid(XMLString1, 20)

            iPos = InStr(XMLString1, "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema""")

            If iPos <> 0 Then

                XMLString1 = Mid(XMLString1, 1, iPos - 1) & Mid(XMLString1, iPos + 99)

            End If



            lErro = AD.Assinar(XMLString1, "infEvento", gobjApp.cert, gobjApp.iDebug)
            If lErro <> SUCESSO Then Throw New System.Exception("")




            Dim xMlD As XmlDocument

            xMlD = AD.XMLDocAssinado()

            Dim xString As String
            xString = AD.XMLStringAssinado

            XMLStringEvento = Mid(xString, 22)

            objEnvCancEvento.idLote = lLote
            objEnvCancEvento.versao = "1.00"

            Dim mySerializerw As New XmlSerializer(GetType(CancXsd.TEnvEvento))

            XMLStream1 = New MemoryStream(10000)

            mySerializerw.Serialize(XMLStream1, objEnvCancEvento)

            Dim xmw As Byte()
            xmw = XMLStream1.ToArray

            XMLString1 = System.Text.Encoding.UTF8.GetString(xmw)

            XMLString2 = Mid(XMLString1, 1, Len(XMLString1) - 12) & XMLStringEvento & Mid(XMLString1, Len(XMLString1) - 12)

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

            If gobjApp.iDebug = 1 Then MsgBox("21")
            gobjApp.sErro = "21"
            gobjApp.sMsg1 = "vai gravar o xml"


            Dim DocDados As XmlDocument = New XmlDocument
            XMLStreamDados.Position = 0
            DocDados.Load(XMLStreamDados)
            sArquivo = gobjApp.sDirXml & sNFeChaveAcesso & "-env-evento-canc.xml"
            '            DocDados.Save(sArquivo)
            Dim writer1 As New XmlTextWriter(sArquivo, Nothing)

            writer1.Formatting = Formatting.None
            DocDados.WriteTo(writer1)
            writer1.Close()

            '???? deserializar evento em objProcEvento.evento
            Dim mySerializerEnvCancAux As New XmlSerializer(GetType(CancXsd.TEnvEvento)), objEnvEvento As New CancXsd.TEnvEvento
            XMLStreamDados.Position = 0
            objEnvEvento = mySerializerEnvCancAux.Deserialize(XMLStreamDados)
            objProcEvento.evento = objEnvEvento.evento(0)
            '????

            If gobjApp.iDebug = 1 Then MsgBox("22")
            gobjApp.sErro = "22"
            gobjApp.sMsg1 = "vai validar o arquivo xml de envio de evento de cancelamento"

            lErro = gobjApp.XML_Valida(sArquivo, gobjApp.sDirXsd & "envEventoCancNFe_v1.00.xsd")
            If lErro = 1 Then

                Call gobjApp.GravarLog("ERRO - Encerrado o envio do evento de cancelamento")

                Exit Try
            End If

            If gobjApp.iDebug = 1 Then MsgBox("23")
            gobjApp.sErro = "23"
            gobjApp.sMsg1 = "vai enviar o evento de cancelamento"

            Dim DocDados1 As New XmlDocument

            Call Salva_Arquivo(DocDados1, XMLString2)

            'NfeCabec.cUF = CStr(gobjApp.iUFCodIBGE)
            'NfeCabec.versaoDados = "1.00"

            'NfeRecEvento.nfeCabecMsgValue = NfeCabec

            Dim sURL As String
            sURL = ""
            Call WS_Obter_URL(sURL, gobjApp.iNFeAmbiente = NFE_AMBIENTE_HOMOLOGACAO, gobjApp.sSiglaEstado, "RecepcaoEvento", "NFCe")

            NfeRecEvento.Url = sURL

            Dim XMLStringRetEnvCanc As String

            NfeRecEvento.ClientCertificates.Add(gobjApp.cert)
            xmlNode1 = NfeRecEvento.nfeRecepcaoEvento(DocDados1)

            XMLStringRetEnvCanc = xmlNode1.OuterXml

            xRet = System.Text.Encoding.UTF8.GetBytes(XMLStringRetEnvCanc)

            XMLStreamRet = New MemoryStream(10000)
            XMLStreamRet.Write(xRet, 0, xRet.Length)

            Dim mySerializerRetEnvCCE As New XmlSerializer(GetType(CancXsd.TRetEnvEvento))

            Dim objRetEnvCanc As CancXsd.TRetEnvEvento = New CancXsd.TRetEnvEvento

            XMLStreamRet.Position = 0

            objRetEnvCanc = mySerializerRetEnvCCE.Deserialize(XMLStreamRet)

            gobjApp.sErro = "24"
            gobjApp.sMsg1 = "trata o retorno do evento de cancelamento"
            If gobjApp.iDebug = 1 Then MsgBox(gobjApp.sErro)


            If Not objRetEnvCanc.retEvento Is Nothing Then

                For iIndice = 0 To objRetEnvCanc.retEvento.Length - 1

                    gobjApp.GravarLog("Retorno do evento " & IIf(objRetEnvCanc.retEvento(iIndice).infEvento.xEvento Is Nothing, "cancelamento", objRetEnvCanc.retEvento(iIndice).infEvento.xEvento) & " nSeqEvento = " & objRetEnvCanc.retEvento(iIndice).infEvento.nSeqEvento & " chave NFe = " & objRetEnvCanc.retEvento(iIndice).infEvento.chNFe)
                    gobjApp.GravarLog("cStat = " & objRetEnvCanc.retEvento(iIndice).infEvento.cStat & " xMotivo1 = " & objRetEnvCanc.retEvento(iIndice).infEvento.xMotivo)

                    If objRetEnvCanc.retEvento(iIndice).infEvento.cStat = "135" Or objRetEnvCanc.retEvento(iIndice).infEvento.cStat = "136" Then

                        objProcEvento.retEvento = objRetEnvCanc.retEvento(iIndice)

                        With gobjVenda.objCupomFiscal
                            .iNFeCancHomologado = 1
                            .iNFeCancPendente = 0
                            .sNFeCancnProt = objRetEnvCanc.retEvento(iIndice).infEvento.nProt
                            .dtNFeCancData = UTCParaDate(objRetEnvCanc.retEvento(iIndice).infEvento.dhRegEvento)
                            .dNFeCancHora = UTCParaHora(objRetEnvCanc.retEvento(iIndice).infEvento.dhRegEvento)
                        End With

                    Else
                        gobjApp.GravarLog("Ocorreu um erro no envio do evento de cancelamento, cStat = " & objRetEnvCanc.cStat & " xMotivo = " & objRetEnvCanc.xMotivo)
                    End If

                    Exit For

                Next

            Else

                gobjApp.GravarLog("Ocorreu um erro no envio do evento de cancelamento, cStat = " & objRetEnvCanc.cStat & " xMotivo = " & objRetEnvCanc.xMotivo)

            End If

            If gobjApp.iDebug = 1 Then MsgBox("26")
            gobjApp.sErro = "26"
            gobjApp.sMsg1 = "vai gravar o xml"

            If gobjVenda.objCupomFiscal.iNFeCancHomologado = 0 Then

                Dim DocDados2 As XmlDocument = New XmlDocument
                XMLStreamRet.Position = 0
                DocDados2.Load(XMLStreamRet)
                sArquivo = gobjApp.sDirXml & sNFeChaveAcesso & "-ret-evento-canc.xml"
                DocDados2.Save(sArquivo)

                Throw New System.Exception("")
            Else
                lErro = Evento_Grava_Xml(objProcEvento, sArquivo)
                If lErro <> SUCESSO Then Throw New System.Exception("")
                gobjVenda.objCupomFiscal.sNFeCancArqXml = sArquivo

                If gobjVenda.objNFeInfo.iNFCeEnviarEmail <> 0 And gobjVenda.objCupomFiscal.sNFCeQRCode <> "" Then
                    Dim sEMailTexto As String = ""
                    Call Email_Monta_Texto_CancVenda(sEMailTexto)
                    Call Email_Enviar(gobjVenda.objNFeInfo.sSMTP, gobjVenda.objNFeInfo.sSMTPUsu, gobjVenda.objNFeInfo.sSMTPSenha, CStr(gobjVenda.objNFeInfo.lSMTPPorta), "Comprovante de Cancelamento de NFCe", sEMailTexto, IIf(gobjVenda.objCupomFiscal.sEndEntEmail <> "", gobjVenda.objCupomFiscal.sEndEntEmail, gobjVenda.objNFeInfo.sSMTPUsu))
                End If

            End If

            NFCE_Cancela = SUCESSO

        Catch ex As Exception

            NFCE_Cancela = 1

            Dim sMsg2 As String

            If ex.InnerException Is Nothing Then
                sMsg2 = ""
            Else
                sMsg2 = " - " & ex.InnerException.Message
            End If

            gobjApp.GravarLog("ERRO - " & ex.Message & sMsg2)
            gobjApp.GravarLog("ERRO - " & gobjApp.sErro & " - " & gobjApp.sMsg1)

            gobjApp.Form1.Show()

            MsgBox("Verifique a mensagem de erro.")

        Finally

            gobjApp.Terminar()
            gobjApp = Nothing
            gobjVenda = Nothing

            NFCE_Cancela = SUCESSO

        End Try

    End Function

End Class


