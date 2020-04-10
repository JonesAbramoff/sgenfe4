Imports System
Imports System.Xml.Serialization
Imports System.IO
Imports System.Text
Imports System.Security.Cryptography.Xml
Imports System.Security.Cryptography.X509Certificates
Imports System.Xml
Imports System.Xml.Schema
Imports Microsoft.Win32
Imports sgenfe4.NFeXsd

<ComClass(NFeCarrega.ClassId, NFeCarrega.InterfaceId, NFeCarrega.EventsId)> _
Public Class NFeCarrega

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "DE73EC69-8449-4feb-A7A0-3C31B48B6B5A"
    Public Const InterfaceId As String = "3E1C93E3-4EF0-4c36-8B13-FFCE6CF1964E"
    Public Const EventsId As String = "4BB1579F-0FFE-4925-B21E-6B3E6DEF70F2"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
        'MsgBox("oi")
    End Sub

    Public Function NFe_Carrega(ByVal sArqXml As String, ByVal objNF As GlobaisCRFAT.ClassNFiscal, ByVal objCliente As GlobaisCRFAT.ClassCliente, ByVal objEndCli As GlobaisAdm.ClassEndereco, ByVal objProt As GlobaisFAT.ClassNFeFedProtNFeView) As Long
        'carrega o ArqXml e preenche informacoes

        Try

            Dim DocDados As XmlDocument = New XmlDocument
            Dim XMLString1 As String
            Dim xm As Byte()
            Dim XMLStreamDados = New MemoryStream(10000)
            Dim XMLStreamDados1 = New MemoryStream(10000)
            Dim xDados1 As Byte()
            Dim objItem As GlobaisCRFAT.ClassItemNF

            Dim objICMS As TNFeInfNFeDetImpostoICMS
            Dim objPIS As TNFeInfNFeDetImpostoPIS
            Dim objCOFINS As TNFeInfNFeDetImpostoCOFINS
            Dim objISS As TNFeInfNFeDetImpostoISSQN
            Dim objPISST As TNFeInfNFeDetImpostoPISST
            Dim objCOFINSST As TNFeInfNFeDetImpostoCOFINSST

            Dim objICMS00 As TNFeInfNFeDetImpostoICMSICMS00
            Dim objICMS10 As TNFeInfNFeDetImpostoICMSICMS10
            Dim objICMS20 As TNFeInfNFeDetImpostoICMSICMS20
            Dim objICMS30 As TNFeInfNFeDetImpostoICMSICMS30
            Dim objICMS40 As TNFeInfNFeDetImpostoICMSICMS40
            Dim objICMS51 As TNFeInfNFeDetImpostoICMSICMS51
            Dim objICMS60 As TNFeInfNFeDetImpostoICMSICMS60
            Dim objICMS70 As TNFeInfNFeDetImpostoICMSICMS70
            Dim objICMS90 As TNFeInfNFeDetImpostoICMSICMS90
            Dim objICMSPart As TNFeInfNFeDetImpostoICMSICMSPart
            Dim objICMSST As TNFeInfNFeDetImpostoICMSICMSST
            Dim objICMS101 As TNFeInfNFeDetImpostoICMSICMSSN101
            Dim objICMS102 As TNFeInfNFeDetImpostoICMSICMSSN102
            Dim objICMS201 As TNFeInfNFeDetImpostoICMSICMSSN201
            Dim objICMS202 As TNFeInfNFeDetImpostoICMSICMSSN202
            Dim objICMS500 As TNFeInfNFeDetImpostoICMSICMSSN500
            Dim objICMS900 As TNFeInfNFeDetImpostoICMSICMSSN900

            Dim objPISAliq As TNFeInfNFeDetImpostoPISPISAliq
            Dim objPISQtde As TNFeInfNFeDetImpostoPISPISQtde
            Dim objPISNT As TNFeInfNFeDetImpostoPISPISNT
            Dim objPISOut As TNFeInfNFeDetImpostoPISPISOutr

            Dim objCOFINSAliq As TNFeInfNFeDetImpostoCOFINSCOFINSAliq
            Dim objCOFINSQtde As TNFeInfNFeDetImpostoCOFINSCOFINSQtde
            Dim objCOFINSNT As TNFeInfNFeDetImpostoCOFINSCOFINSNT
            Dim objCOFINSOut As TNFeInfNFeDetImpostoCOFINSCOFINSOutr
            Dim objNFe As TNFe

            Dim sArqComErros As String = "", iImposto As Integer

            DocDados.Load(sArqXml)

            Dim writer1 As New XmlTextWriter(XMLStreamDados1, Nothing)

            writer1.Formatting = Formatting.None
            DocDados.WriteTo(writer1)
            writer1.Flush()

            '        DocDados.Save(XMLStreamDados1)

            xm = XMLStreamDados1.ToArray

            XMLString1 = Replace(Replace(System.Text.Encoding.UTF8.GetString(xm), "BRASIL", "Brasil"), "<?xml version=""1.0"" encoding=""UTF-8""?>", "")

            xDados1 = System.Text.Encoding.UTF8.GetBytes(XMLString1)

            XMLStreamDados.Write(xDados1, 0, xDados1.Length)

            If InStr(sArqXml, "-env-lot.xml") = 0 Then

                If InStr(sArqXml, "-pre.xml") = 0 Then

                    Dim objNFeProc As New TNfeProc
                    Dim mySerializerImportaXML As New XmlSerializer(GetType(TNfeProc))

                    XMLStreamDados.Position = 0

                    objNFeProc = mySerializerImportaXML.Deserialize(XMLStreamDados)
                    objNFe = objNFeProc.NFe

                    With objNFeProc.protNFe.infProt
                        objProt.scStat = .cStat
                        objProt.schNFe = .chNFe
                        objProt.dtData = UTCParaDate(.dhRecbto)
                        objProt.dHora = UTCParaHora(.dhRecbto)
                        objProt.snProt = .nProt
                        objProt.itpAmb = .tpAmb 'CInt(GetXmlAttrNameFromEnumValue(Of TAmb)(.tpAmb))
                        objProt.sverAplic = .verAplic
                        objProt.sVersao = objNFe.infNFe.versao
                        objProt.sxMotivo = .xMotivo
                    End With

                Else

                    Dim mySerializerImportaXML As New XmlSerializer(GetType(TNFe))

                    XMLStreamDados.Position = 0

                    objNFe = mySerializerImportaXML.Deserialize(XMLStreamDados)

                    objProt.scStat = "100"
                    objProt.schNFe = NFeXml_Conv_Texto(Right(objNFe.infNFe.Id, 44))
                    objProt.dtData = DATA_NULA
                    objProt.dHora = 0
                    objProt.snProt = ""
                    objProt.itpAmb = 0 'CInt(GetXmlAttrNameFromEnumValue(Of TAmb)(.tpAmb))
                    objProt.sverAplic = ""
                    objProt.sVersao = objNFe.infNFe.versao
                    objProt.sxMotivo = "Autorizado o uso da NF-e"

                End If

            Else

                Dim mySerializerImportaXML As New XmlSerializer(GetType(TEnviNFe))
                Dim objEnviNFe As New TEnviNFe

                XMLStreamDados.Position = 0

                objEnviNFe = mySerializerImportaXML.Deserialize(XMLStreamDados)
                objNFe = objEnviNFe.NFe(0)

                objProt.scStat = "100"
                objProt.schNFe = NFeXml_Conv_Texto(Right(objNFe.infNFe.Id, 44))
                objProt.dtData = DATA_NULA
                objProt.dHora = 0
                objProt.snProt = ""
                objProt.itpAmb = 0 'CInt(GetXmlAttrNameFromEnumValue(Of TAmb)(.tpAmb))
                objProt.sverAplic = ""
                objProt.sVersao = objNFe.infNFe.versao
                objProt.sxMotivo = "Autorizado o uso da NF-e"

            End If

                objNF.sChvNFe = NFeXml_Conv_Texto(Right(objNFe.infNFe.Id, 44))

                Select Case objNFe.infNFe.ide.mod

                    Case TMod.Item55
                        objNF.iTipoNFiscal = DOCINFO_NFISFVPAF

                    Case TMod.Item65
                        objNF.iTipoNFiscal = DOCINFO_NFCEPDV

                End Select

                objNF.lNumNotaFiscal = NFeXml_Conv_Long(objNFe.infNFe.ide.nNF)
                objNF.sSerie = NFeXml_Conv_Texto(objNFe.infNFe.ide.serie) & "-e"
                objNF.dtDataEmissao = UTCParaDate(objNFe.infNFe.ide.dhEmi)
                objNF.dtHoraEmissao = System.DateTime.Parse(objNFe.infNFe.ide.dhEmi).ToLocalTime
                objNF.dtDataSaida = UTCParaDate(objNFe.infNFe.ide.dhSaiEnt)
                If objNF.dtDataSaida = DATA_NULA Then objNF.dtDataSaida = objNF.dtDataEmissao
                objNF.dtHoraSaida = objNF.dtHoraEmissao
                objNF.dValorProdutos = NFeXml_Conv_Valor(objNFe.infNFe.total.ICMSTot.vProd)
                objNF.dValorFrete = NFeXml_Conv_Valor(objNFe.infNFe.total.ICMSTot.vFrete)
                objNF.dValorSeguro = NFeXml_Conv_Valor(objNFe.infNFe.total.ICMSTot.vSeg)
                objNF.dValorOutrasDespesas = NFeXml_Conv_Valor(objNFe.infNFe.total.ICMSTot.vOutro)
                objNF.dValorDescontoItens = NFeXml_Conv_Valor(objNFe.infNFe.total.ICMSTot.vDesc)
                objNF.dValorTotal = NFeXml_Conv_Valor(objNFe.infNFe.total.ICMSTot.vNF)

                objNF.objTributacaoNF.dICMSBase = NFeXml_Conv_Valor(objNFe.infNFe.total.ICMSTot.vBC)
                objNF.objTributacaoNF.dICMSValor = NFeXml_Conv_Valor(objNFe.infNFe.total.ICMSTot.vICMS)
                objNF.objTributacaoNF.dICMSSubstBase = NFeXml_Conv_Valor(objNFe.infNFe.total.ICMSTot.vBCST)
                objNF.objTributacaoNF.dICMSSubstValor = NFeXml_Conv_Valor(objNFe.infNFe.total.ICMSTot.vST)
                objNF.objTributacaoNF.dIPIValor = NFeXml_Conv_Valor(objNFe.infNFe.total.ICMSTot.vIPI)
                objNF.objTributacaoNF.dPISValor = NFeXml_Conv_Valor(objNFe.infNFe.total.ICMSTot.vPIS)
                objNF.objTributacaoNF.dCOFINSValor = NFeXml_Conv_Valor(objNFe.infNFe.total.ICMSTot.vCOFINS)

                If Not (objNFe.infNFe.total.ISSQNtot Is Nothing) Then
                    objNF.objTributacaoNF.dISSBase = NFeXml_Conv_Valor(objNFe.infNFe.total.ISSQNtot.vBC)
                    objNF.objTributacaoNF.dISSValor = NFeXml_Conv_Valor(objNFe.infNFe.total.ISSQNtot.vISS)
                    objNF.objTributacaoNF.dPISValor = objNF.objTributacaoNF.dPISValor + NFeXml_Conv_Valor(objNFe.infNFe.total.ISSQNtot.vPIS)
                    objNF.objTributacaoNF.dCOFINSValor = objNF.objTributacaoNF.dCOFINSValor + NFeXml_Conv_Valor(objNFe.infNFe.total.ISSQNtot.vCOFINS)
                End If

                If Not (objNFe.infNFe.total.retTrib Is Nothing) Then

                    objNF.objTributacaoNF.dPISRetido = NFeXml_Conv_Valor(objNFe.infNFe.total.retTrib.vRetPIS)
                    objNF.objTributacaoNF.dCOFINSRetido = NFeXml_Conv_Valor(objNFe.infNFe.total.retTrib.vRetCOFINS)
                    objNF.objTributacaoNF.dCSLLRetido = NFeXml_Conv_Valor(objNFe.infNFe.total.retTrib.vRetCSLL)
                    objNF.objTributacaoNF.dIRRFBase = NFeXml_Conv_Valor(objNFe.infNFe.total.retTrib.vBCIRRF)
                    objNF.objTributacaoNF.dIRRFValor = NFeXml_Conv_Valor(objNFe.infNFe.total.retTrib.vIRRF)
                    objNF.objTributacaoNF.dINSSBase = NFeXml_Conv_Valor(objNFe.infNFe.total.retTrib.vBCRetPrev)
                    objNF.objTributacaoNF.dINSSValor = NFeXml_Conv_Valor(objNFe.infNFe.total.retTrib.vRetPrev)
                    If objNF.objTributacaoNF.dINSSValor <> 0 Then objNF.objTributacaoNF.iINSSRetido = 1

                End If

                If Not (objNFe.infNFe.dest) Is Nothing Then

                    objCliente.sRazaoSocial = NFeXml_Conv_Texto(objNFe.infNFe.dest.xNome)
                If objNFe.infNFe.dest.ItemElementName = ItemChoiceType5.CPF Or objNFe.infNFe.dest.ItemElementName = ItemChoiceType5.CNPJ Then
                    objCliente.sCgc = NFeXml_Conv_Texto(objNFe.infNFe.dest.Item)
                End If
                    objCliente.sInscricaoEstadual = NFeXml_Conv_Texto(objNFe.infNFe.dest.IE)
                    objCliente.sInscricaoSuframa = NFeXml_Conv_Texto(objNFe.infNFe.dest.ISUF)

                    If Not objNFe.infNFe.dest.enderDest Is Nothing Then

                        objEndCli.sLogradouro = NFeXml_Conv_Texto(objNFe.infNFe.dest.enderDest.xLgr)
                        objEndCli.lNumero = NFeXml_Conv_Long(objNFe.infNFe.dest.enderDest.nro)
                        objEndCli.sComplemento = NFeXml_Conv_Texto(objNFe.infNFe.dest.enderDest.xCpl)
                        objEndCli.sBairro = NFeXml_Conv_Texto(objNFe.infNFe.dest.enderDest.xBairro)
                        objEndCli.sCidade = NFeXml_Conv_Texto(objNFe.infNFe.dest.enderDest.xMun)
                        objEndCli.sSiglaEstado = GetXmlAttrNameFromEnumValue(Of TUf)(objNFe.infNFe.dest.enderDest.UF)
                        objEndCli.sCEP = NFeXml_Conv_Texto(objNFe.infNFe.dest.enderDest.CEP)
                        If NFeXml_Conv_Texto(objNFe.infNFe.dest.enderDest.xPais) = "1058" Then
                            objEndCli.iCodigoPais = 1
                        Else
                            objEndCli.iCodigoPais = 1 '?????
                        End If
                        objEndCli.sTelNumero1 = NFeXml_Conv_Texto(objNFe.infNFe.dest.enderDest.fone)
                        objEndCli.sEmail = NFeXml_Conv_Texto(objNFe.infNFe.dest.email)

                    End If

                End If

                objNF.objTributacaoNF.dTotTrib = NFeXml_Conv_Valor(objNFe.infNFe.total.ICMSTot.vTotTrib)

                For iIndice = 1 To objNFe.infNFe.det.Length

                    objItem = New GlobaisCRFAT.ClassItemNF

                    objNF.ColItensNF.Add1(objItem)

                    'Inicializa parte tributária
                    Call objItem.Inicializa_Tributacao()
                    Call objItem.objTributacaoItemNF.Coloca_Auto()

                    If objNFe.infNFe.emit.CRT = TNFeInfNFeEmitCRT.Item1 Then
                        objItem.objTributacao.iRegimeTributario = 1
                    Else
                        objItem.objTributacao.iRegimeTributario = 3
                    End If

                    objPIS = objNFe.infNFe.det(iIndice - 1).imposto.PIS
                    objCOFINS = objNFe.infNFe.det(iIndice - 1).imposto.COFINS
                    objPISST = objNFe.infNFe.det(iIndice - 1).imposto.PISST
                    objCOFINSST = objNFe.infNFe.det(iIndice - 1).imposto.COFINSST

                    objISS = New TNFeInfNFeDetImpostoISSQN

                    If objNFe.infNFe.det(iIndice - 1).imposto.Items(0).GetType.Name = objISS.GetType.Name Then

                        objISS = objNFe.infNFe.det(iIndice - 1).imposto.Items(0)

                        objItem.objTributacao.dISSBase = NFeXml_Conv_Valor(objISS.vBC)
                        objItem.objTributacao.dISSAliquota = NFeXml_Conv_Perc(objISS.vAliq)
                        objItem.objTributacao.dISSValor = NFeXml_Conv_Valor(objISS.vISSQN)

                        objItem.objTributacao.sISSCST = "N" '??? campo eliminado do xml

                    Else

                        objICMS = New TNFeInfNFeDetImpostoICMS

                        For iImposto = 0 To objNFe.infNFe.det(iIndice - 1).imposto.Items.Length - 1

                            Select Case objNFe.infNFe.det(iIndice - 1).imposto.Items(iImposto).GetType.Name

                                Case objICMS.GetType.Name
                                    objICMS = objNFe.infNFe.det(iIndice - 1).imposto.Items(iImposto)
                            End Select

                        Next

                        objICMS00 = New TNFeInfNFeDetImpostoICMSICMS00
                        objICMS10 = New TNFeInfNFeDetImpostoICMSICMS10
                        objICMS20 = New TNFeInfNFeDetImpostoICMSICMS20
                        objICMS30 = New TNFeInfNFeDetImpostoICMSICMS30
                        objICMS40 = New TNFeInfNFeDetImpostoICMSICMS40
                        objICMS51 = New TNFeInfNFeDetImpostoICMSICMS51
                        objICMS60 = New TNFeInfNFeDetImpostoICMSICMS60
                        objICMS70 = New TNFeInfNFeDetImpostoICMSICMS70
                        objICMS90 = New TNFeInfNFeDetImpostoICMSICMS90
                        objICMSPart = New TNFeInfNFeDetImpostoICMSICMSPart
                        objICMSST = New TNFeInfNFeDetImpostoICMSICMSST
                        objICMS101 = New TNFeInfNFeDetImpostoICMSICMSSN101
                        objICMS102 = New TNFeInfNFeDetImpostoICMSICMSSN102
                        objICMS201 = New TNFeInfNFeDetImpostoICMSICMSSN201
                        objICMS202 = New TNFeInfNFeDetImpostoICMSICMSSN202
                        objICMS500 = New TNFeInfNFeDetImpostoICMSICMSSN500
                        objICMS900 = New TNFeInfNFeDetImpostoICMSICMSSN900

                        Select Case objICMS.Item.GetType.Name

                            Case objICMS00.GetType.Name
                                objICMS00 = objICMS.Item
                                objItem.objTributacao.iOrigemMercadoria = objICMS00.orig
                                objItem.objTributacao.sCST = CStr(objItem.objTributacao.iOrigemMercadoria) & "00"
                                objItem.objTributacao.iICMSTipo = 1
                                objItem.objTributacao.iICMSBaseModalidade = objICMS00.modBC
                                objItem.objTributacao.dICMSBase = NFeXml_Conv_Valor(objICMS00.vBC)
                                objItem.objTributacao.dICMSAliquota = NFeXml_Conv_Perc(objICMS00.pICMS)
                                objItem.objTributacao.dICMSValor = NFeXml_Conv_Valor(objICMS00.vICMS)

                            Case objICMS10.GetType.Name
                                objICMS10 = objICMS.Item
                                objItem.objTributacao.iOrigemMercadoria = objICMS10.orig
                                objItem.objTributacao.sCST = CStr(objItem.objTributacao.iOrigemMercadoria) & "10"
                                objItem.objTributacao.iICMSTipo = 6
                                objItem.objTributacao.iICMSBaseModalidade = objICMS10.modBC
                                objItem.objTributacao.dICMSBase = NFeXml_Conv_Valor(objICMS10.vBC)
                                objItem.objTributacao.dICMSAliquota = NFeXml_Conv_Perc(objICMS10.pICMS)
                                objItem.objTributacao.dICMSValor = NFeXml_Conv_Valor(objICMS10.vICMS)
                                objItem.objTributacao.iICMSSubstBaseModalidade = objICMS10.modBCST
                                objItem.objTributacao.dICMSSubstPercMVA = NFeXml_Conv_Perc(objICMS10.pMVAST)
                                objItem.objTributacao.dICMSSubstBase = NFeXml_Conv_Valor(objICMS10.vBCST)
                                objItem.objTributacao.dICMSSubstPercRedBase = NFeXml_Conv_Perc(objICMS10.pRedBCST)
                                objItem.objTributacao.dICMSSubstAliquota = NFeXml_Conv_Perc(objICMS10.pICMSST)
                                objItem.objTributacao.dICMSSubstValor = NFeXml_Conv_Valor(objICMS10.vICMSST)

                            Case objICMS20.GetType.Name
                                objICMS20 = objICMS.Item
                                objItem.objTributacao.iOrigemMercadoria = objICMS20.orig
                                objItem.objTributacao.sCST = CStr(objItem.objTributacao.iOrigemMercadoria) & "20"
                                objItem.objTributacao.iICMSTipo = 7
                                objItem.objTributacao.iICMSBaseModalidade = objICMS20.modBC
                                objItem.objTributacao.dICMSBase = NFeXml_Conv_Valor(objICMS20.vBC)
                                objItem.objTributacao.dICMSPercRedBase = NFeXml_Conv_Perc(objICMS20.pRedBC)
                                objItem.objTributacao.dICMSAliquota = NFeXml_Conv_Perc(objICMS20.pICMS)
                                objItem.objTributacao.dICMSValor = NFeXml_Conv_Valor(objICMS20.vICMS)

                            Case objICMS30.GetType.Name
                                objICMS30 = objICMS.Item
                                objItem.objTributacao.iOrigemMercadoria = objICMS30.orig
                                objItem.objTributacao.sCST = CStr(objItem.objTributacao.iOrigemMercadoria) & "30"
                                objItem.objTributacao.iICMSTipo = 9
                                objItem.objTributacao.iICMSSubstBaseModalidade = objICMS30.modBCST
                                objItem.objTributacao.dICMSSubstPercMVA = NFeXml_Conv_Perc(objICMS30.pMVAST)
                                objItem.objTributacao.dICMSSubstBase = NFeXml_Conv_Valor(objICMS30.vBCST)
                                objItem.objTributacao.dICMSSubstAliquota = NFeXml_Conv_Perc(objICMS30.pICMSST)
                                objItem.objTributacao.dICMSSubstValor = NFeXml_Conv_Valor(objICMS30.vICMSST)
                                objItem.objTributacao.dICMSSubstPercRedBase = NFeXml_Conv_Perc(objICMS30.pRedBCST)

                            Case objICMS40.GetType.Name
                                objICMS40 = objICMS.Item
                                objItem.objTributacao.iOrigemMercadoria = objICMS40.orig
                                objItem.objTributacao.sCST = CStr(objItem.objTributacao.iOrigemMercadoria) & "40"
                                objItem.objTributacao.iICMSTipo = 2

                                Select Case objICMS40.motDesICMS
                                    Case TNFeInfNFeDetImpostoICMSICMS40MotDesICMS.Item1
                                        objItem.objTributacao.iICMSMotivo = 1
                                    Case TNFeInfNFeDetImpostoICMSICMS40MotDesICMS.Item3
                                        objItem.objTributacao.iICMSMotivo = 3
                                    Case TNFeInfNFeDetImpostoICMSICMS40MotDesICMS.Item4
                                        objItem.objTributacao.iICMSMotivo = 4
                                    Case TNFeInfNFeDetImpostoICMSICMS40MotDesICMS.Item5
                                        objItem.objTributacao.iICMSMotivo = 5
                                    Case TNFeInfNFeDetImpostoICMSICMS40MotDesICMS.Item6
                                        objItem.objTributacao.iICMSMotivo = 6
                                    Case TNFeInfNFeDetImpostoICMSICMS40MotDesICMS.Item7
                                        objItem.objTributacao.iICMSMotivo = 7
                                    Case TNFeInfNFeDetImpostoICMSICMS40MotDesICMS.Item9
                                        objItem.objTributacao.iICMSMotivo = 9
                                    Case Else
                                        objItem.objTributacao.iICMSMotivo = 0
                                End Select

                            Case objICMS51.GetType.Name
                                objICMS51 = objICMS.Item
                                objItem.objTributacao.iOrigemMercadoria = objICMS51.orig
                                objItem.objTributacao.sCST = CStr(objItem.objTributacao.iOrigemMercadoria) & "51"
                                objItem.objTributacao.iICMSTipo = 5
                                objItem.objTributacao.iICMSBaseModalidade = objICMS51.modBC
                                objItem.objTributacao.dICMSBase = NFeXml_Conv_Valor(objICMS51.vBC)
                                objItem.objTributacao.dICMSPercRedBase = NFeXml_Conv_Perc(objICMS51.pRedBC)
                                objItem.objTributacao.dICMSAliquota = NFeXml_Conv_Perc(objICMS51.pICMS)
                                objItem.objTributacao.dICMSValor = NFeXml_Conv_Valor(objICMS51.vICMS)

                            Case objICMS60.GetType.Name
                                objICMS60 = objICMS.Item
                                objItem.objTributacao.iOrigemMercadoria = objICMS60.orig
                                objItem.objTributacao.sCST = CStr(objItem.objTributacao.iOrigemMercadoria) & "60"
                                objItem.objTributacao.iICMSTipo = 8
                                objItem.objTributacao.dICMSSTCobrAntBase = NFeXml_Conv_Valor(objICMS60.vBCSTRet)
                                objItem.objTributacao.dICMSSTCobrAntValor = NFeXml_Conv_Valor(objICMS60.vICMSSTRet)

                            Case objICMS70.GetType.Name
                                objICMS70 = objICMS.Item
                                objItem.objTributacao.iOrigemMercadoria = objICMS70.orig
                                objItem.objTributacao.sCST = CStr(objItem.objTributacao.iOrigemMercadoria) & "70"
                                objItem.objTributacao.iICMSTipo = 4
                                objItem.objTributacao.iICMSBaseModalidade = objICMS70.modBC
                                objItem.objTributacao.dICMSPercRedBase = NFeXml_Conv_Perc(objICMS70.pRedBC)
                                objItem.objTributacao.dICMSBase = NFeXml_Conv_Valor(objICMS70.vBC)
                                objItem.objTributacao.dICMSAliquota = NFeXml_Conv_Perc(objICMS70.pICMS)
                                objItem.objTributacao.dICMSValor = NFeXml_Conv_Valor(objICMS70.vICMS)
                                objItem.objTributacao.iICMSSubstBaseModalidade = objICMS70.modBCST
                                objItem.objTributacao.dICMSSubstPercMVA = NFeXml_Conv_Perc(objICMS70.pMVAST)
                                objItem.objTributacao.dICMSSubstBase = NFeXml_Conv_Valor(objICMS70.vBCST)
                                objItem.objTributacao.dICMSSubstAliquota = NFeXml_Conv_Perc(objICMS70.pICMSST)
                                objItem.objTributacao.dICMSSubstValor = NFeXml_Conv_Valor(objICMS70.vICMSST)
                                objItem.objTributacao.dICMSSubstPercRedBase = NFeXml_Conv_Perc(objICMS70.pRedBCST)

                            Case objICMS90.GetType.Name
                                objICMS90 = objICMS.Item
                                objItem.objTributacao.iOrigemMercadoria = objICMS90.orig
                                objItem.objTributacao.sCST = CStr(objItem.objTributacao.iOrigemMercadoria) & "90"
                                objItem.objTributacao.iICMSTipo = 99
                                objItem.objTributacao.iICMSBaseModalidade = objICMS90.modBC
                                objItem.objTributacao.dICMSPercRedBase = NFeXml_Conv_Perc(objICMS90.pRedBC)
                                objItem.objTributacao.dICMSBase = NFeXml_Conv_Valor(objICMS90.vBC)
                                objItem.objTributacao.dICMSAliquota = NFeXml_Conv_Perc(objICMS90.pICMS)
                                objItem.objTributacao.dICMSValor = NFeXml_Conv_Valor(objICMS90.vICMS)
                                objItem.objTributacao.iICMSSubstBaseModalidade = objICMS90.modBCST
                                objItem.objTributacao.dICMSSubstPercMVA = NFeXml_Conv_Perc(objICMS90.pMVAST)
                                objItem.objTributacao.dICMSSubstBase = NFeXml_Conv_Valor(objICMS90.vBCST)
                                objItem.objTributacao.dICMSSubstAliquota = NFeXml_Conv_Perc(objICMS90.pICMSST)
                                objItem.objTributacao.dICMSSubstValor = NFeXml_Conv_Valor(objICMS90.vICMSST)
                                objItem.objTributacao.dICMSSubstPercRedBase = NFeXml_Conv_Perc(objICMS90.pRedBCST)

                            Case objICMSPart.GetType.Name
                                objICMSPart = objICMS.Item
                                objItem.objTributacao.iOrigemMercadoria = objICMSPart.orig
                                If TNFeInfNFeDetImpostoICMSICMSPartCST.Item90 = objICMSPart.CST Then
                                    objItem.objTributacao.sCST = CStr(objItem.objTributacao.iOrigemMercadoria) & "90"
                                    objItem.objTributacao.iICMSTipo = 12
                                Else
                                    objItem.objTributacao.sCST = CStr(objItem.objTributacao.iOrigemMercadoria) & "10"
                                    objItem.objTributacao.iICMSTipo = 11
                                End If
                                objItem.objTributacao.iICMSBaseModalidade = objICMSPart.modBC
                                objItem.objTributacao.dICMSPercRedBase = NFeXml_Conv_Perc(objICMSPart.pRedBC)
                                objItem.objTributacao.dICMSBase = NFeXml_Conv_Valor(objICMSPart.vBC)
                                objItem.objTributacao.dICMSAliquota = NFeXml_Conv_Perc(objICMSPart.pICMS)
                                objItem.objTributacao.dICMSValor = NFeXml_Conv_Valor(objICMSPart.vICMS)
                                objItem.objTributacao.iICMSSubstBaseModalidade = objICMSPart.modBCST
                                objItem.objTributacao.dICMSSubstPercMVA = NFeXml_Conv_Perc(objICMSPart.pMVAST)
                                objItem.objTributacao.dICMSSubstBase = NFeXml_Conv_Valor(objICMSPart.vBCST)
                                objItem.objTributacao.dICMSSubstAliquota = NFeXml_Conv_Perc(objICMSPart.pICMSST)
                                objItem.objTributacao.dICMSSubstValor = NFeXml_Conv_Valor(objICMSPart.vICMSST)
                                objItem.objTributacao.dICMSSubstPercRedBase = NFeXml_Conv_Perc(objICMSPart.pRedBCST)
                                objItem.objTributacao.dICMSpercBaseOperacaoPropria = NFeXml_Conv_Perc(objICMSPart.pBCOp)
                                objItem.objTributacao.sICMSUFDevidoST = objICMSPart.UFST.ToString

                            Case objICMSST.GetType.Name
                                objICMSST = objICMS.Item
                                objItem.objTributacao.iOrigemMercadoria = objICMSST.orig
                                objItem.objTributacao.sCST = CStr(objItem.objTributacao.iOrigemMercadoria) & "41"
                                objItem.objTributacao.iICMSTipo = 13
                                objItem.objTributacao.dICMSvBCSTRet = NFeXml_Conv_Valor(objICMSST.vBCSTRet)
                                objItem.objTributacao.dICMSvICMSSTRet = NFeXml_Conv_Valor(objICMSST.vICMSSTRet)
                                objItem.objTributacao.dICMSvBCSTDest = NFeXml_Conv_Valor(objICMSST.vBCSTDest)
                                objItem.objTributacao.dICMSvICMSSTDest = NFeXml_Conv_Valor(objICMSST.vICMSSTDest)

                            Case objICMS101.GetType.Name
                                objICMS101 = objICMS.Item
                                objItem.objTributacao.iOrigemMercadoria = objICMS101.orig
                                objItem.objTributacao.sCSOSN = "101"
                                objItem.objTributacao.iICMSSimplesTipo = 1
                                objItem.objTributacao.dICMSpCredSN = NFeXml_Conv_Perc(objICMS101.pCredSN)
                                objItem.objTributacao.dICMSvCredSN = NFeXml_Conv_Valor(objICMS101.vCredICMSSN)

                            Case objICMS102.GetType.Name
                                objICMS102 = objICMS.Item
                                objItem.objTributacao.iOrigemMercadoria = objICMS102.orig
                                objItem.objTributacao.sCSOSN = "102"
                                objItem.objTributacao.iICMSSimplesTipo = 2

                            Case objICMS201.GetType.Name
                                objICMS201 = objICMS.Item
                                objItem.objTributacao.iOrigemMercadoria = objICMS201.orig
                                objItem.objTributacao.sCSOSN = "201"
                                objItem.objTributacao.iICMSSimplesTipo = 4
                                objItem.objTributacao.dICMSpCredSN = NFeXml_Conv_Perc(objICMS201.pCredSN)
                                objItem.objTributacao.dICMSvCredSN = NFeXml_Conv_Valor(objICMS201.vCredICMSSN)
                                objItem.objTributacao.iICMSSubstBaseModalidade = objICMS201.modBCST
                                objItem.objTributacao.dICMSSubstPercMVA = NFeXml_Conv_Perc(objICMS201.pMVAST)
                                objItem.objTributacao.dICMSSubstBase = NFeXml_Conv_Valor(objICMS201.vBCST)
                                objItem.objTributacao.dICMSSubstAliquota = NFeXml_Conv_Perc(objICMS201.pICMSST)
                                objItem.objTributacao.dICMSSubstValor = NFeXml_Conv_Valor(objICMS201.vICMSST)
                                objItem.objTributacao.dICMSSubstPercRedBase = NFeXml_Conv_Perc(objICMS201.pRedBCST)

                            Case objICMS202.GetType.Name
                                objICMS202 = objICMS.Item
                                objItem.objTributacao.iOrigemMercadoria = objICMS202.orig
                                objItem.objTributacao.sCSOSN = "202"
                                objItem.objTributacao.iICMSSimplesTipo = 5
                                objItem.objTributacao.iICMSSubstBaseModalidade = objICMS202.modBCST
                                objItem.objTributacao.dICMSSubstPercMVA = NFeXml_Conv_Perc(objICMS202.pMVAST)
                                objItem.objTributacao.dICMSSubstBase = NFeXml_Conv_Valor(objICMS202.vBCST)
                                objItem.objTributacao.dICMSSubstAliquota = NFeXml_Conv_Perc(objICMS202.pICMSST)
                                objItem.objTributacao.dICMSSubstValor = NFeXml_Conv_Valor(objICMS202.vICMSST)
                                objItem.objTributacao.dICMSSubstPercRedBase = NFeXml_Conv_Perc(objICMS202.pRedBCST)

                            Case objICMS500.GetType.Name
                                objICMS500 = objICMS.Item
                                objItem.objTributacao.iOrigemMercadoria = objICMS500.orig
                                objItem.objTributacao.sCSOSN = "500"
                                objItem.objTributacao.iICMSSimplesTipo = 9
                                objItem.objTributacao.dICMSSTCobrAntBase = NFeXml_Conv_Valor(objICMS500.vBCSTRet)
                                objItem.objTributacao.dICMSSTCobrAntValor = NFeXml_Conv_Valor(objICMS500.vICMSSTRet)

                            Case objICMS900.GetType.Name
                                objICMS900 = objICMS.Item
                                objItem.objTributacao.iOrigemMercadoria = objICMS900.orig
                                objItem.objTributacao.sCSOSN = "900"
                                objItem.objTributacao.iICMSSimplesTipo = 10
                                objItem.objTributacao.dICMSpCredSN = NFeXml_Conv_Perc(objICMS900.pCredSN)
                                objItem.objTributacao.dICMSvCredSN = NFeXml_Conv_Valor(objICMS900.vCredICMSSN)
                                objItem.objTributacao.iICMSBaseModalidade = objICMS900.modBC
                                objItem.objTributacao.dICMSPercRedBase = NFeXml_Conv_Perc(objICMS900.pRedBC)
                                objItem.objTributacao.dICMSBase = NFeXml_Conv_Valor(objICMS900.vBC)
                                objItem.objTributacao.dICMSAliquota = NFeXml_Conv_Perc(objICMS900.pICMS)
                                objItem.objTributacao.dICMSValor = NFeXml_Conv_Valor(objICMS900.vICMS)
                                objItem.objTributacao.iICMSSubstBaseModalidade = objICMS900.modBCST
                                objItem.objTributacao.dICMSSubstPercMVA = NFeXml_Conv_Perc(objICMS900.pMVAST)
                                objItem.objTributacao.dICMSSubstBase = NFeXml_Conv_Valor(objICMS900.vBCST)
                                objItem.objTributacao.dICMSSubstAliquota = NFeXml_Conv_Perc(objICMS900.pICMSST)
                                objItem.objTributacao.dICMSSubstValor = NFeXml_Conv_Valor(objICMS900.vICMSST)
                                objItem.objTributacao.dICMSSubstPercRedBase = NFeXml_Conv_Perc(objICMS900.pRedBCST)
                        End Select

                    End If


                    objItem.objTributacaoItemNF.iIPITipoManual = VAR_PREENCH_AUTOMATICO
                    objItem.objTributacao.iIPITipo = 0
                    objItem.objTributacao.iIPITipoCalculo = TRIB_TIPO_CALCULO_PERCENTUAL
                    objItem.objTributacao.dIPIBaseCalculo = 0
                    objItem.objTributacao.dIPIAliquota = 0

                    If Not objPIS Is Nothing Then

                        objPISAliq = New TNFeInfNFeDetImpostoPISPISAliq
                        objPISQtde = New TNFeInfNFeDetImpostoPISPISQtde
                        objPISNT = New TNFeInfNFeDetImpostoPISPISNT
                        objPISOut = New TNFeInfNFeDetImpostoPISPISOutr

                        If objPIS.Item.GetType.Name = objPISAliq.GetType.Name Then

                            objPISAliq = objPIS.Item

                            Select Case objPISAliq.CST
                                Case TNFeInfNFeDetImpostoPISPISAliqCST.Item01
                                    objItem.objTributacao.iPISTipo = 1
                                Case TNFeInfNFeDetImpostoPISPISAliqCST.Item02
                                    objItem.objTributacao.iPISTipo = 2
                            End Select

                            objItem.objTributacao.iPISTipoCalculo = TRIB_TIPO_CALCULO_PERCENTUAL
                            objItem.objTributacao.dPISAliquota = NFeXml_Conv_Perc(objPISAliq.pPIS)
                            objItem.objTributacao.dPISValor = NFeXml_Conv_Valor(objPISAliq.vPIS)
                            objItem.objTributacao.dPISBase = NFeXml_Conv_Valor(objPISAliq.vBC)

                        ElseIf objPIS.Item.GetType.Name = objPISQtde.GetType.Name Then

                            objPISQtde = objPIS.Item

                            Select Case objPISQtde.CST
                                Case TNFeInfNFeDetImpostoPISPISQtdeCST.Item03
                                    objItem.objTributacao.iPISTipo = 3
                            End Select

                            objItem.objTributacao.iPISTipoCalculo = TRIB_TIPO_CALCULO_VALOR
                            objItem.objTributacao.dPISQtde = NFeXml_Conv_Valor(objPISQtde.qBCProd)
                            objItem.objTributacao.dPISAliquotaValor = NFeXml_Conv_Valor(objPISQtde.vAliqProd)
                            objItem.objTributacao.dPISValor = NFeXml_Conv_Valor(objPISQtde.vPIS)

                        ElseIf objPIS.Item.GetType.Name = objPISNT.GetType.Name Then

                            objPISNT = objPIS.Item

                            Select Case objPISNT.CST
                                Case TNFeInfNFeDetImpostoPISPISNTCST.Item04
                                    objItem.objTributacao.iPISTipo = 4
                                Case TNFeInfNFeDetImpostoPISPISNTCST.Item06
                                    objItem.objTributacao.iPISTipo = 6
                                Case TNFeInfNFeDetImpostoPISPISNTCST.Item07
                                    objItem.objTributacao.iPISTipo = 7
                                Case TNFeInfNFeDetImpostoPISPISNTCST.Item08
                                    objItem.objTributacao.iPISTipo = 8
                                Case TNFeInfNFeDetImpostoPISPISNTCST.Item09
                                    objItem.objTributacao.iPISTipo = 9
                            End Select

                        ElseIf objPIS.Item.GetType.Name = objPISOut.GetType.Name Then

                            objPISOut = objPIS.Item

                            Select Case objPISOut.CST
                                Case TNFeInfNFeDetImpostoPISPISOutrCST.Item49
                                    objItem.objTributacao.iPISTipo = 49
                                Case TNFeInfNFeDetImpostoPISPISOutrCST.Item50
                                    objItem.objTributacao.iPISTipo = 50
                                Case TNFeInfNFeDetImpostoPISPISOutrCST.Item51
                                    objItem.objTributacao.iPISTipo = 51
                                Case TNFeInfNFeDetImpostoPISPISOutrCST.Item52
                                    objItem.objTributacao.iPISTipo = 52
                                Case TNFeInfNFeDetImpostoPISPISOutrCST.Item53
                                    objItem.objTributacao.iPISTipo = 53
                                Case TNFeInfNFeDetImpostoPISPISOutrCST.Item54
                                    objItem.objTributacao.iPISTipo = 54
                                Case TNFeInfNFeDetImpostoPISPISOutrCST.Item55
                                    objItem.objTributacao.iPISTipo = 55
                                Case TNFeInfNFeDetImpostoPISPISOutrCST.Item56
                                    objItem.objTributacao.iPISTipo = 56
                                Case TNFeInfNFeDetImpostoPISPISOutrCST.Item60
                                    objItem.objTributacao.iPISTipo = 60
                                Case TNFeInfNFeDetImpostoPISPISOutrCST.Item61
                                    objItem.objTributacao.iPISTipo = 61
                                Case TNFeInfNFeDetImpostoPISPISOutrCST.Item62
                                    objItem.objTributacao.iPISTipo = 62
                                Case TNFeInfNFeDetImpostoPISPISOutrCST.Item63
                                    objItem.objTributacao.iPISTipo = 63
                                Case TNFeInfNFeDetImpostoPISPISOutrCST.Item64
                                    objItem.objTributacao.iPISTipo = 64
                                Case TNFeInfNFeDetImpostoPISPISOutrCST.Item65
                                    objItem.objTributacao.iPISTipo = 65
                                Case TNFeInfNFeDetImpostoPISPISOutrCST.Item66
                                    objItem.objTributacao.iPISTipo = 66
                                Case TNFeInfNFeDetImpostoPISPISOutrCST.Item67
                                    objItem.objTributacao.iPISTipo = 67
                                Case TNFeInfNFeDetImpostoPISPISOutrCST.Item70
                                    objItem.objTributacao.iPISTipo = 70
                                Case TNFeInfNFeDetImpostoPISPISOutrCST.Item71
                                    objItem.objTributacao.iPISTipo = 71
                                Case TNFeInfNFeDetImpostoPISPISOutrCST.Item72
                                    objItem.objTributacao.iPISTipo = 72
                                Case TNFeInfNFeDetImpostoPISPISOutrCST.Item73
                                    objItem.objTributacao.iPISTipo = 73
                                Case TNFeInfNFeDetImpostoPISPISOutrCST.Item74
                                    objItem.objTributacao.iPISTipo = 74
                                Case TNFeInfNFeDetImpostoPISPISOutrCST.Item75
                                    objItem.objTributacao.iPISTipo = 75
                                Case TNFeInfNFeDetImpostoPISPISOutrCST.Item98
                                    objItem.objTributacao.iPISTipo = 98
                                Case TNFeInfNFeDetImpostoPISPISOutrCST.Item99
                                    objItem.objTributacao.iPISTipo = 99
                            End Select

                            Select Case objItem.objTributacao.iPISTipo
                                Case 49 To 56, 60 To 67, 70 To 75, 98, 99
                                    If objPISOut.ItemsElementName(0) = ItemsChoiceType1.vBC Then
                                        objItem.objTributacao.iPISTipoCalculo = TRIB_TIPO_CALCULO_PERCENTUAL
                                        objItem.objTributacao.dPISBase = NFeXml_Conv_Valor(objPISOut.Items(0))
                                        objItem.objTributacao.dPISAliquota = NFeXml_Conv_Perc(objPISOut.Items(1))
                                    Else
                                        objItem.objTributacao.iPISTipoCalculo = TRIB_TIPO_CALCULO_VALOR
                                        objItem.objTributacao.dPISQtde = NFeXml_Conv_Valor(objPISOut.Items(0))
                                        objItem.objTributacao.dPISAliquotaValor = NFeXml_Conv_Valor(objPISOut.Items(1))
                                    End If

                                    objItem.objTributacao.dPISValor = NFeXml_Conv_Valor(objPISOut.vPIS)

                            End Select

                        End If

                        If Not (objPISST Is Nothing) Then
                            If objPISST.ItemsElementName(0) = ItemsChoiceType2.vBC Then
                                objItem.objTributacao.iPISSTTipoCalculo = TRIB_TIPO_CALCULO_PERCENTUAL
                                objItem.objTributacao.dPISSTBase = NFeXml_Conv_Valor(objPISST.Items(0))
                                objItem.objTributacao.dPISSTAliquota = NFeXml_Conv_Perc(objPISST.Items(1))
                            Else
                                objItem.objTributacao.iPISSTTipoCalculo = TRIB_TIPO_CALCULO_VALOR
                                objItem.objTributacao.dPISSTQtde = NFeXml_Conv_Valor(objPISST.Items(0))
                                objItem.objTributacao.dPISSTAliquotaValor = NFeXml_Conv_Valor(objPISST.Items(1))
                            End If

                            objItem.objTributacao.dPISSTValor = NFeXml_Conv_Valor(objPISST.vPIS)
                        End If

                    Else

                        'a parte de PIS nao estava no xml
                        objItem.objTributacaoItemNF.iPISTipoManual = VAR_PREENCH_AUTOMATICO
                        objItem.objTributacao.iPISTipo = 99
                        objItem.objTributacao.iPISTipoCalculo = TRIB_TIPO_CALCULO_PERCENTUAL
                        objItem.objTributacao.iPISSTTipoCalculo = TRIB_TIPO_CALCULO_PERCENTUAL

                    End If

                    If Not objCOFINS Is Nothing Then

                        objCOFINSAliq = New TNFeInfNFeDetImpostoCOFINSCOFINSAliq
                        objCOFINSQtde = New TNFeInfNFeDetImpostoCOFINSCOFINSQtde
                        objCOFINSNT = New TNFeInfNFeDetImpostoCOFINSCOFINSNT
                        objCOFINSOut = New TNFeInfNFeDetImpostoCOFINSCOFINSOutr

                        If objCOFINS.Item.GetType.Name = objCOFINSAliq.GetType.Name Then

                            objCOFINSAliq = objCOFINS.Item

                            Select Case objCOFINSAliq.CST
                                Case TNFeInfNFeDetImpostoCOFINSCOFINSAliqCST.Item01
                                    objItem.objTributacao.iCOFINSTipo = 1
                                Case TNFeInfNFeDetImpostoCOFINSCOFINSAliqCST.Item02
                                    objItem.objTributacao.iCOFINSTipo = 2
                            End Select

                            objItem.objTributacao.iCOFINSTipoCalculo = TRIB_TIPO_CALCULO_PERCENTUAL
                            objItem.objTributacao.dCOFINSAliquota = NFeXml_Conv_Perc(objCOFINSAliq.pCOFINS)
                            objItem.objTributacao.dCOFINSValor = NFeXml_Conv_Valor(objCOFINSAliq.vCOFINS)
                            objItem.objTributacao.dCOFINSBase = NFeXml_Conv_Valor(objCOFINSAliq.vBC)

                        ElseIf objCOFINS.Item.GetType.Name = objCOFINSQtde.GetType.Name Then

                            objCOFINSQtde = objCOFINS.Item

                            Select Case objCOFINSQtde.CST
                                Case TNFeInfNFeDetImpostoCOFINSCOFINSQtdeCST.Item03
                                    objItem.objTributacao.iCOFINSTipo = 3
                            End Select

                            objItem.objTributacao.iCOFINSTipoCalculo = TRIB_TIPO_CALCULO_VALOR
                            objItem.objTributacao.dCOFINSQtde = NFeXml_Conv_Valor(objCOFINSQtde.qBCProd)
                            objItem.objTributacao.dCOFINSAliquotaValor = NFeXml_Conv_Valor(objCOFINSQtde.vAliqProd)
                            objItem.objTributacao.dCOFINSValor = NFeXml_Conv_Valor(objCOFINSQtde.vCOFINS)

                        ElseIf objCOFINS.Item.GetType.Name = objCOFINSNT.GetType.Name Then

                            objCOFINSNT = objCOFINS.Item

                            Select Case objCOFINSNT.CST
                                Case TNFeInfNFeDetImpostoCOFINSCOFINSNTCST.Item04
                                    objItem.objTributacao.iCOFINSTipo = 4
                                Case TNFeInfNFeDetImpostoCOFINSCOFINSNTCST.Item06
                                    objItem.objTributacao.iCOFINSTipo = 6
                                Case TNFeInfNFeDetImpostoCOFINSCOFINSNTCST.Item07
                                    objItem.objTributacao.iCOFINSTipo = 7
                                Case TNFeInfNFeDetImpostoCOFINSCOFINSNTCST.Item08
                                    objItem.objTributacao.iCOFINSTipo = 8
                                Case TNFeInfNFeDetImpostoCOFINSCOFINSNTCST.Item09
                                    objItem.objTributacao.iCOFINSTipo = 9
                            End Select

                        ElseIf objCOFINS.Item.GetType.Name = objCOFINSOut.GetType.Name Then

                            objCOFINSOut = objCOFINS.Item

                            Select Case objCOFINSOut.CST
                                Case TNFeInfNFeDetImpostoCOFINSCOFINSOutrCST.Item49
                                    objItem.objTributacao.iCOFINSTipo = 49
                                Case TNFeInfNFeDetImpostoCOFINSCOFINSOutrCST.Item50
                                    objItem.objTributacao.iCOFINSTipo = 50
                                Case TNFeInfNFeDetImpostoCOFINSCOFINSOutrCST.Item51
                                    objItem.objTributacao.iCOFINSTipo = 51
                                Case TNFeInfNFeDetImpostoCOFINSCOFINSOutrCST.Item52
                                    objItem.objTributacao.iCOFINSTipo = 52
                                Case TNFeInfNFeDetImpostoCOFINSCOFINSOutrCST.Item53
                                    objItem.objTributacao.iCOFINSTipo = 53
                                Case TNFeInfNFeDetImpostoCOFINSCOFINSOutrCST.Item54
                                    objItem.objTributacao.iCOFINSTipo = 54
                                Case TNFeInfNFeDetImpostoCOFINSCOFINSOutrCST.Item55
                                    objItem.objTributacao.iCOFINSTipo = 55
                                Case TNFeInfNFeDetImpostoCOFINSCOFINSOutrCST.Item56
                                    objItem.objTributacao.iCOFINSTipo = 56
                                Case TNFeInfNFeDetImpostoCOFINSCOFINSOutrCST.Item60
                                    objItem.objTributacao.iCOFINSTipo = 60
                                Case TNFeInfNFeDetImpostoCOFINSCOFINSOutrCST.Item61
                                    objItem.objTributacao.iCOFINSTipo = 61
                                Case TNFeInfNFeDetImpostoCOFINSCOFINSOutrCST.Item62
                                    objItem.objTributacao.iCOFINSTipo = 62
                                Case TNFeInfNFeDetImpostoCOFINSCOFINSOutrCST.Item63
                                    objItem.objTributacao.iCOFINSTipo = 63
                                Case TNFeInfNFeDetImpostoCOFINSCOFINSOutrCST.Item64
                                    objItem.objTributacao.iCOFINSTipo = 64
                                Case TNFeInfNFeDetImpostoCOFINSCOFINSOutrCST.Item65
                                    objItem.objTributacao.iCOFINSTipo = 65
                                Case TNFeInfNFeDetImpostoCOFINSCOFINSOutrCST.Item66
                                    objItem.objTributacao.iCOFINSTipo = 66
                                Case TNFeInfNFeDetImpostoCOFINSCOFINSOutrCST.Item67
                                    objItem.objTributacao.iCOFINSTipo = 67
                                Case TNFeInfNFeDetImpostoCOFINSCOFINSOutrCST.Item70
                                    objItem.objTributacao.iCOFINSTipo = 70
                                Case TNFeInfNFeDetImpostoCOFINSCOFINSOutrCST.Item71
                                    objItem.objTributacao.iCOFINSTipo = 71
                                Case TNFeInfNFeDetImpostoCOFINSCOFINSOutrCST.Item72
                                    objItem.objTributacao.iCOFINSTipo = 72
                                Case TNFeInfNFeDetImpostoCOFINSCOFINSOutrCST.Item73
                                    objItem.objTributacao.iCOFINSTipo = 73
                                Case TNFeInfNFeDetImpostoCOFINSCOFINSOutrCST.Item74
                                    objItem.objTributacao.iCOFINSTipo = 74
                                Case TNFeInfNFeDetImpostoCOFINSCOFINSOutrCST.Item75
                                    objItem.objTributacao.iCOFINSTipo = 75
                                Case TNFeInfNFeDetImpostoCOFINSCOFINSOutrCST.Item98
                                    objItem.objTributacao.iCOFINSTipo = 98
                                Case TNFeInfNFeDetImpostoCOFINSCOFINSOutrCST.Item99
                                    objItem.objTributacao.iCOFINSTipo = 99
                            End Select

                            Select Case objItem.objTributacao.iCOFINSTipo
                                Case 49 To 56, 60 To 67, 70 To 75, 98, 99
                                    If objCOFINSOut.ItemsElementName(0) = ItemsChoiceType1.vBC Then
                                        objItem.objTributacao.iCOFINSTipoCalculo = TRIB_TIPO_CALCULO_PERCENTUAL
                                        objItem.objTributacao.dCOFINSBase = NFeXml_Conv_Valor(objCOFINSOut.Items(0))
                                        objItem.objTributacao.dCOFINSAliquota = NFeXml_Conv_Perc(objCOFINSOut.Items(1))
                                    Else
                                        objItem.objTributacao.iCOFINSTipoCalculo = TRIB_TIPO_CALCULO_VALOR
                                        objItem.objTributacao.dCOFINSQtde = NFeXml_Conv_Valor(objCOFINSOut.Items(0))
                                        objItem.objTributacao.dCOFINSAliquotaValor = NFeXml_Conv_Valor(objCOFINSOut.Items(1))
                                    End If

                                    objItem.objTributacao.dCOFINSValor = NFeXml_Conv_Valor(objCOFINSOut.vCOFINS)

                            End Select

                        End If

                        If Not (objCOFINSST Is Nothing) Then
                            If objCOFINSST.ItemsElementName(0) = ItemsChoiceType2.vBC Then
                                objItem.objTributacao.iCOFINSSTTipoCalculo = TRIB_TIPO_CALCULO_PERCENTUAL
                                objItem.objTributacao.dCOFINSSTBase = NFeXml_Conv_Valor(objCOFINSST.Items(0))
                                objItem.objTributacao.dCOFINSSTAliquota = NFeXml_Conv_Perc(objCOFINSST.Items(1))
                            Else
                                objItem.objTributacao.iCOFINSSTTipoCalculo = TRIB_TIPO_CALCULO_VALOR
                                objItem.objTributacao.dCOFINSSTQtde = NFeXml_Conv_Valor(objCOFINSST.Items(0))
                                objItem.objTributacao.dCOFINSSTAliquotaValor = NFeXml_Conv_Valor(objCOFINSST.Items(1))
                            End If

                            objItem.objTributacao.dCOFINSSTValor = NFeXml_Conv_Valor(objCOFINSST.vCOFINS)
                        End If

                    Else

                        'a parte de COFINS nao estava no xml
                        objItem.objTributacaoItemNF.iCOFINSTipoManual = VAR_PREENCH_AUTOMATICO
                        objItem.objTributacao.iCOFINSTipo = 99
                        objItem.objTributacao.iCOFINSTipoCalculo = TRIB_TIPO_CALCULO_PERCENTUAL
                        objItem.objTributacao.iCOFINSSTTipoCalculo = TRIB_TIPO_CALCULO_PERCENTUAL

                    End If

                    objItem.iItem = NFeXml_Conv_Long(objNFe.infNFe.det(iIndice - 1).nItem)
                    objItem.sProduto = NFeXml_Conv_Texto(objNFe.infNFe.det(iIndice - 1).prod.cProd)
                    objItem.sDescricaoItem = NFeXml_Conv_Texto(objNFe.infNFe.det(iIndice - 1).prod.xProd)
                    objItem.objTributacao.sNaturezaOp = Replace(NFeXml_Conv_Texto(objNFe.infNFe.det(iIndice - 1).prod.CFOP.ToString), "Item", "")
                    If iIndice = 1 Then
                        objNF.sNaturezaOp = objItem.objTributacao.sNaturezaOp
                        objNF.objTributacaoNF.sNaturezaOp = objNF.sNaturezaOp
                        objNF.objTributacaoNF.sNaturezaOpInterna = objNF.sNaturezaOp
                    End If

                    objItem.sUnidadeMed = NFeXml_Conv_Texto(objNFe.infNFe.det(iIndice - 1).prod.uCom)
                    objItem.dQuantidade = NFeXml_Conv_Valor(objNFe.infNFe.det(iIndice - 1).prod.qCom)
                    objItem.dPrecoUnitario = NFeXml_Conv_Valor(objNFe.infNFe.det(iIndice - 1).prod.vUnCom)
                    objItem.dValorTotal = NFeXml_Conv_Valor(objNFe.infNFe.det(iIndice - 1).prod.vProd)
                    objItem.sEAN = NFeXml_Conv_Texto(objNFe.infNFe.det(iIndice - 1).prod.cEAN)
                    objItem.objInfoAdicDocItem.sMsg = Left(NFeXml_Conv_Texto(objNFe.infNFe.det(iIndice - 1).infAdProd), 250)
                    objItem.objTributacao.dValorFreteItem = NFeXml_Conv_Valor(objNFe.infNFe.det(iIndice - 1).prod.vFrete)
                    objItem.objTributacao.dValorSeguroItem = NFeXml_Conv_Valor(objNFe.infNFe.det(iIndice - 1).prod.vSeg)
                    objItem.objTributacao.dDescontoGrid = NFeXml_Conv_Valor(objNFe.infNFe.det(iIndice - 1).prod.vDesc)
                    objItem.objTributacao.dValorOutrasDespesasItem = NFeXml_Conv_Valor(objNFe.infNFe.det(iIndice - 1).prod.vOutro)
                    objItem.objTributacao.sIPICodProduto = NFeXml_Conv_Texto(objNFe.infNFe.det(iIndice - 1).prod.NCM)

                    objItem.objTributacao.sUMTrib = NFeXml_Conv_Texto(objNFe.infNFe.det(iIndice - 1).prod.uTrib)
                    objItem.objTributacao.dQtdTrib = NFeXml_Conv_Valor(objNFe.infNFe.det(iIndice - 1).prod.qTrib)
                    objItem.objTributacao.dValorUnitTrib = NFeXml_Conv_Valor(objNFe.infNFe.det(iIndice - 1).prod.vUnTrib)
                    objItem.objTributacao.sEANTrib = NFeXml_Conv_Texto(objNFe.infNFe.det(iIndice - 1).prod.cEANTrib)
                    objItem.objInfoAdicDocItem.iIncluiValorTotal = NFeXml_Conv_Long(objNFe.infNFe.det(iIndice - 1).prod.indTot)
                    objItem.objTributacao.dTotTrib = NFeXml_Conv_Valor(objNFe.infNFe.det(iIndice - 1).imposto.vTotTrib)

            Next

            If Not (objNFe.infNFeSupl Is Nothing) Then
                objNF.sQRCode = objNFe.infNFeSupl.qrCode
            End If

            NFe_Carrega = SUCESSO

        Catch ex As Exception

            Dim sMsg2 As String

            If ex.InnerException Is Nothing Then
                sMsg2 = ""
            Else
                sMsg2 = " - " & ex.InnerException.Message
            End If

            MsgBox("Verifique a mensagem de erro: " & ex.Message & sMsg2)

            NFe_Carrega = 1

        Finally

        End Try

    End Function

End Class


