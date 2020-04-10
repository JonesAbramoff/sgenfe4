Option Strict Off
Option Explicit On

Imports System
Imports System.Xml.Serialization
Imports System.IO
Imports System.Text
Imports System.Security.Cryptography.Xml
Imports System.Security.Cryptography.X509Certificates
Imports System.Xml
Imports System.Xml.Schema
Imports System.Data.Odbc
Imports Microsoft.Win32
Imports System.Net
Imports sgenfe4.NFeXsd

Public Class ClassGlobalApp

    Public Form1 As New Form1
    Public iDebug As Integer
    Public sErro As String
    Public sMsg1 As String

    Public sCertificado As String
    Public iNFeAmbiente As Integer
    Public cert As New X509Certificate2
    Public gobjVenda As GlobaisLoja.ClassVenda

    Public iUFCodIBGE As Integer
    Public sSiglaEstado As String
    Public sDirXml As String
    Public sDirXsd As String

    Public Sub GravarLog(ByVal sTexto As String, Optional ByVal sTipo As String = "INFO", Optional ByVal bGravaForm As Boolean = True, Optional ByVal bGravaECFLog As Boolean = True, Optional ByVal bGravaNFCELog As Boolean = True, Optional ByVal bGravaTxt As Boolean = False, Optional ByVal sNomeTxt As String = "")

        If bGravaForm Then Form1.Msg.Items.Add(sTexto)
        'System.Windows.Forms.Application.DoEvents()
        If bGravaECFLog Then ECF_Grava_Log(sTexto, sTipo)

        If bGravaNFCELog Then NFCE_Grava_Log(sTexto, sTipo)

        If bGravaTxt Then TXT_Grava_Log(sTexto, sNomeTxt)

    End Sub

    Private Function ECF_Grava_Log(ByVal sTexto As String, Optional ByVal sTipo As String = "INFO") As Long

        Dim bAbriuArq As Boolean
        Dim objSW As System.IO.StreamWriter

        On Error GoTo Erro_ECF_Grava_Log

        bAbriuArq = False
        If Not File.Exists(Path.Combine(CurDir, "ECFLog.txt")) Then
            objSW = File.CreateText(Path.Combine(CurDir, "ECFLog.txt"))
            objSW.Close()
        End If
        objSW = New StreamWriter(Path.Combine(CurDir, "ECFLog.txt"), True)
        bAbriuArq = True

        sTexto = Replace(sTexto, vbNewLine, "-")
        sTexto = Replace(sTexto, Chr(10), "")
        sTexto = Replace(sTexto, Chr(13), "")
        sTexto = Replace(sTexto, Chr(0), "")
        sTexto = Trim(sTexto)

        objSW.WriteLine(Now().ToString & " |" & sTipo & "| " & sTexto)

        objSW.Close()

        ECF_Grava_Log = SUCESSO

        Exit Function

Erro_ECF_Grava_Log:

        If bAbriuArq Then objSW.Close()

        ECF_Grava_Log = Err.Number

        Exit Function

    End Function

    Private Function NFCE_Grava_Log(ByVal sTexto As String, Optional ByVal sTipo As String = "INFO") As Long

        Dim bAbriuArq As Boolean
        Dim objSW As System.IO.StreamWriter

        On Error GoTo Erro_NFCE_Grava_Log

        bAbriuArq = False
        If Not File.Exists(Path.Combine(CurDir, "NFCELog.txt")) Then
            objSW = File.CreateText(Path.Combine(CurDir, "NFCELog.txt"))
            objSW.Close()
        End If
        objSW = New StreamWriter(Path.Combine(CurDir, "NFCELog.txt"), True)
        bAbriuArq = True

        sTexto = Replace(sTexto, vbNewLine, "-")
        sTexto = Replace(sTexto, Chr(10), "")
        sTexto = Replace(sTexto, Chr(13), "")
        sTexto = Replace(sTexto, Chr(0), "")
        sTexto = Trim(sTexto)

        objSW.WriteLine(Now().ToString & " |" & sTipo & "| " & sTexto)

        objSW.Close()

        NFCE_Grava_Log = SUCESSO

        Exit Function

Erro_NFCE_Grava_Log:

        If bAbriuArq Then objSW.Close()

        NFCE_Grava_Log = Err.Number

        Exit Function

    End Function

    Private Function TXT_Grava_Log(ByVal sTexto As String, ByVal sNomeArq As String) As Long

        Dim bAbriuArq As Boolean
        Dim objSW As System.IO.StreamWriter

        On Error GoTo Erro_TXT_Grava_Log

        bAbriuArq = False
        objSW = File.CreateText(Path.Combine(sDirXml, sNomeArq))
        bAbriuArq = True

        objSW.Write(sTexto)

        objSW.Close()

        TXT_Grava_Log = SUCESSO

        Exit Function

Erro_TXT_Grava_Log:

        If bAbriuArq Then objSW.Close()

        TXT_Grava_Log = Err.Number

        Exit Function

    End Function

    Public Function Obtem_Certificado() As Long
        '
        '  seleciona certificado do repositório MY do windows
        '
        Try

            Dim certificado As Certificado = New Certificado
            Dim lErro As Long

            lErro = certificado.BuscaNome(gobjVenda.objNFeInfo.sCertificadoA1A3, cert)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            If DateDiff(DateInterval.Day, Now.Date, cert.NotAfter) < 15 And DateDiff(DateInterval.Day, Now.Date, cert.NotAfter) >= 0 Then
                Form1.Msg.Items.Add("ATENÇÃO: FALTAM " & DateDiff(DateInterval.Day, Now.Date, cert.NotAfter) & " DIAS PARA TERMINAR A VALIDADE DO SEU CERTIFICADO. FAVOR RENOVA-LO.")

            ElseIf DateDiff(DateInterval.Day, Now.Date, cert.NotAfter) < 0 Then

                Throw New System.Exception("ATENÇÃO: O CERTIFICADO ESTÁ COM O PRAZO DE VALIDADE VENCIDO. FAVOR RENOVA-LO.")

            End If

            Obtem_Certificado = SUCESSO

        Catch ex As Exception

            Obtem_Certificado = 1

            Form1.Msg.Items.Add("Erro na seleção do certificado digital. " & ex.Message)

        End Try

    End Function

    Private Function Verifica_Status_Servico5(ByVal DocDados1 As XmlDocument, ByVal NfeStatusServico As nfestatusservico2.NFeStatusServico4) As Long

        Dim XMLStringRetStatServ As String
        Dim objRetStatServ As TRetConsStatServ = New TRetConsStatServ
        Dim XMLStreamRet As MemoryStream = New MemoryStream(10000)
        Dim xmlNode1 As XmlNode
        Dim xRet As Byte()

        Try

            If iDebug = 1 Then MsgBox("39.4")
            sErro = "39.4"
            sMsg1 = "vai fazer consulta a Status do Servico"

            NfeStatusServico.ClientCertificates.Add(cert)
            xmlNode1 = NfeStatusServico.nfeStatusServicoNF(DocDados1)

            XMLStringRetStatServ = xmlNode1.OuterXml

            If iDebug = 1 Then
                MsgBox("39.5")
                MsgBox(XMLStringRetStatServ)
            End If
            sErro = "39.5"
            sMsg1 = "consultou o Status do Servico"

            xRet = System.Text.Encoding.UTF8.GetBytes(XMLStringRetStatServ)

            XMLStreamRet = New MemoryStream(10000)
            XMLStreamRet.Write(xRet, 0, xRet.Length)

            Dim mySerializerRetConsStatServ1 As New XmlSerializer(GetType(TRetConsStatServ))

            Dim objRetConsStatServ1 As TRetConsStatServ = New TRetConsStatServ

            XMLStreamRet.Position = 0

            objRetConsStatServ1 = mySerializerRetConsStatServ1.Deserialize(XMLStreamRet)

            If objRetConsStatServ1.cStat <> "107" Then
                Form1.Msg.Items.Add(XMLStringRetStatServ)
            End If

            Verifica_Status_Servico5 = SUCESSO

        Catch ex As Exception

            Verifica_Status_Servico5 = -1

        End Try

    End Function

    Private Sub Verifica_Status_Servico2(ByVal DocDados1 As XmlDocument, ByVal NfeStatusServico As nfestatusservico2.NFeStatusServico4)
        'Obtem URL do web service de status do serviço

        Dim sURL As String
        sURL = ""
        Call WS_Obter_URL(sURL, iNFeAmbiente = NFE_AMBIENTE_HOMOLOGACAO, sSiglaEstado, "NfeStatusServico", "NFCe")

        NfeStatusServico.Url = sURL

        'Dim NFecabec_StatusServ As New nfestatusservico2.nfeCabecMsg

        'NFecabec_StatusServ.cUF = CStr(iUFCodIBGE)
        'NFecabec_StatusServ.versaoDados = NFE_VERSAO_XML

        'NfeStatusServico.nfeCabecMsgValue = NFecabec_StatusServ

    End Sub

    Private Sub Verifica_Status_Servico1(ByVal objStatServ As TConsStatServ)

        objStatServ.cUF = GetCode(Of TCodUfIBGE)(CStr(iUFCodIBGE))

        objStatServ.versao = NFE_VERSAO_XML

        If iNFeAmbiente = NFE_AMBIENTE_HOMOLOGACAO Then
            objStatServ.tpAmb = TAmb.Item2
        Else
            objStatServ.tpAmb = TAmb.Item1
        End If

    End Sub

    Private Sub Verifica_Status_Servico4(ByVal DocDados1 As XmlDocument, ByVal objStatServ As TConsStatServ)
        'salva arquivo

        Dim XMLString4 As String
        Dim iPos As Integer
        Dim XMLStream As MemoryStream = New MemoryStream(10000)

        objStatServ.xServ = TConsStatServXServ.STATUS

        Dim mySerializerZ As New XmlSerializer(GetType(TConsStatServ))

        XMLStream = New MemoryStream(10000)

        mySerializerZ.Serialize(XMLStream, objStatServ)

        Dim xmz As Byte()
        xmz = XMLStream.ToArray

        XMLString4 = System.Text.Encoding.UTF8.GetString(xmz)

        XMLString4 = Mid(XMLString4, 1, 19) & " encoding=""UTF-8"" " & Mid(XMLString4, 20)

        Form1.Msg.Items.Add("Iniciando a verificação do status do serviço")

        '****************  salva o arquivo 

        iPos = InStr(XMLString4, "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema""")


        If iPos <> 0 Then

            XMLString4 = Mid(XMLString4, 1, iPos - 1) & Mid(XMLString4, iPos + 99)

        End If

        Call Salva_Arquivo(DocDados1, XMLString4)

    End Sub

    Public Function Verifica_Status_Servico() As Long

        Dim objStatServ As TConsStatServ = New TConsStatServ
        Dim DocDados1 As XmlDocument = New XmlDocument
        Dim lErro As Long
        Dim NfeStatusServico As New nfestatusservico2.NFeStatusServico4

        Call Verifica_Status_Servico1(objStatServ)

        Call Verifica_Status_Servico4(DocDados1, objStatServ)

        Call Verifica_Status_Servico2(DocDados1, NfeStatusServico)

        lErro = Verifica_Status_Servico5(DocDados1, NfeStatusServico)

        Verifica_Status_Servico = lErro

    End Function

    Public Function XML_Valida(ByVal sArquivo As String, ByVal sXSD As String) As Long
        Dim objValidaXML As ClassValidaXML = New ClassValidaXML
        Dim lErro As Long

        If iDebug = 1 Then MsgBox("39.1")
        sErro = "39.1"
        sMsg1 = "vai validar o XML=" & sArquivo & " xsd=" & sXSD

        Try

            lErro = objValidaXML.validaXML(sArquivo, gobjVenda.objNFeInfo.sDirXsd & sXSD, Form1)
            If lErro <> SUCESSO Then

                GravarLog("ERRO - xml inválido de acordo com os xsd.")

                Throw New System.Exception("")

            End If

            If iDebug = 1 Then MsgBox("39.2")
            sErro = "39.2"
            sMsg1 = "validou o XML"

        Catch ex As Exception

            If lErro = SUCESSO Then lErro = 1

        Finally

            XML_Valida = lErro

        End Try

    End Function

    Public Sub New(ByVal objVenda As GlobaisLoja.ClassVenda, ByVal sModelo As String)

        gobjVenda = objVenda

        If sModelo = "NFe" Then
            iNFeAmbiente = objVenda.objNFeInfo.iNFeAmbiente
        Else
            iNFeAmbiente = objVenda.objNFeInfo.iNFCeAmbiente
        End If

        If UCase(objVenda.objNFeInfo.sCertificadoA1A3) = "FORPRINT" Then
            iNFeAmbiente = NFE_AMBIENTE_HOMOLOGACAO
        End If

        sDirXml = objVenda.objNFeInfo.sDirArqXml
        sSiglaEstado = objVenda.objNFeInfo.sEmitUF
        Call UF_ObterCodIBGE(sSiglaEstado, iUFCodIBGE)
        '        iUFCodIBGE = CInt(GetXmlAttrNameFromEnumValue(Of TCodUfIBGE)(GetCode(Of TUf)(sSiglaEstado)))

        'Form1.Show()

    End Sub

    Public Sub Terminar()

        Form1.Close()
        Form1.Dispose()
        Form1 = Nothing
        gobjVenda = Nothing
        cert = Nothing

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()

    End Sub
End Class
