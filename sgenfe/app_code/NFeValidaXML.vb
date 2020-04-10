Imports System
Imports System.Xml.Serialization
Imports System.IO
Imports System.Text
Imports System.Security.Cryptography.Xml
Imports System.Security.Cryptography.X509Certificates
Imports System.Xml
Imports System.Xml.Schema


Public Class ClassValidaXML

    Dim gsMsg As String
    Dim glErro As Long

    Public Function validaXML(ByVal _arquivo As String, ByVal _schema As String, ByVal objAvisoCtl As Object) As Long

        ' Create a new validating reader

        Dim reader As XmlValidatingReader = New XmlValidatingReader(New XmlTextReader(New StreamReader(_arquivo)))
        'Dim reader As XmlValidatingReader = New XmlValidatingReader()
        '        Dim reader1 As XmlWriter
        '       reader1.
        Dim schema(1) As System.Xml.Schema.XmlSchema

        '// Create a schema collection, add the xsd to it

        Dim schemaCollection As XmlSchemaSet = New XmlSchemaSet()
        Dim iLinha As Integer

        Try

            NFCE_Grava_Log("Validando: " & _arquivo & " Contra: " & _schema)

            glErro = SUCESSO

            schemaCollection.Add("http://www.portalfiscal.inf.br/nfe", _schema)

            schemaCollection.CopyTo(schema, 0)

            '// Add the schema collection to the XmlValidatingReader

            reader.Schemas.Add(schema(0))

            '       Console.Write("Início da validação...\n")

            '    // Wire up the call back.  The ValidationEvent is fired when the
            '    // XmlValidatingReader hits an issue validating a section of the xml

            '            reader. += new ValidationEventHandler(reader_ValidationEventHandler);
            AddHandler reader.ValidationEventHandler, AddressOf reader_ValidationEventHandler

            '            // Iterate through the xml document



            '            while (reader.Read()) {}

            iLinha = 0

            While reader.Read()

                iLinha = iLinha + 1
                If Len(Trim(gsMsg)) > 0 Then

                    objAvisoCtl.Msg.Items.Add(" Linha = " & iLinha & gsMsg)
                    NFCE_Grava_Log(" Linha = " & iLinha & gsMsg)

                    gsMsg = ""

                End If

            End While


        Catch ex As Exception

            Dim sMsg As String

            If ex.InnerException Is Nothing Then
                sMsg = ""
            Else
                sMsg = " - " & ex.InnerException.Message
            End If

            objAvisoCtl.Msg.Items.Add(ex.Message & sMsg)
            NFCE_Grava_Log(ex.Message & sMsg)
            objAvisoCtl.Msg.Items.Add("ERRO - Validação do schema.")
            NFCE_Grava_Log("ERRO - Validação do schema.")

            glErro = 1

        Finally
            If Not (reader Is Nothing) Then
                reader.Close()
            End If
            validaXML = glErro
        End Try
        '          Console.WriteLine("\rFim de validação\n");
        'Console.ReadLine();
    End Function

    Public Function NFCE_Grava_Log(ByVal sTexto As String, Optional ByVal sTipo As String = "INFO") As Long

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

    Sub reader_ValidationEventHandler(ByVal sender As Object, ByVal e As ValidationEventArgs)

        '            // Report back error information to the console...
        '        MessageBox.Show(e.Exception.Message)
        NFCE_Grava_Log("Linha: " & e.Exception.LinePosition & " Coluna: " & e.Exception.LineNumber & " Erro: " & e.Exception.Message & " Name: " & sender.Name & " Valor:  " & sender.Value & "")

        gsMsg = e.Exception.Message
        glErro = 1

    End Sub

End Class
