﻿'------------------------------------------------------------------------------
' <auto-generated>
'     O código foi gerado por uma ferramenta.
'     Versão de Tempo de Execução:4.0.30319.42000
'
'     As alterações ao arquivo poderão causar comportamento incorreto e serão perdidas se
'     o código for gerado novamente.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict Off
Option Explicit On

Imports System
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Xml.Serialization

'
'Este código-fonte foi gerado automaticamente por Microsoft.VSDesigner, Versão 4.0.30319.42000.
'
Namespace nfeconsulta2
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.7.3056.0"),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code"),  _
     System.Web.Services.WebServiceBindingAttribute(Name:="NFeConsultaProtocolo4Soap12", [Namespace]:="http://www.portalfiscal.inf.br/nfe/wsdl/NFeConsultaProtocolo4")>  _
    Partial Public Class NFeConsultaProtocolo4
        Inherits System.Web.Services.Protocols.SoapHttpClientProtocol
        
        Private nfeConsultaNFOperationCompleted As System.Threading.SendOrPostCallback
        
        Private useDefaultCredentialsSetExplicitly As Boolean
        
        '''<remarks/>
        Public Sub New()
            MyBase.New
            Me.SoapVersion = System.Web.Services.Protocols.SoapProtocolVersion.Soap12
            Me.Url = Global.sgenfe4.My.MySettings.Default.sgenfe4_nfeconsulta2_NFeConsultaProtocolo4
            If (Me.IsLocalFileSystemWebService(Me.Url) = true) Then
                Me.UseDefaultCredentials = true
                Me.useDefaultCredentialsSetExplicitly = false
            Else
                Me.useDefaultCredentialsSetExplicitly = true
            End If
        End Sub
        
        Public Shadows Property Url() As String
            Get
                Return MyBase.Url
            End Get
            Set
                If (((Me.IsLocalFileSystemWebService(MyBase.Url) = true)  _
                            AndAlso (Me.useDefaultCredentialsSetExplicitly = false))  _
                            AndAlso (Me.IsLocalFileSystemWebService(value) = false)) Then
                    MyBase.UseDefaultCredentials = false
                End If
                MyBase.Url = value
            End Set
        End Property
        
        Public Shadows Property UseDefaultCredentials() As Boolean
            Get
                Return MyBase.UseDefaultCredentials
            End Get
            Set
                MyBase.UseDefaultCredentials = value
                Me.useDefaultCredentialsSetExplicitly = true
            End Set
        End Property
        
        '''<remarks/>
        Public Event nfeConsultaNFCompleted As nfeConsultaNFCompletedEventHandler
        
        '''<remarks/>
        <System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.portalfiscal.inf.br/nfe/wsdl/NFeConsultaProtocolo4/nfeConsultaNF", Use:=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle:=System.Web.Services.Protocols.SoapParameterStyle.Bare)>  _
        Public Function nfeConsultaNF(<System.Xml.Serialization.XmlElementAttribute([Namespace]:="http://www.portalfiscal.inf.br/nfe/wsdl/NFeConsultaProtocolo4")> ByVal nfeDadosMsg As System.Xml.XmlNode) As <System.Xml.Serialization.XmlElementAttribute("nfeResultMsg", [Namespace]:="http://www.portalfiscal.inf.br/nfe/wsdl/NFeConsultaProtocolo4")> System.Xml.XmlNode
            Dim results() As Object = Me.Invoke("nfeConsultaNF", New Object() {nfeDadosMsg})
            Return CType(results(0),System.Xml.XmlNode)
        End Function
        
        '''<remarks/>
        Public Overloads Sub nfeConsultaNFAsync(ByVal nfeDadosMsg As System.Xml.XmlNode)
            Me.nfeConsultaNFAsync(nfeDadosMsg, Nothing)
        End Sub
        
        '''<remarks/>
        Public Overloads Sub nfeConsultaNFAsync(ByVal nfeDadosMsg As System.Xml.XmlNode, ByVal userState As Object)
            If (Me.nfeConsultaNFOperationCompleted Is Nothing) Then
                Me.nfeConsultaNFOperationCompleted = AddressOf Me.OnnfeConsultaNFOperationCompleted
            End If
            Me.InvokeAsync("nfeConsultaNF", New Object() {nfeDadosMsg}, Me.nfeConsultaNFOperationCompleted, userState)
        End Sub
        
        Private Sub OnnfeConsultaNFOperationCompleted(ByVal arg As Object)
            If (Not (Me.nfeConsultaNFCompletedEvent) Is Nothing) Then
                Dim invokeArgs As System.Web.Services.Protocols.InvokeCompletedEventArgs = CType(arg,System.Web.Services.Protocols.InvokeCompletedEventArgs)
                RaiseEvent nfeConsultaNFCompleted(Me, New nfeConsultaNFCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState))
            End If
        End Sub
        
        '''<remarks/>
        Public Shadows Sub CancelAsync(ByVal userState As Object)
            MyBase.CancelAsync(userState)
        End Sub
        
        Private Function IsLocalFileSystemWebService(ByVal url As String) As Boolean
            If ((url Is Nothing)  _
                        OrElse (url Is String.Empty)) Then
                Return false
            End If
            Dim wsUri As System.Uri = New System.Uri(url)
            If ((wsUri.Port >= 1024)  _
                        AndAlso (String.Compare(wsUri.Host, "localHost", System.StringComparison.OrdinalIgnoreCase) = 0)) Then
                Return true
            End If
            Return false
        End Function
    End Class
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.7.3056.0")>  _
    Public Delegate Sub nfeConsultaNFCompletedEventHandler(ByVal sender As Object, ByVal e As nfeConsultaNFCompletedEventArgs)
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.7.3056.0"),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code")>  _
    Partial Public Class nfeConsultaNFCompletedEventArgs
        Inherits System.ComponentModel.AsyncCompletedEventArgs
        
        Private results() As Object
        
        Friend Sub New(ByVal results() As Object, ByVal exception As System.Exception, ByVal cancelled As Boolean, ByVal userState As Object)
            MyBase.New(exception, cancelled, userState)
            Me.results = results
        End Sub
        
        '''<remarks/>
        Public ReadOnly Property Result() As System.Xml.XmlNode
            Get
                Me.RaiseExceptionIfNecessary
                Return CType(Me.results(0),System.Xml.XmlNode)
            End Get
        End Property
    End Class
End Namespace
