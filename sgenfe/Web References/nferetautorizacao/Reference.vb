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
Namespace nferetautorizacao
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.7.3056.0"),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code"),  _
     System.Web.Services.WebServiceBindingAttribute(Name:="NFeRetAutorizacao4Soap12", [Namespace]:="http://www.portalfiscal.inf.br/nfe/wsdl/NFeRetAutorizacao4")>  _
    Partial Public Class NFeRetAutorizacao4
        Inherits System.Web.Services.Protocols.SoapHttpClientProtocol
        
        Private nfeRetAutorizacaoLoteOperationCompleted As System.Threading.SendOrPostCallback
        
        Private useDefaultCredentialsSetExplicitly As Boolean
        
        '''<remarks/>
        Public Sub New()
            MyBase.New
            Me.SoapVersion = System.Web.Services.Protocols.SoapProtocolVersion.Soap12
            Me.Url = Global.sgenfe4.My.MySettings.Default.sgenfe4_nferetautorizacao_NFeRetAutorizacao4
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
        Public Event nfeRetAutorizacaoLoteCompleted As nfeRetAutorizacaoLoteCompletedEventHandler
        
        '''<remarks/>
        <System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.portalfiscal.inf.br/nfe/wsdl/NFeRetAutorizacao4/nfeRetAutorizacaoLote", Use:=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle:=System.Web.Services.Protocols.SoapParameterStyle.Bare)>  _
        Public Function nfeRetAutorizacaoLote(<System.Xml.Serialization.XmlElementAttribute([Namespace]:="http://www.portalfiscal.inf.br/nfe/wsdl/NFeRetAutorizacao4")> ByVal nfeDadosMsg As System.Xml.XmlNode) As <System.Xml.Serialization.XmlElementAttribute("nfeResultMsg", [Namespace]:="http://www.portalfiscal.inf.br/nfe/wsdl/NFeRetAutorizacao4")> System.Xml.XmlNode
            Dim results() As Object = Me.Invoke("nfeRetAutorizacaoLote", New Object() {nfeDadosMsg})
            Return CType(results(0),System.Xml.XmlNode)
        End Function
        
        '''<remarks/>
        Public Overloads Sub nfeRetAutorizacaoLoteAsync(ByVal nfeDadosMsg As System.Xml.XmlNode)
            Me.nfeRetAutorizacaoLoteAsync(nfeDadosMsg, Nothing)
        End Sub
        
        '''<remarks/>
        Public Overloads Sub nfeRetAutorizacaoLoteAsync(ByVal nfeDadosMsg As System.Xml.XmlNode, ByVal userState As Object)
            If (Me.nfeRetAutorizacaoLoteOperationCompleted Is Nothing) Then
                Me.nfeRetAutorizacaoLoteOperationCompleted = AddressOf Me.OnnfeRetAutorizacaoLoteOperationCompleted
            End If
            Me.InvokeAsync("nfeRetAutorizacaoLote", New Object() {nfeDadosMsg}, Me.nfeRetAutorizacaoLoteOperationCompleted, userState)
        End Sub
        
        Private Sub OnnfeRetAutorizacaoLoteOperationCompleted(ByVal arg As Object)
            If (Not (Me.nfeRetAutorizacaoLoteCompletedEvent) Is Nothing) Then
                Dim invokeArgs As System.Web.Services.Protocols.InvokeCompletedEventArgs = CType(arg,System.Web.Services.Protocols.InvokeCompletedEventArgs)
                RaiseEvent nfeRetAutorizacaoLoteCompleted(Me, New nfeRetAutorizacaoLoteCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState))
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
    Public Delegate Sub nfeRetAutorizacaoLoteCompletedEventHandler(ByVal sender As Object, ByVal e As nfeRetAutorizacaoLoteCompletedEventArgs)
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.7.3056.0"),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code")>  _
    Partial Public Class nfeRetAutorizacaoLoteCompletedEventArgs
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
