VERSION 5.00
Begin VB.Form GdsForm 
   Caption         =   "Gds"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton EasyUAGlobalDiscoveryClient_QueryServers_Main_Command 
      Caption         =   "_EasyUAGlobalDiscoveryClient.QueryServers.Main"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   4815
   End
   Begin VB.CommandButton EasyUAGlobalDiscoveryClient_QueryApplications_Main_Command 
      Caption         =   "_EasyUAGlobalDiscoveryClient.QueryApplications.Main"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4815
   End
   Begin VB.CommandButton EasyUACertificateManagementClient_GetCertificateStatus_Main_Command 
      Caption         =   "_EasyUACertificateManagementClient.GetCertificateStatus.Main"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4815
   End
   Begin VB.TextBox OutputText 
      Height          =   7575
      Left            =   5040
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "GdsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Option Explicit

' Pause
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Public Sub Pause(Optional milliseconds As Long)
    On Error Resume Next
    Dim endTickCount As Long
    endTickCount = GetTickCount + milliseconds
    While GetTickCount < endTickCount: Sleep 1: DoEvents: Wend
End Sub

REM #region Example _EasyUACertificateManagementClient.GetCertificateStatus.Main
REM Shows how to check if an application needs to update its certificate.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Private Sub EasyUACertificateManagementClient_GetCertificateStatus_Main_Command_Click()
    OutputText = ""
        
    ' Define which GDS we will work with.
    Dim gdsEndpointDescriptor As New UAEndpointDescriptor
    gdsEndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:58810/GlobalDiscoveryServer"
    gdsEndpointDescriptor.UserIdentity.UserNameTokenInfo.UserName = "appadmin"
    gdsEndpointDescriptor.UserIdentity.UserNameTokenInfo.Password = "demo"
    
    ' Register our client application with the GDS, so that we obtain an application ID that we need later.
    ' Obtain the application interface
    Dim Application As New EasyUAApplication
    
    ' Create an application registration in the GDS, assigning it a new application ID.
    On Error Resume Next
    Dim applicationId As UANodeId
    Set applicationId = Application.RegisterToGds(gdsEndpointDescriptor)
    If Err.Number <> 0 Then
        OutputText = OutputText & "*** Failure: " & Err.Source & ": " & Err.Description & vbCrLf
        Exit Sub
    End If
    On Error GoTo 0

    OutputText = OutputText & "Application ID: " & applicationId & vbCrLf

    ' Instantiate the certificate management client object
    Dim certificateManagementClient As New EasyUACertificateManagementClient
    
    ' Check if the application needs to update its certificate.
    Dim nullNodeId As New UANodeId
    Dim updateRequired As Boolean: updateRequired = False
    On Error Resume Next
    updateRequired = certificateManagementClient.GetCertificateStatus(gdsEndpointDescriptor, applicationId, nullNodeId, nullNodeId)
    If Err.Number <> 0 Then
        OutputText = OutputText & "*** Failure: " & Err.Source & ": " & Err.Description & vbCrLf
        Exit Sub
    End If
    On Error GoTo 0

    ' Display results
    OutputText = OutputText & "Update required: " & updateRequired & vbCrLf

    ' Example output:
    'Application ID: nsu=http://opcfoundation.org/UA/GDS/applications/ ;ns=2;g=aec94459-f513-4979-8619-8383555fca61
    'Update required: FALSE
End Sub

REM #endregion Example _EasyUACertificateManagementClient.GetCertificateStatus.Main

REM #region Example _EasyUAGlobalDiscoveryClient.QueryApplications.Main
REM Shows how to find client or server applications that meet the specified filters, using the global discovery client.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Private Sub EasyUAGlobalDiscoveryClient_QueryApplications_Main_Command_Click()
    OutputText = ""
    
    ' Define which GDS we will work with.
    Dim gdsEndpointDescriptor As New UAEndpointDescriptor
    gdsEndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:58810/GlobalDiscoveryServer"
    
    ' Instantiate the global discovery client object
    Dim globalDiscoveryClient As New EasyUAGlobalDiscoveryClient
    
    ' Find all (client or server) applications registered in the GDS.
    Dim startingRecordId As Long: startingRecordId = 0
    Dim maximumRecordsToReturn As Long: maximumRecordsToReturn = 0
    Dim applicationName As String: applicationName = ""
    Dim applicationUriString As String: applicationUriString = ""
    Dim productUriString As String: productUriString = ""
    Dim serverCapabilities: serverCapabilities = Array()
    Dim lastCounterResetTime As Date
    Dim nextRecordId As Long
    Dim applicationDescriptionArray As Variant
    On Error Resume Next
    Dim applicationId As UANodeId
    Call globalDiscoveryClient.QueryApplications( _
      gdsEndpointDescriptor, _
      startingRecordId, _
      maximumRecordsToReturn, _
      applicationName, _
      applicationUriString, _
      UAApplicationTypes_All, _
      productUriString, _
      serverCapabilities, _
      lastCounterResetTime, _
      nextRecordId, _
      applicationDescriptionArray)
    If Err.Number <> 0 Then
        OutputText = OutputText & "*** Failure: " & Err.Source & ": " & Err.Description & vbCrLf
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Display results
    Dim i: For i = LBound(applicationDescriptionArray) To UBound(applicationDescriptionArray)
        Dim applicationDescription As UAApplicationDescription
        Set applicationDescription = applicationDescriptionArray(i)
        OutputText = OutputText & vbCrLf
        OutputText = OutputText & "Application name: " & applicationDescription.applicationName & vbCrLf
        OutputText = OutputText & "Application type: " & applicationDescription.ApplicationType & vbCrLf
        OutputText = OutputText & "Application URI string: " & applicationDescription.applicationUriString & vbCrLf
        OutputText = OutputText & "Discovery URI strings: " & applicationDescription.DiscoveryUriStrings.ToString & vbCrLf
    Next
End Sub

REM #endregion Example _EasyUAGlobalDiscoveryClient.QueryApplications.Main

REM #region Example _EasyUAGlobalDiscoveryClient.QueryServers.Main
REM Shows how to find server applications that meet the specified filters, using the global discovery client.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Private Sub EasyUAGlobalDiscoveryClient_QueryServers_Main_Command_Click()
    OutputText = ""
    
    ' Define which GDS we will work with.
    Dim gdsEndpointDescriptor As New UAEndpointDescriptor
    gdsEndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:58810/GlobalDiscoveryServer"
    
    ' Instantiate the global discovery client object
    Dim globalDiscoveryClient As New EasyUAGlobalDiscoveryClient
    
    ' Find all servers registered in the GDS.
    Dim startingRecordId As Integer: startingRecordId = 0
    Dim maximumRecordsToReturn As Integer: maximumRecordsToReturn = 0
    Dim applicationName As String: applicationName = ""
    Dim applicationUriString As String: applicationUriString = ""
    Dim productUriString As String: productUriString = ""
    Dim serverCapabilities: serverCapabilities = Array()
    Dim lastCounterResetTime As Date
    Dim serverOnNetworkArray As Variant
    On Error Resume Next
    Dim applicationId As UANodeId
    Call globalDiscoveryClient.QueryServers( _
      gdsEndpointDescriptor, _
      startingRecordId, _
      maximumRecordsToReturn, _
      applicationName, _
      applicationUriString, _
      productUriString, _
      serverCapabilities, _
      lastCounterResetTime, _
      serverOnNetworkArray)
    If Err.Number <> 0 Then
        OutputText = OutputText & "*** Failure: " & Err.Source & ": " & Err.Description & vbCrLf
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Display results
    Dim i: For i = LBound(serverOnNetworkArray) To UBound(serverOnNetworkArray)
        Dim serverOnNetwork As UAServerOnNetwork
        Set serverOnNetwork = serverOnNetworkArray(i)
        OutputText = OutputText & vbCrLf
        OutputText = OutputText & "ServerName name: " & serverOnNetwork.ServerName & vbCrLf
        OutputText = OutputText & "Discovery URL string: " & serverOnNetwork.DiscoveryUrlString & vbCrLf
        OutputText = OutputText & "Server capabilities: " & serverOnNetwork.serverCapabilities.ToString & vbCrLf
    Next
End Sub

REM #endregion Example _EasyUAGlobalDiscoveryClient.QueryServers.Main


