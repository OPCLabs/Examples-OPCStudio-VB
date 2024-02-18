VERSION 5.00
Begin VB.Form ApplicationForm 
   Caption         =   "Application"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton IEasyUAClientServerApplication_UpdateGdsRegistration_Main_Command 
      Caption         =   "_IEasyUAClientServerApplication.UpdateGdsRegistration.Main"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   4215
   End
   Begin VB.CommandButton IEasyUAClientServerApplication_RefreshTrustLists_Main_Command 
      Caption         =   "_IEasyUAClientServerApplication.RefreshTrustLists.Main"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4215
   End
   Begin VB.CommandButton IEasyUAClientServerApplication_ObtainNewCertificate_Main_Command 
      Caption         =   "_IEasyUAClientServerApplication.ObtainNewCertificate.Main"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
   Begin VB.TextBox OutputText 
      Height          =   7575
      Left            =   4440
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "ApplicationForm"
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

REM #region Example _IEasyUAClientServerApplication.ObtainNewCertificate.Main
REM Shows how to obtain a new application certificate from the certificate manager (GDS),
REM and store it for subsequent usage.
REM
REM Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Private Sub IEasyUAClientServerApplication_ObtainNewCertificate_Main_Command_Click()
    OutputText = ""
    
    ' Define which GDS we will work with.
    Dim gdsEndpointDescriptor As New UAEndpointDescriptor
    gdsEndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:58810/GlobalDiscoveryServer"
    gdsEndpointDescriptor.UserIdentity.UserNameTokenInfo.UserName = "appadmin"
    gdsEndpointDescriptor.UserIdentity.UserNameTokenInfo.Password = "demo"
    
    ' Obtain the application interface
    Dim Application As New EasyUAApplication
    
    ' Display which application we are about to work with.
    OutputText = OutputText & "Application URI string: " & Application.GetApplicationElement.applicationUriString & vbCrLf

    ' Obtain a new application certificate from the certificate manager (GDS), and store it for subsequent usage.
    Dim arguments As New UAObtainCertificateArguments
    Set arguments.Parameters.gdsEndpointDescriptor = gdsEndpointDescriptor
    
    On Error Resume Next
    Dim certificate As PkiCertificate
    Set certificate = Application.ObtainNewCertificate(arguments)
    If Err.Number <> 0 Then
        OutputText = OutputText & "*** Failure: " & Err.Source & ": " & Err.Description & vbCrLf
        Exit Sub
    End If
    On Error GoTo 0

    ' Display results
    OutputText = OutputText & "Certificate: " & certificate & vbCrLf
End Sub

REM #endregion Example _IEasyUAClientServerApplication.ObtainNewCertificate.Main

REM #region Example _IEasyUAClientServerApplication.RefreshTrustLists.Main
REM Shows how to refresh own certificate stores using current trust lists
REM for the application from the certificate manager.
REM
REM Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Private Sub IEasyUAClientServerApplication_RefreshTrustLists_Main_Command_Click()
    OutputText = ""
    
    ' Define which GDS we will work with.
    Dim gdsEndpointDescriptor As New UAEndpointDescriptor
    gdsEndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:58810/GlobalDiscoveryServer"
    gdsEndpointDescriptor.UserIdentity.UserNameTokenInfo.UserName = "appadmin"
    gdsEndpointDescriptor.UserIdentity.UserNameTokenInfo.Password = "demo"
    
    ' Obtain the application interface
    Dim Application As New EasyUAApplication
    
    ' Display which application we are about to work with.
    OutputText = OutputText & "Application URI string: " & Application.GetApplicationElement.applicationUriString & vbCrLf

    ' Refresh own certificate stores using current trust lists for the application from the certificate manager.
    On Error Resume Next
    Dim refreshedTrustLists As UATrustListMasks
    refreshedTrustLists = Application.RefreshTrustLists(gdsEndpointDescriptor, True)
    If Err.Number <> 0 Then
        OutputText = OutputText & "*** Failure: " & Err.Source & ": " & Err.Description & vbCrLf
        Exit Sub
    End If
    On Error GoTo 0

    ' Display results
    OutputText = OutputText & "Refreshed trust lists: " & refreshedTrustLists & vbCrLf
End Sub

REM #endregion Example _IEasyUAClientServerApplication.RefreshTrustLists.Main

REM #region Example _IEasyUAClientServerApplication.UpdateGdsRegistration.Main
REM Shows how to update an application registration in the GDS, keeping its application ID if possible.
REM
REM Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Private Sub IEasyUAClientServerApplication_UpdateGdsRegistration_Main_Command_Click()
    OutputText = ""
    ' Define which GDS we will work with.
    Dim gdsEndpointDescriptor As New UAEndpointDescriptor
    gdsEndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:58810/GlobalDiscoveryServer"
    gdsEndpointDescriptor.UserIdentity.UserNameTokenInfo.UserName = "appadmin"
    gdsEndpointDescriptor.UserIdentity.UserNameTokenInfo.Password = "demo"
    
    ' Obtain the application interface
    Dim Application As New EasyUAApplication
    
    ' Display which application we are about to work with.
    OutputText = OutputText & "Application URI string: " & Application.GetApplicationElement.applicationUriString & vbCrLf

    ' Update an application registration in the GDS, keeping its application ID if possible.
    On Error Resume Next
    Dim applicationId As UANodeId
    Set applicationId = Application.UpdateGdsRegistration(gdsEndpointDescriptor)
    If Err.Number <> 0 Then
        OutputText = OutputText & "*** Failure: " & Err.Source & ": " & Err.Description & vbCrLf
        Exit Sub
    End If
    On Error GoTo 0

    ' Display results
    OutputText = OutputText & "Application ID: " & applicationId & vbCrLf
End Sub

REM #endregion Example _IEasyUAClientServerApplication.UpdateGdsRegistration.Main

