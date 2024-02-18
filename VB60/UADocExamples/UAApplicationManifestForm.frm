VERSION 5.00
Begin VB.Form UAApplicationManifestForm 
   Caption         =   "UAApplicationManifest"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton InstanceOwnStorePath_PlatformSpecific_Command 
      Caption         =   "InstanceOwnStorePath.PlatformSpecific"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   3015
   End
   Begin VB.CommandButton ApplicationName_Main_Command 
      Caption         =   "ApplicationName.Main"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.TextBox OutputText 
      Height          =   7575
      Left            =   3360
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "UAApplicationManifestForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Option Explicit

' The configuration object allows access to static behavior - here, the shared LogEntry event.
Public WithEvents ClientManagement1 As EasyUAClientManagement
Attribute ClientManagement1.VB_VarHelpID = -1

' Pause
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Public Sub Pause(Optional milliseconds As Long)
    On Error Resume Next
    Dim endTickCount As Long
    endTickCount = GetTickCount + milliseconds
    While GetTickCount < endTickCount: Sleep 1: DoEvents: Wend
End Sub

REM #region Example ApplicationName.Main
REM This example demonstrates how to set the application name for the client certificate.
REM
REM Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Private Sub ApplicationName_Main_Command_Click()
    OutputText = ""
    
    ' Obtain the application interface
    Dim Application As New EasyUAApplication
        
    ' Set the application name, which determines the subject of the client certificate.
    ' Note that this only works once in each host process.
    Application.ApplicationParameters.ApplicationManifest.applicationName = "QuickOPC - VB6 example application"

    ' Do something - invoke an OPC read, to trigger some loggable entries.
    Dim client As New EasyUAClient
    On Error Resume Next
    Dim value As Variant
    value = client.ReadValue("opc.tcp://opcua.demo-this.com:51210/UA/SampleServer", "nsu=http://test.org/UA/Data/ ;i=10853")
    If Err.Number <> 0 Then
        OutputText = OutputText & "*** Failure: " & Err.Source & ": " & Err.Description & vbCrLf
        Exit Sub
    End If
    On Error GoTo 0

    ' The certificate will be located or created in a directory similar to:
    ' C:\ProgramData\OPC Foundation\CertificateStores\MachineDefault\certs
    ' and its subject will be as given by the application name.

    OutputText = OutputText & "Processing log entry events for 10 seconds..." & vbCrLf
    Pause 10000
    
    Set Application = Nothing
    OutputText = OutputText & "Finished..." & vbCrLf
End Sub
' Event handler for the LogEntry event. It simply prints out the event.
Private Sub ClientManagement1_LogEntry(ByVal sender As Variant, ByVal eventArgs As OpcLabs_BaseLib.LogEntryEventArgs)
    If eventArgs.eventId = 161 Then
        OutputText = OutputText & eventArgs & vbCrLf
    End If
End Sub
REM #endregion Example ApplicationName.Main

REM #region Example InstanceOwnStorePath.PlatformSpecific
REM  This example demonstrates how to place the client certificate
REM in the platform-specific (Windows, Linux, ...) certificate store.
REM
REM Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Private Sub InstanceOwnStorePath_PlatformSpecific_Command_Click()
    OutputText = ""
    
    ' Obtain the application interface
    Dim Application As New EasyUAApplication
        
    ' Set the application certificate store path, which determines the location of the client certificate.
    ' Note that this only works once in each host process.
    Application.ApplicationParameters.ApplicationManifest.InstanceOwnStorePath = "CurrentUser\My"

    ' Do something - invoke an OPC read, to trigger creation of the certificate.
    Dim client As New EasyUAClient
    On Error Resume Next
    Dim value As Variant
    value = client.ReadValue("opc.tcp://opcua.demo-this.com:51210/UA/SampleServer", "nsu=http://test.org/UA/Data/ ;i=10853")
    If Err.Number <> 0 Then
        OutputText = OutputText & "*** Failure: " & Err.Source & ": " & Err.Description & vbCrLf
        Exit Sub
    End If
    On Error GoTo 0

    ' The certificate will be located or created in the specified platform-specific certificate store.
    ' On Windows, when viewed by the certmgr.msc tool, it will be under
    ' Certificates - Current User -> Personal -> Certificates.
    
    OutputText = OutputText & "Finished..." & vbCrLf
End Sub

REM #endregion Example InstanceOwnStorePath.PlatformSpecific
