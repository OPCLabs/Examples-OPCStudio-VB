VERSION 5.00
Begin VB.Form UAHostAndEndpointDialogForm 
   Caption         =   "UAHostAndEndpointDialog"
   ClientHeight    =   7350
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ShowDialog_Main_Command 
      Caption         =   "ShowDialog.Main"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.TextBox OutputText 
      Height          =   7095
      Left            =   3360
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "UAHostAndEndpointDialogForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Option Explicit On

REM #region Example ShowDialog.Main
REM This example shows how to let the user browse for a host (computer) and an endpoint of an OPC-UA server residing on it.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Private Sub ShowDialog_Main_Command_Click()
    OutputText = ""

    Dim HostAndEndpointDialog As New UAHostAndEndpointDialog
    HostAndEndpointDialog.endpointDescriptor.Host = "opcua.demo-this.com"
    
    Dim DialogResult
    DialogResult = HostAndEndpointDialog.ShowDialog
    
    OutputText = OutputText & DialogResult & vbCrLf
    If DialogResult <> 1 Then   ' OK
        Exit Sub
    End If
    
    ' Display results
    Dim HostElement As HostElement
    Set HostElement = HostAndEndpointDialog.HostElement
    If Not HostElement Is Nothing Then
        OutputText = OutputText & "HostElement: " & HostElement & vbCrLf
    End If
    Dim DiscoveryElement As UADiscoveryElement
    Set DiscoveryElement = HostAndEndpointDialog.DiscoveryElement
    If Not DiscoveryElement Is Nothing Then
        OutputText = OutputText & "DiscoveryElement: " & DiscoveryElement & vbCrLf
    End If
End Sub
REM #endregion Example ShowDialog.Main

