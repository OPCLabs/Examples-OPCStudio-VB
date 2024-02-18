VERSION 5.00
Begin VB.Form UAEndpointDialogForm 
   Caption         =   "UAEndpointDialog"
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
Attribute VB_Name = "UAEndpointDialogForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Option Explicit On

REM #region Example ShowDialog.Main
REM This example shows how to let the user browse for an OPC-UA server endpoint.
REM
REM Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Private Sub ShowDialog_Main_Command_Click()
    OutputText = ""

    Dim EndpointDialog As New UAEndpointDialog
    EndpointDialog.DiscoveryHost = "opcua.demo-this.com"
    
    Dim DialogResult
    DialogResult = EndpointDialog.ShowDialog
    
    OutputText = OutputText & DialogResult & vbCrLf
    If DialogResult <> 1 Then   ' OK
        Exit Sub
    End If
    
    ' Display results
    OutputText = OutputText & EndpointDialog.DiscoveryElement & vbCrLf
End Sub
REM #endregion Example ShowDialog.Main

