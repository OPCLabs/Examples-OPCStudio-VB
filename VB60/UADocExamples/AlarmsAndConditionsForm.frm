VERSION 5.00
Begin VB.Form AlarmsAndConditionsForm 
   Caption         =   "AlarmsAndConditions"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton SubscribeEvent_Main_Command 
      Caption         =   "SubscribeEvent.Main"
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
Attribute VB_Name = "AlarmsAndConditionsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Option Explicit

' The client object, with events
Public WithEvents Client1 As EasyUAClient
Attribute Client1.VB_VarHelpID = -1

' Pause
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Public Sub Pause(Optional milliseconds As Long)
    On Error Resume Next
    Dim endTickCount As Long
    endTickCount = GetTickCount + milliseconds
    While GetTickCount < endTickCount: Sleep 1: DoEvents: Wend
End Sub

REM #region Example SubscribeEvent.Main
REM This example shows how to subscribe to event notifications and display each incoming event.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Private Sub SubscribeEvent_Main_Command_Click()
    OutputText = ""
    
    Const UAObjectIds_Server = "nsu=http://opcfoundation.org/UA/;i=2253"
    
    Dim endpointDescriptor As String
    endpointDescriptor = "opc.tcp://opcua.demo-this.com:62544/Quickstarts/AlarmConditionServer"
    
    ' Instantiate the client object and hook events
    Set Client1 = New EasyUAClient

    OutputText = OutputText & "Subscribing..." & vbCrLf
    Call Client1.SubscribeEvent(endpointDescriptor, UAObjectIds_Server, 1000)

    OutputText = OutputText & "Processing event notifications for 30 seconds..." & vbCrLf
    Pause 30000

    OutputText = OutputText & "Unsubscribing..." & vbCrLf
    Client1.UnsubscribeAllMonitoredItems

    OutputText = OutputText & "Waiting for 5 seconds..." & vbCrLf
    Pause 5000

    OutputText = OutputText & "Finished." & vbCrLf

    Set Client1 = Nothing
End Sub

Private Sub Client1_EventNotification(ByVal sender As Variant, ByVal eventArgs As EasyUAEventNotificationEventArgs)
    ' Display the event
    If eventArgs.Succeeded Then
        OutputText = OutputText & eventArgs & vbCrLf
    Else
        OutputText = OutputText & eventArgs.ErrorMessageBrief & vbCrLf
    End If
End Sub

REM #endregion Example SubscribeEvent.Main



