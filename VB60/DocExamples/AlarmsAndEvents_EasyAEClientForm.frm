VERSION 5.00
Begin VB.Form AlarmsAndEvents_EasyAEClientForm 
   Caption         =   "AlarmsAndEvents._EasyAEClient"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton SubscribeEvents_Main_Command 
      Caption         =   "SubscribeEvents.Main"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   3015
   End
   Begin VB.CommandButton PullNotification_Main_Command 
      Caption         =   "PullNotification.Main"
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
Attribute VB_Name = "AlarmsAndEvents_EasyAEClientForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Option Explicit

' The client object, with events
Public WithEvents Client1 As EasyAEClient
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

REM #region Example PullNotification.Main
REM This example shows how to subscribe to events and obtain the notification events by pulling them.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Private Sub PullNotification_Main_Command_Click()
    OutputText = ""
    
    Dim eventArgs As EasyAENotificationEventArgs
    
    Dim serverDescriptor As New serverDescriptor
    serverDescriptor.ServerClass = "OPCLabs.KitEventServer.2"
        
    ' Instantiate the client object
    Dim client As New EasyAEClient
    
    ' In order to use event pull, you must set a non-zero queue capacity upfront.
    client.PullNotificationQueueCapacity = 1000
    
    OutputText = OutputText & "Subscribing events..." & vbCrLf
    Dim subscriptionParameters As New AESubscriptionParameters
    subscriptionParameters.notificationRate = 1000
    Dim handle
    Dim state
    handle = client.SubscribeEvents(serverDescriptor, subscriptionParameters, True, state)

    OutputText = OutputText & "Processing event notifications for 1 minute..." & vbCrLf
    Dim endTick As Long
    endTick = GetTickCount + 60000
    While GetTickCount < endTick
        Set eventArgs = client.PullNotification(2 * 1000)
        If Not eventArgs Is Nothing Then
            ' Handle the notification event
            OutputText = OutputText & eventArgs & vbCrLf
        End If
    Wend
    
    OutputText = OutputText & "Unsubscribing events..." & vbCrLf
    client.UnsubscribeEvents handle

    OutputText = OutputText & "Finished." & vbCrLf

End Sub
REM #endregion Example PullNotification.Main

REM #region Example SubscribeEvents.Main
REM This example shows how to subscribe to events and display the event message with each notification. It also shows how to
REM unsubscribe afterwards.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Private Sub SubscribeEvents_Main_Command_Click()
    OutputText = ""
    
    Dim serverDescriptor As New serverDescriptor
    serverDescriptor.ServerClass = "OPCLabs.KitEventServer.2"
    
    ' Instantiate the client object and hook events
    Set Client1 = New EasyAEClient
    
    OutputText = OutputText & "Subscribing..." & vbCrLf
    Dim subscriptionParameters As New AESubscriptionParameters
    subscriptionParameters.notificationRate = 1000
    Dim handle
    Dim state
    handle = Client1.SubscribeEvents(serverDescriptor, subscriptionParameters, True, state)

    OutputText = OutputText & "Processing event notifications for 1 minute..." & vbCrLf
    Pause 60000

    OutputText = OutputText & "Unsubscribing events..." & vbCrLf
    Client1.UnsubscribeEvents handle

    OutputText = OutputText & "Waiting for 5 seconds..." & vbCrLf
    Pause 5000

    OutputText = OutputText & "Finished." & vbCrLf

    Set Client1 = Nothing
End Sub

Private Sub Client1_OnNotification(ByVal sender As Variant, ByVal eventArgs As EasyAENotificationEventArgs)
    If Not eventArgs.Succeeded Then
        OutputText = OutputText & eventArgs.ErrorMessageBrief & vbCrLf
        Exit Sub
    End If
    If Not eventArgs.EventData Is Nothing Then
        OutputText = OutputText & eventArgs.EventData.Message & vbCrLf
    End If
End Sub
REM #endregion Example SubscribeEvents.Main


