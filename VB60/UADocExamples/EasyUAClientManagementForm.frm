VERSION 5.00
Begin VB.Form EasyUAClientManagementForm 
   Caption         =   "EasyUAClientManagement"
   ClientHeight    =   7350
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton LogEntry_Main_Command 
      Caption         =   "LogEntry.Main"
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
Attribute VB_Name = "EasyUAClientManagementForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Option Explicit

' The management object allows access to static behavior - here, the shared LogEntry event.
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

REM #region Example LogEntry.Main
REM This example demonstrates the loggable entries originating in the OPC-UA client engine and the EasyUAClient component.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

' The management object allows access to static behavior - here, the shared LogEntry event.
'Public WithEvents ClientManagement1 As EasyUAClientManagement

Private Sub LogEntry_Main_Command_Click()
    OutputText = ""
    
    Set ClientManagement1 = New EasyUAClientManagement
    
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

    OutputText = OutputText & "Processing log entry events for 1 minute..." & vbCrLf
    Pause 60000
    
    Set ClientManagement1 = Nothing
    OutputText = OutputText & "Finished..." & vbCrLf
End Sub

' Event handler for the LogEntry event. It simply prints out the event.
Private Sub ClientManagement1_LogEntry(ByVal sender As Variant, ByVal eventArgs As OpcLabs_BaseLib.LogEntryEventArgs)
    OutputText = OutputText & eventArgs & vbCrLf
End Sub
REM #endregion Example LogEntry.Main
