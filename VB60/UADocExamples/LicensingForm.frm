VERSION 5.00
Begin VB.Form LicensingForm 
   Caption         =   "Licensing"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton LicenseInfo_SerialNumber_Command 
      Caption         =   "LicenseInfo.SerialNumber"
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
Attribute VB_Name = "LicensingForm"
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

REM #region Example LicenseInfo.SerialNumber
REM Shows how to obtain the serial number of the active license, and determine whether it is a stock demo or trial license.
REM
REM Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Private Sub LicenseInfo_SerialNumber_Command_Click()
    OutputText = ""
        
    ' Instantiate the client object
    Dim client As New EasyUAClient

    ' Obtain the serial number from the license info.
    Dim serialNumber As Long
    serialNumber = client.LicenseInfo("Multipurpose.SerialNumber")
    
    ' Display the serial number.
    OutputText = OutputText & "Serial number: " & serialNumber & vbCrLf
    
    ' Determine whether we are running as demo or trial.
    If (1111110000 <= serialNumber) And (serialNumber <= 1111119999) Then
        OutputText = OutputText & "This is a stock demo or trial license." & vbCrLf
    Else
        OutputText = OutputText & "This is not a stock demo or trial license." & vbCrLf
    End If
End Sub
REM #endregion Example LicenseInfo.SerialNumber
