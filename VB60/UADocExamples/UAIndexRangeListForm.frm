VERSION 5.00
Begin VB.Form UAIndexRangeListForm 
   Caption         =   "UAIndexRangeList"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Usage_ReadValue_Command 
      Caption         =   "Usage.ReadValue"
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
Attribute VB_Name = "UAIndexRangeListForm"
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

REM #region Example Usage.ReadValue
REM This example shows how to read a range of values from an array.
REM
REM Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Private Sub Usage_ReadValue_Command_Click()
    OutputText = ""
    
    Dim endpointDescriptor As String
    'endpointDescriptor = "http://opcua.demo-this.com:51211/UA/SampleServer"
    'endpointDescriptor = "https://opcua.demo-this.com:51212/UA/SampleServer/"
    endpointDescriptor = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"

    ' Instantiate the client object
    Dim Client As New EasyUAClient

    ' Prepare the arguments, indicating that just the elements 2 to 4 should be returned.
    Dim indexRangeList As New UAIndexRangeList
    Dim indexRange As New UAIndexRange
    indexRange.Minimum = 2
    indexRange.Maximum = 4
    indexRangeList.Add indexRange
    
    Dim readArguments1 As New UAReadArguments
    readArguments1.endpointDescriptor.UrlString = endpointDescriptor
    readArguments1.NodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;ns=2;i=10305"
    Set readArguments1.indexRangeList = indexRangeList

    Dim arguments(0) As Variant
    Set arguments(0) = readArguments1

    ' Obtain value.
    Dim results() As Variant
    results = Client.ReadMultipleValues(arguments)

    Dim valueResult As valueResult
    Set valueResult = results(0)
    If Not valueResult.Succeeded Then
        OutputText = OutputText & "*** Failure: " & valueResult.Exception.GetBaseException.Message & vbCrLf
        Exit Sub
    End If
    
    ' Display results
    Dim arrayValue() As Long
    arrayValue = valueResult.value
    
    Dim i: For i = LBound(arrayValue) To UBound(arrayValue)
        OutputText = OutputText & "arrayValue(" & i & "):" & arrayValue(i) & vbCrLf
    Next

    ' Example output:
    'arrayValue(0): 180410224
    'arrayValue(1): 1919239969
    'arrayValue(2): 1700185172
End Sub
REM #endregion Example Usage.ReadValue

