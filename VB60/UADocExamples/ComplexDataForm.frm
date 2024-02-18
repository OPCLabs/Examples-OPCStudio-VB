VERSION 5.00
Begin VB.Form ComplexDataForm 
   Caption         =   "ComplexData"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton EasyUAClient_ReadValue_Main_Command 
      Caption         =   "_EasyUAClient.ReadValue.Main"
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
Attribute VB_Name = "ComplexDataForm"
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

REM #region Example _EasyUAClient.ReadValue.Main
REM Shows how to read complex data with OPC UA Complex Data plug-in.
REM
REM Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Private Sub EasyUAClient_ReadValue_Main_Command_Click()
    OutputText = ""
    
    ' Define which server and node we will work with.
    Dim endpointDescriptor As String
    'endpointDescriptor = "http://opcua.demo-this.com:51211/UA/SampleServer"
    'endpointDescriptor = "https://opcua.demo-this.com:51212/UA/SampleServer/"
    endpointDescriptor = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    Dim nodeDescriptor As String
    nodeDescriptor = "nsu=http://test.org/UA/Data/ ;i=10239" ' [ObjectsFolder]/Data.Static.Scalar.StructureValue
        
    ' Instantiate the client object
    Dim client As New EasyUAClient

    ' Read a node which returns complex data. This is done in the same way as regular reads - just the data
    ' returned is different.
    On Error Resume Next
    Dim value As Variant
    Set value = client.ReadValue(endpointDescriptor, nodeDescriptor)
    If Err.Number <> 0 Then
        OutputText = OutputText & "*** Failure: " & Err.Source & ": " & Err.Description & vbCrLf
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Display basic information about what we have read.
    OutputText = OutputText & "value: " & value & vbCrLf
    
    ' We know that this node returns complex data, so it is a UAGenericObject.
    Dim genericObject As UAGenericObject
    Set genericObject = value
        
    ' The actual data is in the GenericData property of the UAGenericObject.
    '
    ' If we want to see the whole hierarchy of the received complex data, we can format it with the "V" (verbose)
    ' specifier. In the debugger, you can view the same by displaying the private DebugView property.
    OutputText = OutputText & vbCrLf
    OutputText = OutputText & genericObject.genericData.ToString_2("V", Nothing) & vbCrLf

    ' For processing the internals of the data, refer to examples for GenericData and DataType classes.

    ' Example output (truncated):
    '
    '(ScalarValueDataType) structured
    '
    '(ScalarValueDataType) structured
    '  [BooleanValue] (Boolean) primitive; True {System.Boolean}
    '  [ByteStringValue] (ByteString) primitive; System.Byte[] {System.Byte[]}
    '  [ByteValue] (Byte) primitive; 153 {System.Byte}
    '  [DateTimeValue] (DateTime) primitive; 5/11/2013 4:32:00 PM {System.DateTime}
    '  [DoubleValue] (Double) primitive; -8.93178007363702E+27 {System.Double}
    '  [EnumerationValue] (Int32) primitive; 0 {System.Int32}
    '  [ExpandedNodeIdValue] (ExpandedNodeId) structured
    '    [NamespaceURI] (CharArray) primitive; "http://samples.org/UA/memorybuffer/Instance" {System.String}
    '    [NamespaceURISpecified] (Bit) primitive; True {System.Boolean}
    '    [NodeIdType] (NodeIdType) enumeration; 3 (String)
    '    [ServerIndexSpecified] (Bit) primitive; False {System.Boolean}
    '    [String] (StringNodeId) structured
    '      [Identifier] (CharArray) primitive; "????" {System.String}
    '      [NamespaceIndex] (UInt16) primitive; 0 {System.UInt16}
    '  [FloatValue] (Float) primitive; 78.37176 {System.Single}
    '  [GuidValue] (Guid) primitive; 8129cdaf-24d9-8140-64f2-3a6d7a957fd7 {System.Guid}
    '  [Int16Value] (Int16) primitive; 2793 {System.Int16}
    '  [Int32Value] (Int32) primitive; 1133391074 {System.Int32}
    '  [Int64Value] (Int64) primitive; -1039109760798965779 {System.Int64}
    '  [Integer] (Variant) structured
    '    [ArrayDimensionsSpecified] sequence[1]
    '      [0] (Bit) primitive; False {System.Boolean}
    '    [ArrayLengthSpecified] sequence[1]
    '      [0] (Bit) primitive; False {System.Boolean}
    '    [Int64] sequence[1]
    '      [0] (Int64) primitive; 0 {System.Int64}
    '    [VariantType] sequence[6]
    '      [0] (Bit) primitive; False {System.Boolean}
    '      [1] (Bit) primitive; False {System.Boolean}
    '      [2] (Bit) primitive; False {System.Boolean}
    '      [3] (Bit) primitive; True {System.Boolean}
    '      [4] (Bit) primitive; False {System.Boolean}
    '      [5] (Bit) primitive; False {System.Boolean}
    '  [LocalizedTextValue] (LocalizedText) structured
    '    [Locale] (CharArray) primitive; "ko" {System.String}
    '    [LocaleSpecified] (Bit) primitive; True {System.Boolean}
    '    [Reserved1] sequence[6]
    '      [0] (Bit) primitive; False {System.Boolean}
    '      [1] (Bit) primitive; False {System.Boolean}
    '      [2] (Bit) primitive; False {System.Boolean}
    '      [3] (Bit) primitive; False {System.Boolean}
    '      [4] (Bit) primitive; False {System.Boolean}
    '      [5] (Bit) primitive; False {System.Boolean}
    '    [Text] (CharArray) primitive; "? ?? ??+ ??? ??) ?: ???? ?! ?!" {System.String}
    '    [TextSpecified] (Bit) primitive; True {System.Boolean}
    '  [NodeIdValue] (NodeId) structured
End Sub
REM #endregion Example _EasyUAClient.ReadValue.Main
