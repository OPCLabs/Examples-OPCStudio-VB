VERSION 5.00
Begin VB.Form EasyUAClientForm 
   Caption         =   "EasyUAClient"
   ClientHeight    =   8670
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
   Begin VB.TextBox OutputText 
      Height          =   8415
      Left            =   3840
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "EasyUAClientForm"
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

' The client object, with events
Public WithEvents Client2 As EasyUAClient
Attribute Client2.VB_VarHelpID = -1

' The client object, with events
Public WithEvents Client3 As EasyUAClient
Attribute Client3.VB_VarHelpID = -1

' The client object, with events
Public WithEvents Client4 As EasyUAClient
Attribute Client4.VB_VarHelpID = -1

' The client object, with events
Public WithEvents Client5 As EasyUAClient
Attribute Client5.VB_VarHelpID = -1

' The client object, with events
Public WithEvents Client6 As EasyUAClient
Attribute Client6.VB_VarHelpID = -1

' The client object, with events
Public WithEvents Client7 As EasyUAClient
Attribute Client7.VB_VarHelpID = -1

' The client object, with events
Public WithEvents Client8 As EasyUAClient
Attribute Client8.VB_VarHelpID = -1

' Pause
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Public Sub Pause(Optional milliseconds As Long)
    On Error Resume Next
    Dim endTickCount As Long
    endTickCount = GetTickCount + milliseconds
    While GetTickCount < endTickCount: Sleep 1: DoEvents: Wend
End Sub

Private Sub Form_Load()
    Dim actions
    Dim index As Integer: index = 0
    Dim top As Integer: top = 120
    Dim action
    
    actions = Array( _
    "BrowseDataNodes.Main", "BrowseDataVariables.Main", _
    "BrowseMethods.Main", "BrowseObjects.Main", _
    "BrowseProperties.Main", "Browse.Main", _
    "CallMethod.Main", "CallMultipleMethods.Main", _
    "ChangeMonitoredItemSubscription.Main", _
    "DiscoverGlobalServers.Main", "DiscoverLocalServers.Main", _
    "DiscoverNetworkServers.Main", "FindLocalApplications.Main", _
    "PullDataChangeNotification.Main", _
    "Read.Main", _
    "ReadMultiple.BrowsePath", "ReadMultiple.Main", _
    "ReadMultipleValues.DataType", "ReadMultipleValues.Main", _
    "ReadValue.Main", _
    "SubscribeDataChange.Filter", "SubscribeDataChange.Main", _
    "SubscribeMultipleMonitoredItems.Filter", _
    "SubscribeMultipleMonitoredItems.Main", _
    "SubscribeMultipleMonitoredItems.StateAsInteger", _
    "UnsubscribeAllMonitoredItems.Main", "UnsubscribeMultipleMonitoredItems.Some", _
    "Write.Main", _
    "WriteMultipleValues.TestSuccess", _
    "WriteMultipleValues.ValueTypeCode", "WriteMultipleValues.ValueTypeFullName", _
    "WriteValue.ByteString", "WriteValue.Main", _
    "WriteValue.TypeCode" _
    )
    
    For Each action In actions
        If index > 0 Then
            Load Command(index)
        End If
        Command(index).Caption = action
        Command(index).top = top
        Command(index).Left = 120
        Command(index).Height = 250 ' 375
        Command(index).Width = 3615
        Command(index).TabIndex = index + 1
        Command(index).Visible = True

        index = index + 1
        top = top + 240 ' 360
    Next
    
End Sub

Private Sub Command_Click(index As Integer)
    Dim subName As String
    subName = Replace((Command(index).Caption), ".", "_") & "_Command_Click"
    CallByName Me, subName, VbMethod
End Sub

REM #region Example BrowseDataNodes.Main
REM This example shows how to obtain all data nodes (objects and variables) under a given node of the OPC-UA address space.
REM For each node, it displays its browse name and node ID.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Public Sub BrowseDataNodes_Main_Command_Click()
    OutputText = ""

    ' Instantiate the client object
    Dim Client As New EasyUAClient

    ' Perform the operation
    On Error Resume Next
    Dim NodeElements As UANodeElementCollection
    Set NodeElements = Client.BrowseDataNodes("opc.tcp://opcua.demo-this.com:51210/UA/SampleServer", "nsu=http://test.org/UA/Data/ ;i=10791")
    If Err.Number <> 0 Then
        OutputText = OutputText & "*** Failure: " & Err.Source & ": " & Err.Description & vbCrLf
        Exit Sub
    End If
    On Error GoTo 0

    ' Display results
    Dim NodeElement: For Each NodeElement In NodeElements
        OutputText = OutputText & NodeElement.BrowseName & ": " & NodeElement.NodeId & vbCrLf
    Next
End Sub
REM #endregion Example BrowseDataNodes.Main

REM #region Example BrowseDataVariables.Main
REM This example shows how to obtain data variables under the "Server" node
REM in the address space.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Public Sub BrowseDataVariables_Main_Command_Click()
    OutputText = ""

    Dim endpointDescriptor As String
    'endpointDescriptor = "http://opcua.demo-this.com:51211/UA/SampleServer"
    'endpointDescriptor = "https://opcua.demo-this.com:51212/UA/SampleServer/"
    endpointDescriptor = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"

    ' Instantiate the client object
    Dim Client As New EasyUAClient

    ' Obtain variables under "Server" node
    Dim serverNodeId As New UANodeId
    serverNodeId.StandardName = "Server"
    
    On Error Resume Next
    Dim NodeElements As UANodeElementCollection
    Set NodeElements = Client.BrowseDataVariables(endpointDescriptor, serverNodeId.expandedText)
    If Err.Number <> 0 Then
        OutputText = OutputText & "*** Failure: " & Err.Source & ": " & Err.Description & vbCrLf
        Exit Sub
    End If
    On Error GoTo 0

    ' Display results
    Dim NodeElement: For Each NodeElement In NodeElements
        OutputText = OutputText & vbCrLf
        OutputText = OutputText & "nodeElement.NodeId: " & NodeElement.NodeId & vbCrLf
        OutputText = OutputText & "nodeElement.NodeId.ExpandedText: " & NodeElement.NodeId.expandedText & vbCrLf
        OutputText = OutputText & "nodeElement.DisplayName: " & NodeElement.DisplayName & vbCrLf
    Next
    
    ' Example output:
    '
    'nodeElement.NodeId: Server_ServerStatus
    'nodeElement.NodeId.ExpandedText: nsu=http://opcfoundation.org/UA/ ;i=2256
    'nodeElement.DisplayName: ServerStatus
End Sub
REM #endregion Example BrowseDataVariables.Main

REM #region Example BrowseObjects.Main
REM This example shows how to obtain objects under the "Server" node
REM in the address space.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Public Sub BrowseObjects_Main_Command_Click()
    OutputText = ""

    Dim endpointDescriptor As String
    'endpointDescriptor = "http://opcua.demo-this.com:51211/UA/SampleServer"
    'endpointDescriptor = "https://opcua.demo-this.com:51212/UA/SampleServer/"
    endpointDescriptor = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"

    ' Instantiate the client object
    Dim Client As New EasyUAClient

    ' Obtain objects under "Server" node
    Dim serverNodeId As New UANodeId
    serverNodeId.StandardName = "Server"
    
    On Error Resume Next
    Dim NodeElements As UANodeElementCollection
    Set NodeElements = Client.BrowseObjects(endpointDescriptor, serverNodeId.expandedText)
    If Err.Number <> 0 Then
        OutputText = OutputText & "*** Failure: " & Err.Source & ": " & Err.Description & vbCrLf
        Exit Sub
    End If
    On Error GoTo 0

    ' Display results
    Dim NodeElement: For Each NodeElement In NodeElements
        OutputText = OutputText & vbCrLf
        OutputText = OutputText & "nodeElement.NodeId: " & NodeElement.NodeId & vbCrLf
        OutputText = OutputText & "nodeElement.NodeId.ExpandedText: " & NodeElement.NodeId.expandedText & vbCrLf
        OutputText = OutputText & "nodeElement.DisplayName: " & NodeElement.DisplayName & vbCrLf
    Next
    
    ' Example output:
    '
    'nodeElement.NodeId: Server_ServerCapabilities
    'nodeElement.NodeId.ExpandedText: nsu=http://opcfoundation.org/UA/ ;i=2268
    'nodeElement.DisplayName: ServerCapabilities
    '
    'nodeElement.NodeId: Server_ServerDiagnostics
    'nodeElement.NodeId.ExpandedText: nsu=http://opcfoundation.org/UA/ ;i=2274
    'nodeElement.DisplayName: ServerDiagnostics
    '
    'nodeElement.NodeId: Server_VendorServerInfo
    'nodeElement.NodeId.ExpandedText: nsu=http://opcfoundation.org/UA/ ;i=2295
    'nodeElement.DisplayName: VendorServerInfo
    '
    'nodeElement.NodeId: Server_ServerRedundancy
    'nodeElement.NodeId.ExpandedText: nsu=http://opcfoundation.org/UA/ ;i=2296
    'nodeElement.DisplayName: ServerRedundancy
    '
    'nodeElement.NodeId: Server_Namespaces
    'nodeElement.NodeId.ExpandedText: nsu=http://opcfoundation.org/UA/ ;i=11715
    'nodeElement.DisplayName: Namespaces
End Sub
REM #endregion Example BrowseObjects.Main

REM #region Example BrowseProperties.Main
REM This example shows how to obtain properties under the "Server" node
REM in the address space.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Public Sub BrowseProperties_Main_Command_Click()
    OutputText = ""

    Dim endpointDescriptor As String
    'endpointDescriptor = "http://opcua.demo-this.com:51211/UA/SampleServer"
    'endpointDescriptor = "https://opcua.demo-this.com:51212/UA/SampleServer/"
    endpointDescriptor = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"

    ' Instantiate the client object
    Dim Client As New EasyUAClient

    ' Obtain properties under "Server" node
    Dim serverNodeId As New UANodeId
    serverNodeId.StandardName = "Server"
    
    On Error Resume Next
    Dim NodeElements As UANodeElementCollection
    Set NodeElements = Client.BrowseProperties(endpointDescriptor, serverNodeId.expandedText)
    If Err.Number <> 0 Then
        OutputText = OutputText & "*** Failure: " & Err.Source & ": " & Err.Description & vbCrLf
        Exit Sub
    End If
    On Error GoTo 0

    ' Display results
    Dim NodeElement: For Each NodeElement In NodeElements
        OutputText = OutputText & vbCrLf
        OutputText = OutputText & "nodeElement.NodeId: " & NodeElement.NodeId & vbCrLf
        OutputText = OutputText & "nodeElement.NodeId.ExpandedText: " & NodeElement.NodeId.expandedText & vbCrLf
        OutputText = OutputText & "nodeElement.DisplayName: " & NodeElement.DisplayName & vbCrLf
    Next
    
    ' Example output:
    '
    'nodeElement.NodeId: Server_ServerArray
    'nodeElement.NodeId.ExpandedText: nsu=http://opcfoundation.org/UA/ ;i=2254
    'nodeElement.DisplayName: ServerArray
    '
    'nodeElement.NodeId: Server_NamespaceArray
    'nodeElement.NodeId.ExpandedText: nsu=http://opcfoundation.org/UA/ ;i=2255
    'nodeElement.DisplayName: NamespaceArray
    '
    'nodeElement.NodeId: Server_ServiceLevel
    'nodeElement.NodeId.ExpandedText: nsu=http://opcfoundation.org/UA/ ;i=2267
    'nodeElement.DisplayName: ServiceLevel
    '
    'nodeElement.NodeId: Server_Auditing
    'nodeElement.NodeId.ExpandedText: nsu=http://opcfoundation.org/UA/ ;i=2994
    'nodeElement.DisplayName: Auditing
End Sub
REM #endregion Example BrowseProperties.Main


REM #region Example BrowseMethods.Main
REM This example shows how to obtain all method nodes under a given node of the OPC-UA address space.
REM For each node, it displays its browse name and node ID.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Public Sub BrowseMethods_Main_Command_Click()
    OutputText = ""

    ' Instantiate the client object
    Dim Client As New EasyUAClient

    ' Perform the operation
    On Error Resume Next
    Dim NodeElements As UANodeElementCollection
    Set NodeElements = Client.BrowseMethods("opc.tcp://opcua.demo-this.com:51210/UA/SampleServer", "nsu=http://test.org/UA/Data/ ;i=10755")
    If Err.Number <> 0 Then
        OutputText = OutputText & "*** Failure: " & Err.Source & ": " & Err.Description & vbCrLf
        Exit Sub
    End If
    On Error GoTo 0

    ' Display results
    Dim NodeElement: For Each NodeElement In NodeElements
        OutputText = OutputText & NodeElement.BrowseName & ": " & NodeElement.NodeId & vbCrLf
    Next
End Sub
REM #endregion Example BrowseMethods.Main

REM #region Example Browse.Main
REM This example shows how to obtain nodes under a given node of the OPC-UA address space.
REM For each node, it displays its browse name and node ID.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Public Sub Browse_Main_Command_Click()
    OutputText = ""

    Dim endpointDescriptor As New UAEndpointDescriptor
    endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"

    Dim nodeDescriptor As New UANodeDescriptor
    Dim BrowsePathParser As New UABrowsePathParser
    BrowsePathParser.DefaultNamespaceUriString = "http://test.org/UA/Data/"
    Set nodeDescriptor.browsePath = BrowsePathParser.Parse("[ObjectsFolder]/Data/Static/UserScalar")

    Dim browseParameters As New UABrowseParameters
    browseParameters.StandardName = "AllForwardReferences"

    ' Instantiate the client object
    Dim Client As New EasyUAClient

    ' Perform the operation
    On Error Resume Next
    Dim NodeElements As UANodeElementCollection
    Set NodeElements = Client.Browse(endpointDescriptor, nodeDescriptor, browseParameters)
    If Err.Number <> 0 Then
        OutputText = OutputText & "*** Failure: " & Err.Source & ": " & Err.Description & vbCrLf
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Display results
    Dim NodeElement: For Each NodeElement In NodeElements
        OutputText = OutputText & NodeElement.BrowseName & ": " & NodeElement.NodeId & vbCrLf
    Next
End Sub
REM #endregion Example Browse.Main

REM #region Example CallMethod.Main
REM This example shows how to call a single method, and pass arguments to and from it.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Public Sub CallMethod_Main_Command_Click()
    OutputText = ""

    Dim inputs(10)
    inputs(0) = False
    inputs(1) = 1
    inputs(2) = 2
    inputs(3) = 3
    inputs(4) = 4
    inputs(5) = 5
    inputs(6) = 6
    inputs(7) = 7
    inputs(8) = 8
    inputs(9) = 9
    inputs(10) = 10

    Dim typeCodes(10)
    typeCodes(0) = 3    ' TypeCode.Boolean
    typeCodes(1) = 5    ' TypeCode.SByte
    typeCodes(2) = 6    ' TypeCode.Byte
    typeCodes(3) = 7    ' TypeCode.Int16
    typeCodes(4) = 8    ' TypeCode.UInt16
    typeCodes(5) = 9    ' TypeCode.Int32
    typeCodes(6) = 10   ' TypeCode.UInt32
    typeCodes(7) = 11   ' TypeCode.Int64
    typeCodes(8) = 12   ' TypeCode.UInt64
    typeCodes(9) = 13   ' TypeCode.Single
    typeCodes(10) = 14  ' TypeCode.Double

    ' Instantiate the client object
    Dim Client As New EasyUAClient

    ' Perform the operation
    On Error Resume Next
    Dim outputs As Variant
    outputs = Client.CallMethod( _
        "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer", _
        "nsu=http://test.org/UA/Data/ ;i=10755", _
        "nsu=http://test.org/UA/Data/ ;i=10756", _
        inputs, _
        typeCodes)
    If Err.Number <> 0 Then
        OutputText = OutputText & "*** Failure: " & Err.Source & ": " & Err.Description & vbCrLf
        Exit Sub
    End If
    On Error GoTo 0

    ' Display results
    Dim i: For i = LBound(outputs) To UBound(outputs)
        On Error Resume Next
        OutputText = OutputText & "outputs(" & i & "): " & outputs(i) & vbCrLf
        If Err <> 0 Then OutputText = OutputText & "*** Error" & vbCrLf ' occurrs with types not recognized by VB6
        On Error GoTo 0
    Next
End Sub
REM #endregion Example CallMethod.Main

REM #region Example CallMultipleMethods.Main
REM This example shows how to call multiple methods, and pass arguments to and from them.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Public Sub CallMultipleMethods_Main_Command_Click()
    OutputText = ""

    Dim inputs1(10)
    inputs1(0) = False
    inputs1(1) = 1
    inputs1(2) = 2
    inputs1(3) = 3
    inputs1(4) = 4
    inputs1(5) = 5
    inputs1(6) = 6
    inputs1(7) = 7
    inputs1(8) = 8
    inputs1(9) = 9
    inputs1(10) = 10

    Dim typeCodes1(10)
    typeCodes1(0) = 3    ' TypeCode.Boolean
    typeCodes1(1) = 5    ' TypeCode.SByte
    typeCodes1(2) = 6    ' TypeCode.Byte
    typeCodes1(3) = 7    ' TypeCode.Int16
    typeCodes1(4) = 8    ' TypeCode.UInt16
    typeCodes1(5) = 9    ' TypeCode.Int32
    typeCodes1(6) = 10   ' TypeCode.UInt32
    typeCodes1(7) = 11   ' TypeCode.Int64
    typeCodes1(8) = 12   ' TypeCode.UInt64
    typeCodes1(9) = 13   ' TypeCode.Single
    typeCodes1(10) = 14  ' TypeCode.Double

    Dim CallArguments1 As New UACallArguments
    CallArguments1.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    CallArguments1.nodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;i=10755"
    CallArguments1.MethodNodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;i=10756"
    ' Use SetXXXX methods instead of array-type property setters in Visual Basic 6.0
    CallArguments1.SetInputArguments inputs1
    CallArguments1.SetInputTypeCodes typeCodes1

    Dim inputs2(11)
    inputs2(0) = False
    inputs2(1) = 1
    inputs2(2) = 2
    inputs2(3) = 3
    inputs2(4) = 4
    inputs2(5) = 5
    inputs2(6) = 6
    inputs2(7) = 7
    inputs2(8) = 8
    inputs2(9) = 9
    inputs2(10) = 10
    inputs2(11) = "eleven"

    Dim typeCodes2(11)
    typeCodes2(0) = 3    ' TypeCode.Boolean
    typeCodes2(1) = 5    ' TypeCode.SByte
    typeCodes2(2) = 6    ' TypeCode.Byte
    typeCodes2(3) = 7    ' TypeCode.Int16
    typeCodes2(4) = 8    ' TypeCode.UInt16
    typeCodes2(5) = 9    ' TypeCode.Int32
    typeCodes2(6) = 10   ' TypeCode.UInt32
    typeCodes2(7) = 11   ' TypeCode.Int64
    typeCodes2(8) = 12   ' TypeCode.UInt64
    typeCodes2(9) = 13   ' TypeCode.Single
    typeCodes2(10) = 14  ' TypeCode.Double
    typeCodes2(11) = 18  ' TypeCode.String

    Dim CallArguments2 As New UACallArguments
    CallArguments2.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    CallArguments2.nodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;i=10755"
    CallArguments2.MethodNodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;i=10774"
    ' Use SetXXXX methods instead of array-type property setters in Visual Basic 6.0
    CallArguments2.SetInputArguments inputs2
    CallArguments2.SetInputTypeCodes typeCodes2

    Dim arguments(1) As Variant
    Set arguments(0) = CallArguments1
    Set arguments(1) = CallArguments2

    ' Instantiate the client object
    Dim Client As New EasyUAClient

    ' Perform the operation
    Dim results As Variant
    results = Client.CallMultipleMethods(arguments)

    ' Display results
    Dim i: For i = LBound(results) To UBound(results)
        OutputText = OutputText & vbCrLf
        OutputText = OutputText & "results(" & i & "):" & vbCrLf
        Dim Result As ValueArrayResult: Set Result = results(i)

        If Result.Exception Is Nothing Then
            Dim outputs As Variant: outputs = Result.ValueArray
            Dim j: For j = LBound(outputs) To UBound(outputs)
                On Error Resume Next
                OutputText = OutputText & Space(4) & "outputs(" & j & "): " & outputs(j) & vbCrLf
                If Err <> 0 Then OutputText = OutputText & Space(4) & "*** Error" & vbCrLf ' occurrs with types not recognized by VB6
                On Error GoTo 0
            Next
        Else
            OutputText = OutputText & "*** Error: " & Result.Exception & vbCrLf
        End If
    Next
End Sub
REM #endregion Example CallMultipleMethods.Main

REM #region Example ChangeMonitoredItemSubscription.Main
REM This example shows how to change the sampling rate of an existing monitored
REM item subscription.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

' The client object, with events
'Public WithEvents Client5 As EasyUAClient

Public Sub ChangeMonitoredItemSubscription_Main_Command_Click()
    OutputText = ""
    
    Set Client5 = New EasyUAClient

    OutputText = OutputText & "Subscribing..." & vbCrLf
    Dim Handle As Long
    Handle = Client5.SubscribeDataChange("opc.tcp://opcua.demo-this.com:51210/UA/SampleServer", "nsu=http://test.org/UA/Data/ ;i=10853", 1000)

    OutputText = OutputText & "Processing monitored item changed events for 10 seconds..." & vbCrLf
    Pause 10000

    Call Client5.ChangeMonitoredItemSubscription(Handle, 100)

    OutputText = OutputText & "Processing monitored item changed events for 10 seconds..." & vbCrLf
    Pause 10000

    OutputText = OutputText & "Unsubscribing..." & vbCrLf
    Client5.UnsubscribeAllMonitoredItems

    OutputText = OutputText & "Waiting for 5 seconds..." & vbCrLf
    Pause 5000

    OutputText = OutputText & "Finished." & vbCrLf
    Set Client5 = Nothing
End Sub

Public Sub Client5_DataChangeNotification(ByVal sender As Variant, ByVal eventArgs As EasyUADataChangeNotificationEventArgs)
    ' Display the data
    If eventArgs.Exception Is Nothing Then
        OutputText = OutputText & eventArgs.AttributeData & vbCrLf
    Else
        OutputText = OutputText & eventArgs.ErrorMessageBrief & vbCrLf
    End If
End Sub
REM #endregion Example ChangeMonitoredItemSubscription.Main

REM #region Example DiscoverGlobalServers.Main
REM This example shows how to obtain information about OPC UA servers from the Global Discovery Server (GDS).
REM The result is flat, i.e. each discovery URL is returned in separate element, with possible repetition of the servers.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Public Sub DiscoverGlobalServers_Main_Command_Click()
    OutputText = ""

    ' Instantiate the client object
    Dim Client As New EasyUAClient

    ' Obtain collection of application elements
    On Error Resume Next
    Dim DiscoveryElementCollection As UADiscoveryElementCollection
    Set DiscoveryElementCollection = Client.DiscoverGlobalServers("opc.tcp://opcua.demo-this.com:58810/GlobalDiscoveryServer")
    If Err.Number <> 0 Then
        OutputText = OutputText & "*** Failure: " & Err.Source & ": " & Err.Description & vbCrLf
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Display results
    Dim DiscoveryElement As UADiscoveryElement: For Each DiscoveryElement In DiscoveryElementCollection
        OutputText = OutputText & vbCrLf
        OutputText = OutputText & "Server name: " & DiscoveryElement.ServerName & vbCrLf
        OutputText = OutputText & "Discovery URI string: " & DiscoveryElement.DiscoveryUriString & vbCrLf
        OutputText = OutputText & "Server capabilities: " & DiscoveryElement.serverCapabilities.ToString & vbCrLf
    Next
End Sub
REM #endregion Example DiscoverGlobalServers.Main

REM #region Example DiscoverLocalServers.Main
REM This example shows how to obtain application URLs of all OPC Unified Architecture servers on the specified host.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Public Sub DiscoverLocalServers_Main_Command_Click()
    OutputText = ""

    ' Instantiate the client object
    Dim Client As New EasyUAClient

    ' Obtain collection of server elements
    On Error Resume Next
    Dim DiscoveryElementCollection As UADiscoveryElementCollection
    Set DiscoveryElementCollection = Client.DiscoverLocalServers("opcua.demo-this.com")
    If Err.Number <> 0 Then
        OutputText = OutputText & "*** Failure: " & Err.Source & ": " & Err.Description & vbCrLf
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Display results
    Dim DiscoveryElement As UADiscoveryElement: For Each DiscoveryElement In DiscoveryElementCollection
        OutputText = OutputText & "DiscoveryElementCollection[""" & DiscoveryElement.DiscoveryUriString & """].ApplicationUriString: " & _
            DiscoveryElement.applicationUriString & vbCrLf
    Next
End Sub
REM #endregion Example DiscoverLocalServers.Main

REM #region Example DiscoverNetworkServers.Main
REM This example shows how to obtain information about OPC UA servers available on the network.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Public Sub DiscoverNetworkServers_Main_Command_Click()
    OutputText = ""

    ' Instantiate the client object
    Dim Client As New EasyUAClient

    ' Obtain collection of application elements
    On Error Resume Next
    Dim DiscoveryElementCollection As UADiscoveryElementCollection
    Set DiscoveryElementCollection = Client.DiscoverNetworkServers(Nothing, "opcua.demo-this.com")
    If Err.Number <> 0 Then
        OutputText = OutputText & "*** Failure: " & Err.Source & ": " & Err.Description & vbCrLf
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Display results
    Dim DiscoveryElement As UADiscoveryElement: For Each DiscoveryElement In DiscoveryElementCollection
        OutputText = OutputText & vbCrLf
        OutputText = OutputText & "Server name: " & DiscoveryElement.ServerName & vbCrLf
        OutputText = OutputText & "Discovery URI string: " & DiscoveryElement.DiscoveryUriString & vbCrLf
        OutputText = OutputText & "Server capabilities: " & DiscoveryElement.serverCapabilities.ToString & vbCrLf
    Next
End Sub
REM #endregion Example DiscoverNetworkServers.Main

REM #region Example FindLocalApplications.Main
REM This example shows how to obtain application URLs of all OPC Unified Architecture servers, using specified discovery URI strings.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Public Sub FindLocalApplications_Main_Command_Click()
    OutputText = ""
    
    Dim DiscoveryUriStrings(2)
    DiscoveryUriStrings(0) = "opc.tcp://opcua.demo-this.com:4840/UADiscovery"
    DiscoveryUriStrings(1) = "http://opcua.demo-this.com/UADiscovery/Default.svc"
    DiscoveryUriStrings(2) = "http://opcua.demo-this.com:52601/UADiscovery"

    ' Instantiate the client object
    Dim Client As New EasyUAClient

    ' Obtain collection of application elements
    On Error Resume Next
    Dim DiscoveryElementCollection As UADiscoveryElementCollection
    Set DiscoveryElementCollection = Client.FindLocalApplications(DiscoveryUriStrings, UAApplicationTypes_Server)
    If Err.Number <> 0 Then
        OutputText = OutputText & "*** Failure: " & Err.Source & ": " & Err.Description & vbCrLf
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Display results
    Dim DiscoveryElement As UADiscoveryElement: For Each DiscoveryElement In DiscoveryElementCollection
        OutputText = OutputText & "DiscoveryElementCollection[""" & DiscoveryElement.DiscoveryUriString & """].ApplicationUriString: " & _
            DiscoveryElement.applicationUriString & vbCrLf
    Next
End Sub
REM #endregion Example FindLocalApplications.Main

REM #region Example PullDataChangeNotification.Main
REM This example shows how to subscribe to changes of a single monitored item, pull events, and display each change.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Public Sub PullDataChangeNotification_Main_Command_Click()
    OutputText = ""
    
    Dim eventArgs As EasyUADataChangeNotificationEventArgs
    
    ' Instantiate the client object
    Dim Client As New EasyUAClient
    
    ' In order to use event pull, you must set a non-zero queue capacity upfront.
    Client.PullDataChangeNotificationQueueCapacity = 1000
    
    OutputText = OutputText & "Subscribing..." & vbCrLf
    Call Client.SubscribeDataChange("opc.tcp://opcua.demo-this.com:51210/UA/SampleServer", "nsu=http://test.org/UA/Data/ ;i=10853", 1000)

    OutputText = OutputText & "Processing data changed notification events for 1 minute..." & vbCrLf
    
    Dim EndTick As Long
    EndTick = GetTickCount + 60000
    While GetTickCount < EndTick
        Set eventArgs = Client.PullDataChangeNotification(2 * 1000)
        If Not eventArgs Is Nothing Then
            ' Handle the notification event
            OutputText = OutputText & eventArgs & vbCrLf
        End If
    Wend
    
    OutputText = OutputText & "Unsubscribing..." & vbCrLf
    Client.UnsubscribeAllMonitoredItems

    OutputText = OutputText & "Finished." & vbCrLf
End Sub
REM #endregion Example PullDataChangeNotification.Main

REM #region Example Read.Main
REM This example shows how to read and display data of an attribute (value, timestamps, and status code).
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Public Sub Read_Main_Command_Click()
    OutputText = ""

    ' Instantiate the client object
    Dim Client As New EasyUAClient

    ' Obtain attribute data. By default, the Value attribute of a node will be read.
    On Error Resume Next
    Dim AttributeData As UAAttributeData
    Set AttributeData = Client.Read("opc.tcp://opcua.demo-this.com:51210/UA/SampleServer", _
                                    "nsu=http://test.org/UA/Data/ ;i=10853")
    If Err.Number <> 0 Then
        OutputText = OutputText & "*** Failure: " & Err.Source & ": " & Err.Description & vbCrLf
        Exit Sub
    End If
    On Error GoTo 0

    ' Display results
    OutputText = OutputText & "Value: " & AttributeData.value & vbCrLf
    OutputText = OutputText & "ServerTimestamp: " & AttributeData.ServerTimestamp & vbCrLf
    OutputText = OutputText & "SourceTimestamp: " & AttributeData.SourceTimestamp & vbCrLf
    OutputText = OutputText & "StatusCode: " & AttributeData.StatusCode & vbCrLf

    ' Example output:
    '
    'Value: -2.230064E-31
    'ServerTimestamp: 11/6/2011 1:34:30 PM
    'SourceTimestamp: 11/6/2011 1:34:30 PM
    'StatusCode: Good
End Sub
REM #endregion Example Read.Main

REM #region Example ReadMultiple.BrowsePath
REM This example shows how to read the attributes of 4 OPC-UA nodes specified
REM by browse paths at once, and display the results.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Public Sub ReadMultiple_BrowsePath_Command_Click()
    OutputText = ""

    Dim BrowsePathParser As New UABrowsePathParser
    BrowsePathParser.DefaultNamespaceUriString = "http://test.org/UA/Data/"

    Dim readArguments1 As New UAReadArguments
    readArguments1.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    ' Note: Add error handling around the following statement if the browse path is not guaranteed to be syntactically valid.
    Set readArguments1.nodeDescriptor.browsePath = BrowsePathParser.Parse("[ObjectsFolder]/Data/Dynamic/Scalar/FloatValue")

    Dim ReadArguments2 As New UAReadArguments
    ReadArguments2.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    ' Note: Add error handling around the following statement if the browse path is not guaranteed to be syntactically valid.
    Set ReadArguments2.nodeDescriptor.browsePath = BrowsePathParser.Parse("[ObjectsFolder]/Data/Dynamic/Scalar/SByteValue")

    Dim ReadArguments3 As New UAReadArguments
    ReadArguments3.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    ' Note: Add error handling around the following statement if the browse path is not guaranteed to be syntactically valid.
    Set ReadArguments3.nodeDescriptor.browsePath = BrowsePathParser.Parse("[ObjectsFolder]/Data/Static/Array/UInt16Value")

    Dim ReadArguments4 As New UAReadArguments
    ReadArguments4.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    ' Note: Add error handling around the following statement if the browse path is not guaranteed to be syntactically valid.
    Set ReadArguments4.nodeDescriptor.browsePath = BrowsePathParser.Parse("[ObjectsFolder]/Data/Static/UserScalar/Int32Value")

    Dim arguments(3) As Variant
    Set arguments(0) = readArguments1
    Set arguments(1) = ReadArguments2
    Set arguments(2) = ReadArguments3
    Set arguments(3) = ReadArguments4

    ' Instantiate the client object
    Dim Client As New EasyUAClient

    ' Obtain values. By default, the Value attributes of the nodes will be read.
    Dim results() As Variant
    results = Client.ReadMultiple(arguments)

    ' Display results
    Dim i: For i = LBound(results) To UBound(results)
        Dim Result As UAAttributeDataResult: Set Result = results(i)
        If Result.Succeeded Then
            OutputText = OutputText & "results(" & i & ").AttributeData: " & Result.AttributeData & vbCrLf
        Else
            OutputText = OutputText & "results(" & i & ") *** Failure: " & Result.ErrorMessageBrief & vbCrLf
        End If
    Next
    
    ' Example output:
    'results(0).AttributeData: 4.187603E+21 {System.Single} @2019-11-09T14:05:46.268 @@2019-11-09T14:05:46.268; Good
    'results(1).AttributeData: -98 {System.Int16} @2019-11-09T14:05:46.268 @@2019-11-09T14:05:46.268; Good
    'results(2).AttributeData: [58] {38240, 11129, 64397, 22845, 30525, ...} {System.Int32[]} @2019-11-09T14:00:07.543 @@2019-11-09T14:05:46.268; Good
    'results(3).AttributeData: 1280120396 {System.Int32} @2019-11-09T14:00:07.590 @@2019-11-09T14:05:46.268; Good
End Sub
REM #endregion Example ReadMultiple.BrowsePath

REM #region Example ReadMultiple.Main
REM This example shows how to read the attributes of 4 OPC-UA nodes at once, and
REM display the results.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Public Sub ReadMultiple_Main_Command_Click()
    OutputText = ""

    Dim readArguments1 As New UAReadArguments
    readArguments1.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    readArguments1.nodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;i=10853"

    Dim ReadArguments2 As New UAReadArguments
    ReadArguments2.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    ReadArguments2.nodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;i=10845"

    Dim ReadArguments3 As New UAReadArguments
    ReadArguments3.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    ReadArguments3.nodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;i=10304"

    Dim ReadArguments4 As New UAReadArguments
    ReadArguments4.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    ReadArguments4.nodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;i=10389"

    Dim arguments(3) As Variant
    Set arguments(0) = readArguments1
    Set arguments(1) = ReadArguments2
    Set arguments(2) = ReadArguments3
    Set arguments(3) = ReadArguments4

    ' Instantiate the client object
    Dim Client As New EasyUAClient

    ' Obtain values. By default, the Value attributes of the nodes will be read.
    Dim results() As Variant
    results = Client.ReadMultiple(arguments)

    ' Display results
    Dim i: For i = LBound(results) To UBound(results)
        Dim Result As UAAttributeDataResult: Set Result = results(i)
        If Result.Succeeded Then
            OutputText = OutputText & "results(" & i & ").AttributeData: " & Result.AttributeData & vbCrLf
        Else
            OutputText = OutputText & "results(" & i & ") *** Failure: " & Result.ErrorMessageBrief & vbCrLf
        End If
    Next
End Sub
REM #endregion Example ReadMultiple.Main

REM #region Example ReadMultipleValues.DataType
REM This example shows how to read the DataType attributes of 3 different nodes at once. Using the same method, it is also possible
REM to read multiple attributes of the same node.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Public Sub ReadMultipleValues_DataType_Command_Click()
    OutputText = ""
    
    ' Instantiate the client object
    Dim Client As New EasyUAClient

    Dim readArguments1 As New UAReadArguments
    readArguments1.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    readArguments1.nodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;i=10845"
    readArguments1.AttributeId = UAAttributeId_DataType

    Dim ReadArguments2 As New UAReadArguments
    ReadArguments2.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    ReadArguments2.nodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;i=10853"
    ReadArguments2.AttributeId = UAAttributeId_DataType

    Dim ReadArguments3 As New UAReadArguments
    ReadArguments3.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    ReadArguments3.nodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;i=10855"
    ReadArguments3.AttributeId = UAAttributeId_DataType

    Dim arguments(2) As Variant
    Set arguments(0) = readArguments1
    Set arguments(1) = ReadArguments2
    Set arguments(2) = ReadArguments3

    ' Obtain values.
    Dim results() As Variant
    results = Client.ReadMultipleValues(arguments)

    ' Display results
    Dim i: For i = LBound(results) To UBound(results)
        OutputText = OutputText & vbCrLf
        
        Dim Result As valueResult: Set Result = results(i)
        If Result.Succeeded Then
            OutputText = OutputText & "Value: " & Result.value & vbCrLf
            On Error Resume Next
            OutputText = OutputText & "Value.ExpandedText: " & Result.value.expandedText & vbCrLf
            OutputText = OutputText & "Value.NamespaceUriString: " & Result.value.NamespaceUriString & vbCrLf
            OutputText = OutputText & "Value.NamespaceIndex: " & Result.value.NamespaceIndex & vbCrLf
            OutputText = OutputText & "Value.NumericIdentifier: " & Result.value.NumericIdentifier & vbCrLf
            On Error GoTo 0
        Else
            OutputText = OutputText & "*** Failure: " & Result.ErrorMessageBrief & vbCrLf
        End If
    Next

    ' Example output:
    '
    'Value: SByte
    'Value.ExpandedText: nsu=http://opcfoundation.org/UA/ ;i=2
    'Value.NamespaceUriString: http://opcfoundation.org/UA/
    'Value.NamespaceIndex: 0
    'Value.NumericIdentifier: 2
    '
    'Value: Float
    'Value.ExpandedText: nsu=http://opcfoundation.org/UA/ ;i=10
    'Value.NamespaceUriString: http://opcfoundation.org/UA/
    'Value.NamespaceIndex: 0
    'Value.NumericIdentifier: 10
    '
    'Value: String
    'Value.ExpandedText: nsu=http://opcfoundation.org/UA/ ;i=12
    'Value.NamespaceUriString: http://opcfoundation.org/UA/
    'Value.NamespaceIndex: 0
    'Value.NumericIdentifier: 12
End Sub
REM #endregion Example ReadMultipleValues.DataType

REM #region Example ReadMultipleValues.Main
REM This example shows how to read the Value attributes of 3 different nodes at once. Using the same method, it is also possible
REM to read multiple attributes of the same node.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Public Sub ReadMultipleValues_Main_Command_Click()
    OutputText = ""

    ' Instantiate the client object
    Dim Client As New EasyUAClient

    Dim readArguments1 As New UAReadArguments
    readArguments1.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    readArguments1.nodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;i=10845"

    Dim ReadArguments2 As New UAReadArguments
    ReadArguments2.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    ReadArguments2.nodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;i=10853"

    Dim ReadArguments3 As New UAReadArguments
    ReadArguments3.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    ReadArguments3.nodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;i=10855"

    Dim arguments(2) As Variant
    Set arguments(0) = readArguments1
    Set arguments(1) = ReadArguments2
    Set arguments(2) = ReadArguments3

    ' Obtain values. By default, the Value attributes of the nodes will be read.
    Dim results() As Variant
    results = Client.ReadMultipleValues(arguments)

    ' Display results
    Dim i: For i = LBound(results) To UBound(results)
        Dim Result As valueResult: Set Result = results(i)
        If Result.Succeeded Then
            OutputText = OutputText & "Value: " & Result.value & vbCrLf
        Else
            OutputText = OutputText & "*** Failure: " & Result.ErrorMessageBrief & vbCrLf
        End If
    Next

    ' Example output:
    '
    'Value: 8
    'Value: -8.06803E+21
    'Value: Strawberry Pig Banana Snake Mango Purple Grape Monkey Purple? Blueberry Lemon^
End Sub
REM #endregion Example ReadMultipleValues.Main

REM #region Example ReadValue.Main
REM This example shows how to read value of a single node, and display it.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Public Sub ReadValue_Main_Command_Click()
    OutputText = ""

    ' Instantiate the client object
    Dim Client As New EasyUAClient

    ' Perform the operation
    On Error Resume Next
    Dim value As Variant
    value = Client.ReadValue("opc.tcp://opcua.demo-this.com:51210/UA/SampleServer", "nsu=http://test.org/UA/Data/ ;i=10853")
    If Err.Number <> 0 Then
        OutputText = OutputText & "*** Failure: " & Err.Source & ": " & Err.Description & vbCrLf
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Display results
    OutputText = OutputText & "value: " & value & vbCrLf
End Sub
REM #endregion Example ReadValue.Main

REM #region Example SubscribeDataChange.Filter
REM This example shows how to subscribe to changes of a monitored item with data change filter.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

' The client object, with events
'Public WithEvents Client3 As EasyUAClient

Public Sub SubscribeDataChange_Filter_Command_Click()
    OutputText = ""

    Dim endpointDescriptor As String
    'endpointDescriptor = "http://opcua.demo-this.com:51211/UA/SampleServer"
    'endpointDescriptor = "https://opcua.demo-this.com:51212/UA/SampleServer/"
    endpointDescriptor = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"

    ' Instantiate the client object and hook events
    Set Client3 = New EasyUAClient
    
    ' Prepare the arguments.
    ' Report a notification if either the StatusCode or the value change.
    Dim DataChangeFilter As New UADataChangeFilter
    DataChangeFilter.Trigger = UADataChangeTrigger_StatusValue
    
    Dim MonitoringParameters As New UAMonitoringParameters
    Set MonitoringParameters.DataChangeFilter = DataChangeFilter
    MonitoringParameters.SamplingInterval = 1000
    
    Dim MonitoredItemArguments1 As New EasyUAMonitoredItemArguments
    MonitoredItemArguments1.endpointDescriptor.UrlString = endpointDescriptor
    MonitoredItemArguments1.nodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;i=10853"
    Set MonitoredItemArguments1.MonitoringParameters = MonitoringParameters

    Dim arguments(0) As Variant
    Set arguments(0) = MonitoredItemArguments1

    OutputText = OutputText & "Subscribing..." & vbCrLf
    Call Client3.SubscribeMultipleMonitoredItems(arguments)

    OutputText = OutputText & "Processing monitored item changed events for 20 seconds..." & vbCrLf
    Pause 20000

    OutputText = OutputText & "Unsubscribing..." & vbCrLf
    Client3.UnsubscribeAllMonitoredItems

    OutputText = OutputText & "Waiting for 5 seconds..." & vbCrLf
    Pause 5000

    Set Client3 = Nothing
End Sub

Public Sub Client3_DataChangeNotification(ByVal sender As Variant, ByVal eventArgs As EasyUADataChangeNotificationEventArgs)
    ' Display the data
    If eventArgs.Exception Is Nothing Then
        OutputText = OutputText & eventArgs.AttributeData & vbCrLf
    Else
        OutputText = OutputText & eventArgs.ErrorMessageBrief & vbCrLf
    End If
End Sub
REM #endregion Example SubscribeDataChange.Filter

REM #region Example SubscribeDataChange.Main
REM This example shows how to subscribe to changes of a single monitored item and display each change.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

' The client object, with events
'Public WithEvents Client1 As EasyUAClient

Public Sub SubscribeDataChange_Main_Command_Click()
    OutputText = ""
    
    Set Client1 = New EasyUAClient

    OutputText = OutputText & "Subscribing..." & vbCrLf
    Call Client1.SubscribeDataChange("opc.tcp://opcua.demo-this.com:51210/UA/SampleServer", "nsu=http://test.org/UA/Data/ ;i=10853", 1000)

    OutputText = OutputText & "Processing data changed notification events for 1 minute..." & vbCrLf
    Pause 60000

    OutputText = OutputText & "Unsubscribing..." & vbCrLf
    Client1.UnsubscribeAllMonitoredItems

    OutputText = OutputText & "Waiting for 5 seconds..." & vbCrLf
    Pause 5000

    OutputText = OutputText & "Finished." & vbCrLf
    Set Client1 = Nothing
End Sub

Public Sub Client1_DataChangeNotification(ByVal sender As Variant, ByVal eventArgs As EasyUADataChangeNotificationEventArgs)
    ' Display the data
    If eventArgs.Exception Is Nothing Then
        OutputText = OutputText & eventArgs.AttributeData & vbCrLf
    Else
        OutputText = OutputText & eventArgs.ErrorMessageBrief & vbCrLf
    End If
End Sub
REM #endregion Example SubscribeDataChange.Main

REM #region Example SubscribeMultipleMonitoredItems.Filter
REM This example shows how to subscribe to changes of multiple monitored items
REM and use a data change filter.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

' The client object, with events
'Public WithEvents Client4 As EasyUAClient

Public Sub SubscribeMultipleMonitoredItems_Filter_Command_Click()
    OutputText = ""

    ' Instantiate the client object and hook events
    Set Client4 = New EasyUAClient

    ' Prepare the arguments.
    ' Report a notification if either the StatusCode or the value change.
    Dim DataChangeFilter As New UADataChangeFilter
    DataChangeFilter.Trigger = UADataChangeTrigger_StatusValue
    '
    Dim MonitoringParameters As New UAMonitoringParameters
    Set MonitoringParameters.DataChangeFilter = DataChangeFilter
    MonitoringParameters.SamplingInterval = 1000
    
    Dim MonitoredItemArguments1 As New EasyUAMonitoredItemArguments
    MonitoredItemArguments1.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    MonitoredItemArguments1.nodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;i=10845"
    Set MonitoredItemArguments1.MonitoringParameters = MonitoringParameters
    MonitoredItemArguments1.SetState ("Item1")
    
    Dim MonitoredItemArguments2 As New EasyUAMonitoredItemArguments
    MonitoredItemArguments2.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    MonitoredItemArguments2.nodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;i=10853"
    Set MonitoredItemArguments2.MonitoringParameters = MonitoringParameters
    MonitoredItemArguments2.SetState ("Item2")
    
    Dim MonitoredItemArguments3 As New EasyUAMonitoredItemArguments
    MonitoredItemArguments3.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    MonitoredItemArguments3.nodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;i=10855"
    Set MonitoredItemArguments3.MonitoringParameters = MonitoringParameters
    MonitoredItemArguments3.SetState ("Item3")
    
    Dim arguments(2) As Variant
    Set arguments(0) = MonitoredItemArguments1
    Set arguments(1) = MonitoredItemArguments2
    Set arguments(2) = MonitoredItemArguments3
    
    OutputText = OutputText & "Subscribing..." & vbCrLf
    Dim handleArray As Variant
    handleArray = Client4.SubscribeMultipleMonitoredItems(arguments)

    Dim i As Long: For i = LBound(handleArray) To UBound(handleArray)
        OutputText = OutputText & "handleArray(" & i & "): " & handleArray(i) & vbCrLf
    Next

    OutputText = OutputText & "Processing monitored item changed events for 10 seconds..." & vbCrLf
    Pause 10000

    OutputText = OutputText & "Unsubscribing..." & vbCrLf
    Call Client4.UnsubscribeAllMonitoredItems

    OutputText = OutputText & "Waiting for 5 seconds..." & vbCrLf
    Pause 5000

    Set Client4 = Nothing
    OutputText = OutputText & "Finished." & vbCrLf
End Sub

Public Sub Client4_DataChangeNotification(ByVal sender As Variant, ByVal eventArgs As EasyUADataChangeNotificationEventArgs)
    ' Display the data
    If eventArgs.Exception Is Nothing Then
        OutputText = OutputText & "[" & eventArgs.arguments.State & "] " & eventArgs.arguments.nodeDescriptor & ": " & eventArgs.AttributeData & vbCrLf
    Else
        OutputText = OutputText & "[" & eventArgs.arguments.State & "] " & eventArgs.arguments.nodeDescriptor & ": " & eventArgs.ErrorMessageBrief & vbCrLf
    End If
End Sub
REM #endregion Example SubscribeMultipleMonitoredItems.Filter

REM #region Example SubscribeMultipleMonitoredItems.Main
REM This example shows how to subscribe to changes of multiple monitored items and display the value of the monitored item with
REM each change.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

' The client object, with events
'Public WithEvents Client2 As EasyUAClient

Public Sub SubscribeMultipleMonitoredItems_Main_Command_Click()
    OutputText = ""

    Set Client2 = New EasyUAClient

    OutputText = OutputText & "Subscribing..." & vbCrLf
    Dim MonitoringParameters As New UAMonitoringParameters
    MonitoringParameters.SamplingInterval = 1000
    
    Dim MonitoredItemArguments1 As New EasyUAMonitoredItemArguments
    MonitoredItemArguments1.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    MonitoredItemArguments1.nodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;i=10845"
    Set MonitoredItemArguments1.MonitoringParameters = MonitoringParameters
    MonitoredItemArguments1.SetState ("Item1")
    
    Dim MonitoredItemArguments2 As New EasyUAMonitoredItemArguments
    MonitoredItemArguments2.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    MonitoredItemArguments2.nodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;i=10853"
    Set MonitoredItemArguments2.MonitoringParameters = MonitoringParameters
    MonitoredItemArguments2.SetState ("Item2")
    
    Dim MonitoredItemArguments3 As New EasyUAMonitoredItemArguments
    MonitoredItemArguments3.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    MonitoredItemArguments3.nodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;i=10855"
    Set MonitoredItemArguments3.MonitoringParameters = MonitoringParameters
    MonitoredItemArguments3.SetState ("Item3")
    
    Dim arguments(2) As Variant
    Set arguments(0) = MonitoredItemArguments1
    Set arguments(1) = MonitoredItemArguments2
    Set arguments(2) = MonitoredItemArguments3
    Dim handleArray As Variant
    handleArray = Client2.SubscribeMultipleMonitoredItems(arguments)

    Dim i As Long: For i = LBound(handleArray) To UBound(handleArray)
        OutputText = OutputText & "handleArray(" & i & "): " & handleArray(i) & vbCrLf
    Next

    OutputText = OutputText & "Processing monitored item changed events for 10 seconds..." & vbCrLf
    Pause 10000

    OutputText = OutputText & "Unsubscribing..." & vbCrLf
    Call Client2.UnsubscribeAllMonitoredItems

    OutputText = OutputText & "Waiting for 5 seconds..." & vbCrLf
    Pause 5000

    Set Client2 = Nothing
    OutputText = OutputText & "Finished." & vbCrLf
End Sub

Public Sub Client2_DataChangeNotification(ByVal sender As Variant, ByVal eventArgs As EasyUADataChangeNotificationEventArgs)
    ' Display the data
    If eventArgs.Exception Is Nothing Then
        OutputText = OutputText & "[" & eventArgs.arguments.State & "] " & eventArgs.arguments.nodeDescriptor & ": " & eventArgs.AttributeData & vbCrLf
    Else
        OutputText = OutputText & "[" & eventArgs.arguments.State & "] " & eventArgs.arguments.nodeDescriptor & ": " & eventArgs.ErrorMessageBrief & vbCrLf
    End If
End Sub
REM #endregion Example SubscribeMultipleMonitoredItems.Main

REM #region Example SubscribeMultipleMonitoredItems.StateAsInteger
REM This example shows how to subscribe to changes of multiple monitored items
REM and display each change, identifying the different subscriptions by an
REM integer.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

' The client object, with events
'Public WithEvents Client6 As EasyUAClient

Public Sub SubscribeMultipleMonitoredItems_StateAsInteger_Command_Click()
    OutputText = ""

    Set Client6 = New EasyUAClient

    OutputText = OutputText & "Subscribing..." & vbCrLf
    Dim MonitoringParameters As New UAMonitoringParameters
    MonitoringParameters.SamplingInterval = 1000
    
    Dim MonitoredItemArguments1 As New EasyUAMonitoredItemArguments
    MonitoredItemArguments1.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    MonitoredItemArguments1.nodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;i=10845"
    Set MonitoredItemArguments1.MonitoringParameters = MonitoringParameters
    MonitoredItemArguments1.SetState 1 ' An integer we have chosen to identify the subscription
    
    Dim MonitoredItemArguments2 As New EasyUAMonitoredItemArguments
    MonitoredItemArguments2.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    MonitoredItemArguments2.nodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;i=10853"
    Set MonitoredItemArguments2.MonitoringParameters = MonitoringParameters
    MonitoredItemArguments2.SetState 2 ' An integer we have chosen to identify the subscription
    
    Dim MonitoredItemArguments3 As New EasyUAMonitoredItemArguments
    MonitoredItemArguments3.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    MonitoredItemArguments3.nodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;i=10855"
    Set MonitoredItemArguments3.MonitoringParameters = MonitoringParameters
    MonitoredItemArguments3.SetState 3 ' An integer we have chosen to identify the subscription
    
    Dim arguments(2) As Variant
    Set arguments(0) = MonitoredItemArguments1
    Set arguments(1) = MonitoredItemArguments2
    Set arguments(2) = MonitoredItemArguments3
    Dim handleArray As Variant
    handleArray = Client6.SubscribeMultipleMonitoredItems(arguments)

    Dim i As Long: For i = LBound(handleArray) To UBound(handleArray)
        OutputText = OutputText & "handleArray(" & i & "): " & handleArray(i) & vbCrLf
    Next

    OutputText = OutputText & "Processing monitored item changed events for 10 seconds..." & vbCrLf
    Pause 10000

    OutputText = OutputText & "Unsubscribing..." & vbCrLf
    Call Client6.UnsubscribeAllMonitoredItems

    OutputText = OutputText & "Waiting for 5 seconds..." & vbCrLf
    Pause 5000

    Set Client2 = Nothing
    OutputText = OutputText & "Finished." & vbCrLf
End Sub

Public Sub Client6_DataChangeNotification(ByVal sender As Variant, ByVal eventArgs As EasyUADataChangeNotificationEventArgs)
    ' Obtain the integer state we have passed in.
    Dim stateAsInteger As Integer: stateAsInteger = eventArgs.arguments.State
    If eventArgs.Succeeded Then
        OutputText = OutputText & stateAsInteger & ": " & eventArgs.AttributeData & vbCrLf
    Else
        OutputText = OutputText & stateAsInteger & " *** Failure: " & eventArgs.ErrorMessageBrief & vbCrLf
    End If
End Sub
REM #endregion Example SubscribeMultipleMonitoredItems.StateAsInteger

REM #region Example UnsubscribeAllMonitoredItems.Main
REM This example shows how to unsubscribe from changes of all items.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

' The client object, with events
'Public WithEvents Client7 As EasyUAClient

Public Sub UnsubscribeAllMonitoredItems_Main_Command_Click()
    OutputText = ""

    Set Client7 = New EasyUAClient

    OutputText = OutputText & "Subscribing..." & vbCrLf
    Dim MonitoringParameters As New UAMonitoringParameters
    MonitoringParameters.SamplingInterval = 1000
    
    Dim MonitoredItemArguments1 As New EasyUAMonitoredItemArguments
    MonitoredItemArguments1.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    MonitoredItemArguments1.nodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;i=10845"
    Set MonitoredItemArguments1.MonitoringParameters = MonitoringParameters
    MonitoredItemArguments1.SetState ("Item1")
    
    Dim MonitoredItemArguments2 As New EasyUAMonitoredItemArguments
    MonitoredItemArguments2.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    MonitoredItemArguments2.nodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;i=10853"
    Set MonitoredItemArguments2.MonitoringParameters = MonitoringParameters
    MonitoredItemArguments2.SetState ("Item2")
    
    Dim MonitoredItemArguments3 As New EasyUAMonitoredItemArguments
    MonitoredItemArguments3.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    MonitoredItemArguments3.nodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;i=10855"
    Set MonitoredItemArguments3.MonitoringParameters = MonitoringParameters
    MonitoredItemArguments3.SetState ("Item3")
    
    Dim arguments(2) As Variant
    Set arguments(0) = MonitoredItemArguments1
    Set arguments(1) = MonitoredItemArguments2
    Set arguments(2) = MonitoredItemArguments3
    Dim handleArray As Variant
    handleArray = Client7.SubscribeMultipleMonitoredItems(arguments)

    OutputText = OutputText & "Processing monitored item changed events for 10 seconds..." & vbCrLf
    Pause 10000

    OutputText = OutputText & "Unsubscribing..." & vbCrLf
    Call Client7.UnsubscribeAllMonitoredItems

    OutputText = OutputText & "Waiting for 5 seconds..." & vbCrLf
    Pause 5000

    Set Client7 = Nothing
    OutputText = OutputText & "Finished." & vbCrLf
End Sub

Public Sub Client7_DataChangeNotification(ByVal sender As Variant, ByVal eventArgs As EasyUADataChangeNotificationEventArgs)
    ' Display the data
    If eventArgs.Succeeded Then
        OutputText = OutputText & eventArgs.arguments.nodeDescriptor & ": " & eventArgs.AttributeData & vbCrLf
    Else
        OutputText = OutputText & eventArgs.arguments.nodeDescriptor & ": " & eventArgs.ErrorMessageBrief & vbCrLf
    End If
End Sub
REM #endregion Example UnsubscribeAllMonitoredItems.Main

REM #region Example UnsubscribeMultipleMonitoredItems.Some
REM This example shows how to unsubscribe from changes of just some monitored
REM items.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

' The client object, with events
'Public WithEvents Client8 As EasyUAClient

Public Sub UnsubscribeMultipleMonitoredItems_Some_Command_Click()
    OutputText = ""

    Set Client8 = New EasyUAClient

    OutputText = OutputText & "Subscribing..." & vbCrLf
    Dim MonitoringParameters As New UAMonitoringParameters
    MonitoringParameters.SamplingInterval = 1000
    
    Dim MonitoredItemArguments1 As New EasyUAMonitoredItemArguments
    MonitoredItemArguments1.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    MonitoredItemArguments1.nodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;i=10845"
    Set MonitoredItemArguments1.MonitoringParameters = MonitoringParameters
    MonitoredItemArguments1.SetState ("Item1")
    
    Dim MonitoredItemArguments2 As New EasyUAMonitoredItemArguments
    MonitoredItemArguments2.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    MonitoredItemArguments2.nodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;i=10853"
    Set MonitoredItemArguments2.MonitoringParameters = MonitoringParameters
    MonitoredItemArguments2.SetState ("Item2")
    
    Dim MonitoredItemArguments3 As New EasyUAMonitoredItemArguments
    MonitoredItemArguments3.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    MonitoredItemArguments3.nodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;i=10855"
    Set MonitoredItemArguments3.MonitoringParameters = MonitoringParameters
    MonitoredItemArguments3.SetState ("Item3")
    
    Dim arguments(2) As Variant
    Set arguments(0) = MonitoredItemArguments1
    Set arguments(1) = MonitoredItemArguments2
    Set arguments(2) = MonitoredItemArguments3
    Dim handleArray As Variant
    handleArray = Client8.SubscribeMultipleMonitoredItems(arguments)

    Dim i As Long: For i = LBound(handleArray) To UBound(handleArray)
        OutputText = OutputText & "handleArray(" & i & "): " & handleArray(i) & vbCrLf
    Next

    OutputText = OutputText & vbCrLf
    OutputText = OutputText & "Processing monitored item changed events for 10 seconds..." & vbCrLf
    Pause 10000

    OutputText = OutputText & vbCrLf
    OutputText = OutputText & "Unsubscribing from 2 monitored items..." & vbCrLf

    Dim SelectedHandles(1) As Variant
    ' We will unsubscribe from the first and third monitored item we have
    ' previously subscribed to.
    SelectedHandles(0) = handleArray(0)
    SelectedHandles(1) = handleArray(2)
    Client8.UnsubscribeMultipleMonitoredItems (SelectedHandles)

    OutputText = OutputText & vbCrLf
    OutputText = OutputText & "Processing monitored item changed events for 10 seconds..." & vbCrLf
    Pause 10000

    OutputText = OutputText & vbCrLf
    OutputText = OutputText & "Unsubscribing from all remaining monitored items..." & vbCrLf
    Call Client8.UnsubscribeAllMonitoredItems

    OutputText = OutputText & "Waiting for 5 seconds..." & vbCrLf
    Pause 5000

    Set Client8 = Nothing
    OutputText = OutputText & "Finished." & vbCrLf
End Sub

Public Sub Client8_DataChangeNotification(ByVal sender As Variant, ByVal eventArgs As EasyUADataChangeNotificationEventArgs)
    ' Display the data
    If eventArgs.Succeeded Then
        OutputText = OutputText & eventArgs.arguments.nodeDescriptor & ": " & eventArgs.AttributeData & vbCrLf
    Else
        OutputText = OutputText & eventArgs.arguments.nodeDescriptor & ": " & eventArgs.ErrorMessageBrief & vbCrLf
    End If
End Sub
REM #endregion Example UnsubscribeMultipleMonitoredItems.Some

REM #region Example Write.Main
REM This example shows how to write data (a value, timestamps and status code)
REM into a single attribute of a node.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Public Sub Write_Main_Command_Click()
    OutputText = ""

    Const GoodOrSuccess = 0

    ' Instantiate the client object
    Dim Client As New EasyUAClient

    ' Modify data of a node's attribute
    Dim StatusCode As New UAStatusCode
    StatusCode.Severity = GoodOrSuccess
    Dim AttributeData As New UAAttributeData
    AttributeData.SetValue 12345
    Set AttributeData.StatusCode = StatusCode
    AttributeData.SourceTimestamp = Now()
    
    ' Perform the operation
    On Error Resume Next
    Call Client.Write("opc.tcp://opcua.demo-this.com:51210/UA/SampleServer", "nsu=http://test.org/UA/Data/ ;i=10221", AttributeData)
    If Err.Number <> 0 Then
        OutputText = OutputText & "*** Failure: " & Err.Source & ": " & Err.Description & vbCrLf
        Exit Sub
    End If
    On Error GoTo 0
End Sub
REM #endregion Example Write.Main

REM #region Example WriteMultipleValues.TestSuccess
REM This example shows how to write values into 3 nodes at once, test for success of each write and display the exception
REM message in case of failure.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Public Sub WriteMultipleValues_TestSuccess_Command_Click()
    OutputText = ""

    ' Instantiate the client object
    Dim Client As New EasyUAClient

    Dim WriteValueArguments1 As New UAWriteValueArguments
    WriteValueArguments1.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    WriteValueArguments1.nodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;i=10221"
    WriteValueArguments1.SetValue 23456

    Dim WriteValueArguments2 As New UAWriteValueArguments
    WriteValueArguments2.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    WriteValueArguments2.nodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;i=10226"
    WriteValueArguments2.SetValue "This string cannot be converted to Double"

    Dim WriteValueArguments3 As New UAWriteValueArguments
    WriteValueArguments3.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    WriteValueArguments3.nodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;s=UnknownNode"
    WriteValueArguments3.SetValue "ABC"

    Dim arguments(2) As Variant
    Set arguments(0) = WriteValueArguments1
    Set arguments(1) = WriteValueArguments2
    Set arguments(2) = WriteValueArguments3

    ' Modify values of nodes
    Dim results As Variant
    results = Client.WriteMultipleValues(arguments)

    ' Display results
    Dim i: For i = LBound(results) To UBound(results)
        Dim Result As UAWriteResult: Set Result = results(i)
        If Result.Succeeded Then
            OutputText = OutputText & "Result " & i & " success" & vbCrLf
        Else
            OutputText = OutputText & "Result " & i & ": " & Result.Exception.GetBaseException().Message & vbCrLf
        End If
    Next
End Sub
REM #endregion Example WriteMultipleValues.TestSuccess

REM #region Example WriteMultipleValues.ValueTypeCode
REM This example shows how to write values into 3 nodes at once, specifying a type code explicitly. It tests for success of
REM each write and displays the exception message in case of failure.
REM
REM Reasons for specifying the type explicitly might be:
REM - The data type in the server has subtypes, and the client therefore needs to pick the subtype to be written.
REM - The data type that the reports is incorrect.
REM - Writing with an explicitly specified type is more efficient.
REM
REM Alternative ways of specifying the type are using the ValueType or ValueTypeFullName properties.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Public Sub WriteMultipleValues_ValueTypeCode_Command_Click()
    OutputText = ""

    Const TypeCode_Int32 = 9
    Const TypeCode_Double = 14
    Const TypeCode_String = 18

    ' Instantiate the client object
    Dim Client As New EasyUAClient

    Dim WriteValueArguments1 As New UAWriteValueArguments
    WriteValueArguments1.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    WriteValueArguments1.nodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;i=10221"
    WriteValueArguments1.SetValue 23456
    WriteValueArguments1.ValueTypeCode = TypeCode_Int32    ' here is the type explicitly specified

    Dim WriteValueArguments2 As New UAWriteValueArguments
    WriteValueArguments2.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    WriteValueArguments2.nodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;i=10226"
    WriteValueArguments2.SetValue "This string cannot be converted to Double"
    WriteValueArguments2.ValueTypeCode = TypeCode_Double    ' here is the type explicitly specified

    Dim WriteValueArguments3 As New UAWriteValueArguments
    WriteValueArguments3.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    WriteValueArguments3.nodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;s=UnknownNode"
    WriteValueArguments3.SetValue "ABC"
    WriteValueArguments3.ValueTypeCode = TypeCode_String    ' here is the type explicitly specified

    Dim arguments(2) As Variant
    Set arguments(0) = WriteValueArguments1
    Set arguments(1) = WriteValueArguments2
    Set arguments(2) = WriteValueArguments3

    ' Modify values of nodes
    Dim results As Variant
    results = Client.WriteMultipleValues(arguments)

    ' Display results
    Dim i: For i = LBound(results) To UBound(results)
        Dim Result As UAWriteResult: Set Result = results(i)
        If Result.Succeeded Then
            OutputText = OutputText & "Result " & i & " success" & vbCrLf
        Else
            OutputText = OutputText & "Result " & i & ": " & Result.Exception.GetBaseException().Message & vbCrLf
        End If
    Next
End Sub
REM #endregion Example WriteMultipleValues.ValueTypeCode

REM #region Example WriteMultipleValues.ValueTypeFullName
REM This example shows how to write values into 3 nodes at once, specifying a type's full name explicitly. It tests for
REM success of each write and displays the exception message in case of failure.
REM
REM Reasons for specifying the type explicitly might be:
REM - The data type in the server has subtypes, and the client therefore needs to pick the subtype to be written.
REM - The data type that the reports is incorrect.
REM - Writing with an explicitly specified type is more efficient.
REM
REM Alternative ways of specifying the type are using the ValueType or ValueTypeCode properties.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Public Sub WriteMultipleValues_ValueTypeFullName_Command_Click()
    OutputText = ""

    ' Instantiate the client object
    Dim Client As New EasyUAClient

    Dim WriteValueArguments1 As New UAWriteValueArguments
    WriteValueArguments1.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    WriteValueArguments1.nodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;i=10221"
    WriteValueArguments1.SetValue 23456
    WriteValueArguments1.ValueTypeFullName = "System.Int32"    ' here is the type explicitly specified

    Dim WriteValueArguments2 As New UAWriteValueArguments
    WriteValueArguments2.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    WriteValueArguments2.nodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;i=10226"
    WriteValueArguments2.SetValue "This string cannot be converted to Double"
    WriteValueArguments2.ValueTypeFullName = "System.Double"    ' here is the type explicitly specified

    Dim WriteValueArguments3 As New UAWriteValueArguments
    WriteValueArguments3.endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    WriteValueArguments3.nodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;s=UnknownNode"
    WriteValueArguments3.SetValue "ABC"
    WriteValueArguments3.ValueTypeFullName = "System.String"    ' here is the type explicitly specified

    Dim arguments(2) As Variant
    Set arguments(0) = WriteValueArguments1
    Set arguments(1) = WriteValueArguments2
    Set arguments(2) = WriteValueArguments3

    ' Modify values of nodes
    Dim results As Variant
    results = Client.WriteMultipleValues(arguments)

    ' Display results
    Dim i: For i = LBound(results) To UBound(results)
        Dim Result As UAWriteResult: Set Result = results(i)
        If Result.Succeeded Then
            OutputText = OutputText & "Result " & i & " success" & vbCrLf
        Else
            OutputText = OutputText & "Result " & i & ": " & Result.Exception.GetBaseException().Message & vbCrLf
        End If
    Next
End Sub
REM #endregion Example WriteMultipleValues.ValueTypeFullName

REM #region Example WriteValue.ByteString
REM This example shows how to write a value into a single node that is of type ByteString.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Public Sub WriteValue_ByteString_Command_Click()
    OutputText = ""

    Dim Values(4) As Byte
    Values(0) = 11
    Values(1) = 22
    Values(2) = 33
    Values(3) = 44
    Values(4) = 55

    ' Instantiate the client object
    Dim Client As New EasyUAClient

    ' Perform the operation
    On Error Resume Next
    Call Client.WriteValue("opc.tcp://opcua.demo-this.com:51210/UA/SampleServer", "nsu=http://test.org/UA/Data/ ;i=10230", Values)
    If Err.Number <> 0 Then
        OutputText = OutputText & "*** Failure: " & Err.Source & ": " & Err.Description & vbCrLf
        Exit Sub
    End If
    On Error GoTo 0

End Sub
REM #endregion Example WriteValue.ByteString

REM #region Example WriteValue.Main
REM This example shows how to write a value into a single node.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Public Sub WriteValue_Main_Command_Click()
    OutputText = ""

    ' Instantiate the client object
    Dim Client As New EasyUAClient

    ' Perform the operation
    On Error Resume Next
    Call Client.WriteValue("opc.tcp://opcua.demo-this.com:51210/UA/SampleServer", "nsu=http://test.org/UA/Data/ ;i=10221", 12345)
    If Err.Number <> 0 Then
        OutputText = OutputText & "*** Failure: " & Err.Source & ": " & Err.Description & vbCrLf
        Exit Sub
    End If
    On Error GoTo 0

End Sub
REM #endregion Example WriteValue.Main

REM #region Example WriteValue.TypeCode
REM This example shows how to write a value into a single node, specifying a type code explicitly.
REM
REM Reasons for specifying the type explicitly might be:
REM - The data type in the server has subtypes, and the client therefore needs to pick the subtype to be written.
REM - The data type that the reports is incorrect.
REM - Writing with an explicitly specified type is more efficient.
REM
REM TypeCode is easy to use, but it does not cover all possible types. It is also possible to specify the .NET Type, using
REM a different overload of the WriteValue method.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Public Sub WriteValue_TypeCode_Command_Click()
    OutputText = ""
    
    Const TypeCode_Int32 = 9

    Dim endpointDescriptor As String
    'endpointDescriptor = "http://opcua.demo-this.com:51211/UA/SampleServer"
    'endpointDescriptor = "https://opcua.demo-this.com:51212/UA/SampleServer/"
    endpointDescriptor = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"

    ' Instantiate the client object
    Dim Client As New EasyUAClient

    ' Prepare the arguments
    Dim WriteValueArguments1 As New UAWriteValueArguments
    WriteValueArguments1.endpointDescriptor.UrlString = endpointDescriptor
    WriteValueArguments1.nodeDescriptor.NodeId.expandedText = "nsu=http://test.org/UA/Data/ ;i=10221"
    WriteValueArguments1.SetValue 12345
    WriteValueArguments1.ValueTypeCode = TypeCode_Int32

    Dim arguments(0) As Variant
    Set arguments(0) = WriteValueArguments1

    ' Modify value of node
    Dim results As Variant
    results = Client.WriteMultipleValues(arguments)

    ' Display results
    Dim Result As UAWriteResult: Set Result = results(0)
    If Not Result.Succeeded Then
        OutputText = OutputText & "*** Failure: " & Result.Exception.GetBaseException().Message & vbCrLf
    End If
End Sub
REM #endregion Example WriteValue.TypeCode

