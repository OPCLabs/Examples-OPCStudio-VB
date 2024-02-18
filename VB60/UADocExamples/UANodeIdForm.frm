VERSION 5.00
Begin VB.Form UANodeIdForm 
   Caption         =   "UANodeId"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Construction_Main_Command 
      Caption         =   "_Construction.Main"
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
Attribute VB_Name = "UANodeIdForm"
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

REM #region Example _Construction.Main
REM This example shows different ways of constructing OPC UA node IDs.
REM
REM Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Private Sub Construction_Main_Command_Click()
    OutputText = ""
    
    ' A node ID specifies a namespace (either by an URI or by an index), and an identifier.
    ' The identifier can be numeric (an integer), string, GUID, or opaque.

    ' A node ID can be specified in string form (so-called expanded text).
    ' The code below specifies a namespace URI (nsu=...), and an integer identifier (i=...).
    ' Assigning an expanded text to a node ID parses the value being assigned and sets all corresponding
    ' properties accordingly.
    Dim nodeId1 As New UANodeId
    nodeId1.expandedText = "nsu=http://test.org/UA/Data/ ;i=10853"
    OutputText = OutputText & nodeId1 & vbCrLf


    ' Similarly, with a string identifier (s=...).
    Dim nodeId2 As New UANodeId
    nodeId2.expandedText = "nsu=http://test.org/UA/Data/ ;s=someIdentifier"
    OutputText = OutputText & nodeId2 & vbCrLf
    
    
    ' Actually, "s=" can be omitted (not recommended, though)
    Dim nodeId3 As New UANodeId
    nodeId3.expandedText = "nsu=http://test.org/UA/Data/ ;someIdentifier"
    OutputText = OutputText & nodeId3 & vbCrLf
    ' Notice that the output is normalized - the "s=" is added again.
     
    ' Similarly, with a GUID identifier (g=...)
    Dim nodeId4 As New UANodeId
    nodeId4.expandedText = "nsu=http://test.org/UA/Data/ ;g=BAEAF004-1E43-4A06-9EF0-E52010D5CD10"
    OutputText = OutputText & nodeId4 & vbCrLf
    ' Notice that the output is normalized - uppercase letters in the GUI are converted to lowercase, etc.
    
    
    ' Similarly, with an opaque identifier (b=..., in Base64 encoding).
    Dim nodeId5 As New UANodeId
    nodeId5.expandedText = "nsu=http://test.org/UA/Data/ ;b=AP8="
    OutputText = OutputText & nodeId5 & vbCrLf
       
       
    ' Namespace index can be used instead of namespace URI. The server is allowed to change the namespace
    ' indices between sessions (except for namespace 0), and for this reason, you should avoid the use of
    ' namespace indices, and rather use the namespace URIs whenever possible.
    Dim nodeId6 As New UANodeId
    nodeId6.expandedText = "ns=2;i=10853"
    OutputText = OutputText & nodeId6 & vbCrLf
       
       
    ' Namespace index can be also specified together with namespace URI. This is still safe, but may be
    ' a bit quicker to perform, because the client can just verify the namespace URI instead of looking
    ' it up.
    Dim nodeId7 As New UANodeId
    nodeId7.expandedText = "nsu=http://test.org/UA/Data/ ;ns=2;i=10853"
    OutputText = OutputText & nodeId7 & vbCrLf
       
       
    ' When neither namespace URI nor namespace index are given, the node ID is assumed to be in namespace
    ' with index 0 and URI "http://opcfoundation.org/UA/", which is reserved by OPC UA standard. There are
    ' many standard nodes that live in this reserved namespace, but no nodes specific to your servers will
    ' be in the reserved namespace, and hence the need to specify the namespace with server-specific nodes.
    Dim nodeId8 As New UANodeId
    nodeId8.expandedText = "i=2254"
    OutputText = OutputText & nodeId8 & vbCrLf
       
       
    ' If you attempt to pass in a string that does not conform to the syntax rules,
    ' a UANodeIdFormatException is thrown.
    Dim nodeId9 As New UANodeId
    On Error Resume Next
    nodeId9.expandedText = "nsu=http://test.org/UA/Data/ ;i=notAnInteger"
    OutputText = OutputText & nodeId9 & vbCrLf
    If Err.Number <> 0 Then
        OutputText = OutputText & "*** Failure: " & Err.Source & ": " & Err.Description & vbCrLf
    End If
    On Error GoTo 0


    ' There is a parser object that can be used to parse the expanded texts of node IDs.
    Dim nodeIdParser10 As New UANodeIdParser
    Dim nodeId10 As UANodeId
    Set nodeId10 = nodeIdParser10.Parse("nsu=http://test.org/UA/Data/ ;i=10853", False)
    OutputText = OutputText & nodeId10 & vbCrLf
       
       
    ' The parser can be used if you want to parse the expanded text of the node ID but do not want
    ' exceptions be thrown.
    Dim nodeIdParser11 As New UANodeIdParser
    Dim stringParsingError As stringParsingError
    Dim nodeId11 As UANodeId
    Set stringParsingError = nodeIdParser11.TryParse("nsu=http://test.org/UA/Data/ ;i=notAnInteger", False, nodeId11)
    If stringParsingError Is Nothing Then
        OutputText = OutputText & nodeId11 & vbCrLf
    Else
        OutputText = OutputText & "*** Failure: " & stringParsingError.Message & vbCrLf
    End If
       
       
    ' You can also use the parser if you have node IDs where you want the default namespace be different
    ' from the standard "http://opcfoundation.org/UA/".
    Dim nodeIdParser12 As New UANodeIdParser
    nodeIdParser12.DefaultNamespaceUriString = "http://test.org/UA/Data/"
    Dim nodeId12 As UANodeId
    Set nodeId12 = nodeIdParser12.Parse("i=10853", False)
    OutputText = OutputText & nodeId12 & vbCrLf
       
       
    ' You can create a "null" node ID. Such node ID does not actually identify any valid node in OPC UA, but
    ' is useful as a placeholder or as a starting point for further modifications of its properties.
    Dim nodeId14 As New UANodeId
    OutputText = OutputText & nodeId14 & vbCrLf
       
       
    ' If you know the type of the identifier upfront, it is safer to use typed properties that correspond
    ' to specific types of identifier. Here, with an integer identifier.
    Dim nodeId17 As New UANodeId
    nodeId17.NamespaceUriString = "http://test.org/UA/Data/"
    nodeId17.NumericIdentifier = 10853
    OutputText = OutputText & nodeId17 & vbCrLf
       
       
    ' Similarly, with a string identifier.
    Dim nodeId18 As New UANodeId
    nodeId18.NamespaceUriString = "http://test.org/UA/Data/"
    nodeId18.StringIdentifier = "someIdentifier"
    OutputText = OutputText & nodeId18 & vbCrLf
       
       
    ' If you have GUID in its string form, the node ID object can parse it for you.
    Dim nodeId20 As New UANodeId
    nodeId20.NamespaceUriString = "http://test.org/UA/Data/"
    nodeId20.GuidIdentifierString = "BAEAF004-1E43-4A06-9EF0-E52010D5CD10"
    OutputText = OutputText & nodeId20 & vbCrLf
       
       
    ' And, with an opaque identifier.
    Dim nodeId21 As New UANodeId
    nodeId21.NamespaceUriString = "http://test.org/UA/Data/"
    Dim opaqueIdentifier21(1) As Byte
    opaqueIdentifier21(0) = &H0&
    opaqueIdentifier21(1) = &HFF&
    nodeId21.SetOpaqueIdentifier opaqueIdentifier21
    OutputText = OutputText & nodeId21 & vbCrLf


    ' We have built-in a list of all standard nodes specified by OPC UA. You can simply refer to these node IDs in your code.
    ' You can refer to any standard node using its name (in a string form).
    ' Note that assigning a non-existing standard name is not allowed, and throws ArgumentException.
    Dim nodeId26 As New UANodeId
    nodeId26.StandardName = "TypesFolder"
    OutputText = OutputText & nodeId26 & vbCrLf
    ' When the UANodeId equals to one of the standard nodes, it is output in the shortened form - as the standard name only.
    
      
    ' When you browse for nodes in the OPC UA server, every returned node element contains a node ID that
    ' you can use further.
    Dim client27 As New EasyUAClient
    Dim endpointDescriptor As New UAEndpointDescriptor
    endpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    ' Browse from the Server node.
    Dim serverNodeId As New UANodeId
    serverNodeId.StandardName = "Server"
    Dim serverNodedescriptor As New UANodeDescriptor
    Set serverNodedescriptor.NodeId = serverNodeId
    ' Browse all References.
    Dim referencesNodeId As New UANodeId
    referencesNodeId.StandardName = "References"
    
    Dim browseParameters As New UABrowseParameters
    browseParameters.NodeClasses = UANodeClass_All ' this is the default, anyway
    browseParameters.ReferenceTypeIds.Add referencesNodeId
    
    On Error Resume Next
    Dim nodeElements27 As UANodeElementCollection
    Set nodeElements27 = client27.Browse(endpointDescriptor, serverNodedescriptor, browseParameters)
    If Err.Number <> 0 Then
        OutputText = OutputText & "*** Failure: " & Err.Source & ": " & Err.Description & vbCrLf
        Exit Sub
    End If
    On Error GoTo 0

    If nodeElements27.Count <> 0 Then
        Dim nodeId27 As UANodeId
        Set nodeId27 = nodeElements27(0).NodeId
        OutputText = OutputText & nodeId27 & vbCrLf
    End If
   
End Sub
REM #endregion Example _Construction.Main
