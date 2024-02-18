VERSION 5.00
Begin VB.Form DataAccessXmlForm 
   Caption         =   "DataAccess.Xml"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton WriteItemValue_MainXml_Command 
      Caption         =   "WriteItemValue_MainXml"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   3015
   End
   Begin VB.CommandButton SubscribeItem_MainXml_Command 
      Caption         =   "SubscribeItem.MainXml"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   3015
   End
   Begin VB.CommandButton ReadMultipleItems_MainXml_Command 
      Caption         =   "ReadMultipleItems.MainXml"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   3015
   End
   Begin VB.CommandButton PullItemChanged_MainXml_Command 
      Caption         =   "PullItemChanged.MainXml"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   3015
   End
   Begin VB.CommandButton GetMultiplePropertyValues_DataTypeXml_Command 
      Caption         =   "GetMultiplePropertyValues.DataTypeXml"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   3015
   End
   Begin VB.CommandButton ChangeItemSubscription_MainXml_Command 
      Caption         =   "ChangeItemSubscription.MainXml"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   3015
   End
   Begin VB.CommandButton BrowseNodes_RecursiveXml_Command 
      Caption         =   "BrowseNodes.RecursiveXml"
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
Attribute VB_Name = "DataAccessXmlForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Option Explicit

' The client object, with events
Public WithEvents Client1 As EasyDAClient
Attribute Client1.VB_VarHelpID = -1

' The client object, with events
Public WithEvents Client2 As EasyDAClient
Attribute Client2.VB_VarHelpID = -1

Public branchCount As Integer
Public leafCount As Integer

' Pause
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Public Sub Pause(Optional milliseconds As Long)
    On Error Resume Next
    Dim endTickCount As Long
    endTickCount = GetTickCount + milliseconds
    While GetTickCount < endTickCount: Sleep 1: DoEvents: Wend
End Sub

Rem #region Example BrowseNodes.RecursiveXml
Rem This example shows how to recursively browse the nodes in the OPC XML-DA address space.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

'Public branchCount As Integer
'Public leafCount As Integer

Private Sub BrowseNodes_RecursiveXml_Command_Click()
    OutputText = ""
    branchCount = 0
    leafCount = 0
    Dim beginTime: beginTime = Timer
        
    Dim serverDescriptor As New serverDescriptor
    serverDescriptor.UrlString = "http://opcxml.demo-this.com/XmlDaSampleServer/Service.asmx"
    
    Dim NodeDescriptor As New DANodeDescriptor
    NodeDescriptor.ItemId = ""
    
    ' Instantiate the client object
    Dim client As New EasyDAClient

    On Error Resume Next
        BrowseFromNode client, serverDescriptor, NodeDescriptor
    If Err.Number <> 0 Then
        OutputText = OutputText & "*** Failure: " & Err.Source & ": " & Err.Description & vbCrLf
        Exit Sub
    End If
    On Error GoTo 0

    Dim endTime: endTime = Timer
    
    OutputText = OutputText & vbCrLf
    OutputText = OutputText & "Browsing has taken (milliseconds): " & (endTime - beginTime) * 1000 & vbCrLf
    OutputText = OutputText & "Branch count: " & branchCount & vbCrLf
    OutputText = OutputText & "Leaf count: " & leafCount & vbCrLf
End Sub

Public Sub BrowseFromNode(client, serverDescriptor, ParentNodeDescriptor)
    ' Obtain all node elements under ParentNodeDescriptor
    Dim BrowseParameters As New DABrowseParameters
    Dim NodeElementCollection As DANodeElementCollection
    Set NodeElementCollection = client.BrowseNodes(serverDescriptor, ParentNodeDescriptor, BrowseParameters)
    ' Remark: that BrowseNodes(...) may also throw OpcException; a production code should contain handling for
    ' it, here omitted for brevity.

    Dim NodeElement: For Each NodeElement In NodeElementCollection
        OutputText = OutputText & NodeElement & vbCrLf
        
        ' If the node is a branch, browse recursively into it.
        If NodeElement.IsBranch Then
            branchCount = branchCount + 1
            BrowseFromNode client, serverDescriptor, NodeElement.ToDANodeDescriptor
        Else
            leafCount = leafCount + 1
        End If
    Next
End Sub
Rem #endregion Example BrowseNodes.RecursiveXml

Rem #region Example ChangeItemSubscription.MainXml
Rem This example shows how change the update rate of an existing OPC XML-DA subscription.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

' The client object, with events
'Public WithEvents Client1 As EasyDAClient

Private Sub ChangeItemSubscription_MainXml_Command_Click()
    OutputText = ""
    
    Dim ItemSubscriptionArguments1 As New EasyDAItemSubscriptionArguments
    ItemSubscriptionArguments1.serverDescriptor.UrlString = "http://opcxml.demo-this.com/XmlDaSampleServer/Service.asmx"
    ItemSubscriptionArguments1.ItemDescriptor.ItemId = "Dynamic/Analog Types/Int"
    ItemSubscriptionArguments1.GroupParameters.RequestedUpdateRate = 2000

    Dim arguments(0) As Variant
    Set arguments(0) = ItemSubscriptionArguments1

    ' Instantiate the client object and hook events
    Set Client1 = New EasyDAClient

    OutputText = OutputText & "Subscribing..." & vbCrLf
    Dim handles() As Variant
    handles = Client1.SubscribeMultipleItems(arguments)
    
    OutputText = OutputText & "Waiting for 20 seconds..." & vbCrLf
    Pause 20000

    OutputText = OutputText & "Changing subscription..." & vbCrLf
    Call Client1.ChangeItemSubscription(handles(0), 500)

    OutputText = OutputText & "Waiting for 10 seconds..." & vbCrLf
    Pause 10000

    OutputText = OutputText & "Unsubscribing..." & vbCrLf
    Client1.UnsubscribeAllItems

    OutputText = OutputText & "Waiting for 10 seconds..." & vbCrLf
    Pause 10000

    Set Client1 = Nothing
End Sub

Public Sub Client1_ItemChanged(ByVal sender As Variant, ByVal eventArgs As EasyDAItemChangedEventArgs)
    If Not eventArgs.Succeeded Then
        OutputText = OutputText & "*** Failure: " & eventArgs.ErrorMessageBrief & vbCrLf
        Exit Sub
    End If
    OutputText = OutputText & eventArgs.vtq & vbCrLf
End Sub
Rem #endregion Example ChangeItemSubscription.MainXml

Rem #region Example GetMultiplePropertyValues.DataTypeXml
Rem This example shows how to obtain a data type of all OPC XML-DA items under a branch.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Private Sub GetMultiplePropertyValues_DataTypeXml_Command_Click()
    OutputText = ""
    
    Dim serverDescriptor As New serverDescriptor
    serverDescriptor.UrlString = "http://opcxml.demo-this.com/XmlDaSampleServer/Service.asmx"

    Dim NodeDescriptor As New DANodeDescriptor
    NodeDescriptor.ItemId = "Static/Analog Types"
    
    Dim BrowseParameters As New DABrowseParameters
    BrowseParameters.BrowseFilter = DABrowseFilter_Leaves
    
    ' Instantiate the client object
    Dim client As New EasyDAClient

    ' Browse for all leaves under the "Static/Analog Types" branch
    Dim NodeElementCollection As DANodeElementCollection
    Set NodeElementCollection = client.BrowseNodes(serverDescriptor, NodeDescriptor, BrowseParameters)
    
    ' Create list of node descriptors, one for each leaf obtained
    Dim arguments()
    ReDim arguments(NodeElementCollection.Count)
    Dim i: i = 0
    Dim NodeElement: For Each NodeElement In NodeElementCollection
        ' filter out hint leafs that do not represent real OPC XML-DA items (rare)
        If Not NodeElement.IsHint Then
            Dim PropertyArguments As New DAPropertyArguments
            Set PropertyArguments.serverDescriptor = serverDescriptor
            Set PropertyArguments.NodeDescriptor = NodeElement.ToDANodeDescriptor
            PropertyArguments.PropertyDescriptor.PropertyId.NumericalValue = DAPropertyIds_DataType
            Set arguments(i) = PropertyArguments
            Set PropertyArguments = Nothing
            i = i + 1
        End If
    Next

    Dim propertyArgumentArray()
    ReDim propertyArgumentArray(i - 1)
    Dim j: For j = 0 To i - 1
        Set propertyArgumentArray(j) = arguments(j)
    Next

    ' Get the value of DataType property; it is a 16-bit signed integer
    Dim valueResultArray() As Variant
    valueResultArray = client.GetMultiplePropertyValues(propertyArgumentArray)

    ' Display results
    For j = 0 To i - 1
        Dim NodeDescriptor2 As DANodeDescriptor
        Set NodeDescriptor2 = propertyArgumentArray(j).NodeDescriptor
                
        Dim ValueResult As ValueResult: Set ValueResult = valueResultArray(j)
        ' Check if there has been an error getting the property value
        If Not ValueResult.Exception Is Nothing Then
            OutputText = OutputText & NodeDescriptor2.NodeId & " *** Failure: " & ValueResult.Exception.Message & vbCrLf
        Else
          ' Display the obtained data type
          Dim VarType As New VarType
          VarType.InternalValue = ValueResult.value
          OutputText = OutputText & NodeDescriptor2.NodeId & ": " & VarType & vbCrLf
        End If
    Next
End Sub
Rem #endregion Example GetMultiplePropertyValues.DataTypeXml

Rem #region Example PullItemChanged.MainXml
Rem This example shows how to subscribe to OPC XML-DA item changes and obtain the events by pulling them.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Private Sub PullItemChanged_MainXml_Command_Click()
    OutputText = ""
    
    Dim eventArgs As EasyDAItemChangedEventArgs
    
    Dim ItemSubscriptionArguments1 As New EasyDAItemSubscriptionArguments
    ItemSubscriptionArguments1.serverDescriptor.UrlString = "http://opcxml.demo-this.com/XmlDaSampleServer/Service.asmx"
    ItemSubscriptionArguments1.ItemDescriptor.ItemId = "Dynamic/Analog Types/Int"
    ItemSubscriptionArguments1.GroupParameters.RequestedUpdateRate = 1000
    
    Dim arguments(0) As Variant
    Set arguments(0) = ItemSubscriptionArguments1


    ' Instantiate the client object
    Dim client As New EasyDAClient
    
    ' In order to use event pull, you must set a non-zero queue capacity upfront.
    client.PullItemChangedQueueCapacity = 1000
    
    OutputText = OutputText & "Subscribing item changes..." & vbCrLf
    Call client.SubscribeMultipleItems(arguments)
    
    OutputText = OutputText & "Processing item changes for 1 minute..." & vbCrLf
    Dim endTick As Long
    endTick = GetTickCount + 60000
    While GetTickCount < endTick
        Set eventArgs = client.PullItemChanged(2 * 1000)
        If Not eventArgs Is Nothing Then
            ' Handle the notification event
            OutputText = OutputText & eventArgs & vbCrLf
        End If
    Wend
    
    OutputText = OutputText & "Unsubscribing item changes..." & vbCrLf
    client.UnsubscribeAllItems

    OutputText = OutputText & "Finished." & vbCrLf
End Sub
Rem #endregion Example PullItemChanged.MainXml

Rem #region Example ReadMultipleItems.MainXml
Rem This example shows how to read 4 items from an OPC XML-DA server at once, and display their values, timestamps
Rem and qualities.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Private Sub ReadMultipleItems_MainXml_Command_Click()
    OutputText = ""
    
    Dim readArguments1 As New DAReadItemArguments
    readArguments1.serverDescriptor.UrlString = "http://opcxml.demo-this.com/XmlDaSampleServer/Service.asmx"
    readArguments1.ItemDescriptor.ItemId = "Dynamic/Analog Types/Double"
    
    Dim readArguments2 As New DAReadItemArguments
    readArguments2.serverDescriptor.UrlString = "http://opcxml.demo-this.com/XmlDaSampleServer/Service.asmx"
    readArguments2.ItemDescriptor.ItemId = "Dynamic/Analog Types/Double[]"
    
    Dim readArguments3 As New DAReadItemArguments
    readArguments3.serverDescriptor.UrlString = "http://opcxml.demo-this.com/XmlDaSampleServer/Service.asmx"
    readArguments3.ItemDescriptor.ItemId = "Dynamic/Analog Types/Int"
    
    Dim readArguments4 As New DAReadItemArguments
    readArguments4.serverDescriptor.UrlString = "http://opcxml.demo-this.com/XmlDaSampleServer/Service.asmx"
    readArguments4.ItemDescriptor.ItemId = "SomeUnknownItem"
    
    Dim arguments(3) As Variant
    Set arguments(0) = readArguments1
    Set arguments(1) = readArguments2
    Set arguments(2) = readArguments3
    Set arguments(3) = readArguments4

    ' Instantiate the client object
    Dim client As New EasyDAClient

    Dim vtqResults() As Variant
    vtqResults = client.ReadMultipleItems(arguments)

    ' Display results
    Dim i: For i = LBound(vtqResults) To UBound(vtqResults)
        Dim vtqResult As DAVtqResult: Set vtqResult = vtqResults(i)
        If Not (vtqResult.Exception Is Nothing) Then
            OutputText = OutputText & "results(" & i & ") *** Failure: " & vtqResult.ErrorMessageBrief & vbCrLf
        Else
        OutputText = OutputText & "results(" & i & ").Vtq.ToString(): " & vtqResult.vtq & vbCrLf
        End If
    Next
End Sub
Rem #endregion Example ReadMultipleItems.MainXml

Rem #region Example SubscribeItem.MainXml
Rem This example shows how to subscribe to changes of a single OPC XML-DA item and display the value of the item with each change.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

' The client object, with events
'Public WithEvents Client2 As EasyDAClient

Private Sub SubscribeItem_MainXml_Command_Click()
    OutputText = ""
    
    Dim ItemSubscriptionArguments1 As New EasyDAItemSubscriptionArguments
    ItemSubscriptionArguments1.serverDescriptor.UrlString = "http://opcxml.demo-this.com/XmlDaSampleServer/Service.asmx"
    ItemSubscriptionArguments1.ItemDescriptor.ItemId = "Dynamic/Analog Types/Int"
    ItemSubscriptionArguments1.GroupParameters.RequestedUpdateRate = 1000

    Dim arguments(0) As Variant
    Set arguments(0) = ItemSubscriptionArguments1

    ' Instantiate the client object and hook events
    Set Client2 = New EasyDAClient

    OutputText = OutputText & "Subscribing..." & vbCrLf
    Dim handleArray() As Variant
    handleArray = Client2.SubscribeMultipleItems(arguments)
    
    OutputText = OutputText & "Processing item changed events for 1 minute..." & vbCrLf
    Pause 60000

    OutputText = OutputText & "Unsubscribing..." & vbCrLf
    Client2.UnsubscribeAllItems

    OutputText = OutputText & "Waiting for 5 seconds..." & vbCrLf
    Pause 5000

    OutputText = OutputText & "Finished." & vbCrLf
    Set Client2 = Nothing
End Sub

Public Sub Client2_ItemChanged(ByVal sender As Variant, ByVal eventArgs As EasyDAItemChangedEventArgs)
    If eventArgs.Succeeded Then
        OutputText = OutputText & eventArgs.vtq & vbCrLf
    Else
        OutputText = OutputText & "*** Failure: " & eventArgs.ErrorMessageBrief & vbCrLf
    End If
End Sub
Rem #endregion Example SubscribeItem.MainXml

Rem #region Example WriteItemValue.MainXml
Rem This example shows how to write a value into a single OPC XML-DA item.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Private Sub WriteItemValue_MainXml_Command_Click()
    OutputText = ""
    
    Dim itemValueArguments1 As New DAItemValueArguments
    itemValueArguments1.serverDescriptor.UrlString = "http://opcxml.demo-this.com/XmlDaSampleServer/Service.asmx"
    itemValueArguments1.ItemDescriptor.ItemId = "Static/Analog Types/Int"
    itemValueArguments1.SetValue 12345
    
    Dim arguments(0) As Variant
    Set arguments(0) = itemValueArguments1

    ' Instantiate the client object
    Dim client As New EasyDAClient

    ' Modify values of nodes
    Dim results() As Variant
    results = client.WriteMultipleItemValues(arguments)

    Dim operationResult As operationResult: Set operationResult = results(0)
    If Not operationResult.Succeeded Then
        OutputText = OutputText & "*** Failure: " & operationResult.ErrorMessageBrief & vbCrLf
    End If
End Sub
Rem #endregion Example WriteItemValue.MainXml


