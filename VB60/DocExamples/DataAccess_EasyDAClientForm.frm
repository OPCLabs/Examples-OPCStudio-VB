VERSION 5.00
Begin VB.Form DataAccess_EasyDAClientForm 
   Caption         =   "DataAccess._EasyDAClient"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton WriteMultipleItemValues_Main_Command 
      Caption         =   "WriteMultipleItemValues.Main"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   6840
      Width           =   3015
   End
   Begin VB.CommandButton WriteMultipleItems_Main_Command 
      Caption         =   "WriteMultipleItems.Main"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   6360
      Width           =   3015
   End
   Begin VB.CommandButton WriteItemValue_Main_Command 
      Caption         =   "WriteItemValue.Main"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   5880
      Width           =   3015
   End
   Begin VB.CommandButton WriteItem_Main_Command 
      Caption         =   "WriteItem.Main"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   5400
      Width           =   3015
   End
   Begin VB.CommandButton SubscribeMultipleItems_Main_Command 
      Caption         =   "SubscribeMultipleItems.Main"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   4920
      Width           =   3015
   End
   Begin VB.CommandButton SubscribeItem_Main_Command 
      Caption         =   "SubscribeItem.Main"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   4440
      Width           =   3015
   End
   Begin VB.CommandButton ReadMultipleItemValues_Main_Command 
      Caption         =   "ReadMultipleItemValues.Main"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3960
      Width           =   3015
   End
   Begin VB.CommandButton ReadMultipleItems_Main_Command 
      Caption         =   "ReadMultipleItems.Main"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   3480
      Width           =   3015
   End
   Begin VB.CommandButton ReadItemValue_Main_Command 
      Caption         =   "ReadItemValue.Main"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   3015
   End
   Begin VB.CommandButton ReadItem_Main_Command 
      Caption         =   "ReadItem.Main"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   3015
   End
   Begin VB.CommandButton PullItemChanged_MultipleItems_Command 
      Caption         =   "PullItemChanged.MultipleItems"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   3015
   End
   Begin VB.CommandButton PullItemChanged_Main_Command 
      Caption         =   "PullItemChanged.Main"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   3015
   End
   Begin VB.CommandButton GetMultiplePropertyValues_Main_Command 
      Caption         =   "GetMultiplePropertyValues.Main"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   3015
   End
   Begin VB.CommandButton BrowseProperties_Main_Command 
      Caption         =   "BrowseProperties.Main"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   3015
   End
   Begin VB.CommandButton BrowseNodes_Main_Command 
      Caption         =   "BrowseNodes.Main"
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
Attribute VB_Name = "DataAccess_EasyDAClientForm"
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

' Pause
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Public Sub Pause(Optional milliseconds As Long)
    On Error Resume Next
    Dim endTickCount As Long
    endTickCount = GetTickCount + milliseconds
    While GetTickCount < endTickCount: Sleep 1: DoEvents: Wend
End Sub

REM #region Example BrowseNodes.Main
REM This example shows how to obtain all nodes under the "Simulation" branch of the address space. For each node, it displays
REM whether the node is a branch or a leaf.
REM
REM Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Private Sub BrowseNodes_Main_Command_Click()
    OutputText = ""
    
    Dim serverDescriptor As New serverDescriptor
    serverDescriptor.ServerClass = "OPCLabs.KitServer.2"
    
    Dim nodeDescriptor As New DANodeDescriptor
    nodeDescriptor.itemId = "Simulation"
    
    Dim browseParameters As New DABrowseParameters
    
    ' Instantiate the client object
    Dim client As New EasyDAClient

    On Error Resume Next
    Dim nodeElements As DANodeElementCollection
    Set nodeElements = client.BrowseNodes(serverDescriptor, nodeDescriptor, browseParameters)
    If Err.Number <> 0 Then
        OutputText = OutputText & "*** Failure: " & Err.Source & ": " & Err.Description & vbCrLf
        Exit Sub
    End If
    On Error GoTo 0

    Dim nodeElement: For Each nodeElement In nodeElements
        OutputText = OutputText & "BrowseElements(""" & nodeElement.Name & """): " & nodeElement.itemId & vbCrLf
        OutputText = OutputText & "    .IsBranch: " & nodeElement.IsBranch & vbCrLf
        OutputText = OutputText & "    .IsLeaf: " & nodeElement.IsLeaf & vbCrLf
    Next
    
End Sub
REM #endregion Example BrowseNodes.Main

REM #region Example BrowseProperties.Main
REM This example shows how to enumerate all properties of an OPC item. For each property, it displays its Id and description.
REM
REM Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Private Sub BrowseProperties_Main_Command_Click()
    OutputText = ""
    
    Dim serverDescriptor As New serverDescriptor
    serverDescriptor.ServerClass = "OPCLabs.KitServer.2"
    
    Dim nodeDescriptor As New DANodeDescriptor
    nodeDescriptor.itemId = "Simulation.Random"
    
    ' Instantiate the client object
    Dim client As New EasyDAClient

    On Error Resume Next
    Dim propertyElements As DAPropertyElementCollection
    Set propertyElements = client.BrowseProperties(serverDescriptor, nodeDescriptor)
    If Err.Number <> 0 Then
        OutputText = OutputText & "*** Failure: " & Err.Source & ": " & Err.Description & vbCrLf
        Exit Sub
    End If
    On Error GoTo 0

    Dim propertyElement: For Each propertyElement In propertyElements
        OutputText = OutputText & "propertyElements(""" & propertyElement.propertyId.NumericalValue & """).Description: " & propertyElement.Description & vbCrLf
    Next
End Sub
REM #endregion Example BrowseProperties.Main

REM #region Example GetMultiplePropertyValues.Main
REM This example shows how to get value of multiple OPC properties.
REM
REM Note that some properties may not have a useful value initially (e.g. until the item is activated in a group), which also the
REM case with Timestamp property as implemented by the demo server. This behavior is server-dependent, and normal. You can run
REM IEasyDAClient.ReadMultipleItemValues.Main.vbs shortly before this example, in order to obtain better property values. Your
REM code may also subscribe to the items in order to assure that they remain active.
REM
REM Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Private Sub GetMultiplePropertyValues_Main_Command_Click()
    OutputText = ""
    
    ' Get the values of Timestamp and AccessRights properties of two items.

    Dim propertyArguments1 As New DAPropertyArguments
    propertyArguments1.serverDescriptor.ServerClass = "OPCLabs.KitServer.2"
    propertyArguments1.nodeDescriptor.itemId = "Simulation.Random"
    propertyArguments1.PropertyDescriptor.propertyId.NumericalValue = DAPropertyIds_Timestamp
    
    Dim propertyArguments2 As New DAPropertyArguments
    propertyArguments2.serverDescriptor.ServerClass = "OPCLabs.KitServer.2"
    propertyArguments2.nodeDescriptor.itemId = "Simulation.Random"
    propertyArguments2.PropertyDescriptor.propertyId.NumericalValue = DAPropertyIds_AccessRights
    
    Dim propertyArguments3 As New DAPropertyArguments
    propertyArguments3.serverDescriptor.ServerClass = "OPCLabs.KitServer.2"
    propertyArguments3.nodeDescriptor.itemId = "Trends.Ramp (1 min)"
    propertyArguments3.PropertyDescriptor.propertyId.NumericalValue = DAPropertyIds_Timestamp
    
    Dim propertyArguments4 As New DAPropertyArguments
    propertyArguments4.serverDescriptor.ServerClass = "OPCLabs.KitServer.2"
    propertyArguments4.nodeDescriptor.itemId = "Trends.Ramp (1 min)"
    propertyArguments4.PropertyDescriptor.propertyId.NumericalValue = DAPropertyIds_AccessRights
    
    Dim arguments(3) As Variant
    Set arguments(0) = propertyArguments1
    Set arguments(1) = propertyArguments2
    Set arguments(2) = propertyArguments3
    Set arguments(3) = propertyArguments4

    ' Instantiate the client object
    Dim client As New EasyDAClient

    ' Obtain values. By default, the Value attributes of the nodes will be read.
    Dim results() As Variant
    results = client.GetMultiplePropertyValues(arguments)

    ' Display results
    Dim i: For i = LBound(results) To UBound(results)
        Dim valueResult As valueResult: Set valueResult = results(i)
        ' Check if there has been an error getting the property value
        If Not valueResult.Exception Is Nothing Then
            OutputText = OutputText & arguments(i).nodeDescriptor.NodeId & " *** Failure: " & valueResult.Exception.Message & vbCrLf
        Else
            OutputText = OutputText & "results(" & i & ").Value: " & valueResult.value & vbCrLf
        End If
    Next
End Sub
REM #endregion Example GetMultiplePropertyValues.Main

REM #region Example PullItemChanged.Main
REM This example shows how to subscribe to item changes and obtain the events by pulling them.
REM
REM Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Private Sub PullItemChanged_Main_Command_Click()
    OutputText = ""
    
    Dim eventArgs As EasyDAItemChangedEventArgs
    
    ' Instantiate the client object
    Dim client As New EasyDAClient
    
    ' In order to use event pull, you must set a non-zero queue capacity upfront.
    client.PullItemChangedQueueCapacity = 1000
    
    OutputText = OutputText & "Subscribing item changes..." & vbCrLf
    Call client.SubscribeItem("", "OPCLabs.KitServer.2", "Simulation.Random", 1000)
    
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
REM #endregion Example PullItemChanged.Main

REM #region Example PullItemChanged.MultipleItems
REM This example shows how to subscribe to changes of multiple items and obtain the item changed events by pulling them.
REM
REM Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Private Sub PullItemChanged_MultipleItems_Command_Click()
    OutputText = ""
    
    Dim eventArgs As EasyDAItemChangedEventArgs
    
    Dim itemSubscriptionArguments1 As New EasyDAItemSubscriptionArguments
    itemSubscriptionArguments1.serverDescriptor.ServerClass = "OPCLabs.KitServer.2"
    itemSubscriptionArguments1.ItemDescriptor.itemId = "Simulation.Random"
    itemSubscriptionArguments1.GroupParameters.requestedUpdateRate = 1000

    Dim itemSubscriptionArguments2 As New EasyDAItemSubscriptionArguments
    itemSubscriptionArguments2.serverDescriptor.ServerClass = "OPCLabs.KitServer.2"
    itemSubscriptionArguments2.ItemDescriptor.itemId = "Trends.Ramp (1 min)"
    itemSubscriptionArguments2.GroupParameters.requestedUpdateRate = 1000

    Dim itemSubscriptionArguments3 As New EasyDAItemSubscriptionArguments
    itemSubscriptionArguments3.serverDescriptor.ServerClass = "OPCLabs.KitServer.2"
    itemSubscriptionArguments3.ItemDescriptor.itemId = "Trends.Sine (1 min)"
    itemSubscriptionArguments3.GroupParameters.requestedUpdateRate = 1000

    ' Intentionally specifying an unknown item here, to demonstrate its behavior.
    Dim itemSubscriptionArguments4 As New EasyDAItemSubscriptionArguments
    itemSubscriptionArguments4.serverDescriptor.ServerClass = "OPCLabs.KitServer.2"
    itemSubscriptionArguments4.ItemDescriptor.itemId = "SomeUnknownItem"
    itemSubscriptionArguments4.GroupParameters.requestedUpdateRate = 1000

    Dim arguments(3) As Variant
    Set arguments(0) = itemSubscriptionArguments1
    Set arguments(1) = itemSubscriptionArguments2
    Set arguments(2) = itemSubscriptionArguments3
    Set arguments(3) = itemSubscriptionArguments4

    ' Instantiate the client object
    Dim client As New EasyDAClient
    ' In order to use event pull, you must set a non-zero queue capacity upfront.
    client.PullItemChangedQueueCapacity = 1000

    OutputText = OutputText & "Subscribing item changes..." & vbCrLf
    Dim handleArray() As Variant
    handleArray = client.SubscribeMultipleItems(arguments)
    
    OutputText = OutputText & "Processing item changes for 1 minute..." & vbCrLf
    Dim endTick As Long
    endTick = GetTickCount + 60000
    While GetTickCount < endTick
        Set eventArgs = client.PullItemChanged(2 * 1000)
        If Not eventArgs Is Nothing Then
            ' Handle the notification event
            If eventArgs.Succeeded Then
                OutputText = OutputText & eventArgs.arguments.ItemDescriptor.itemId & ": " & eventArgs.vtq & vbCrLf
            Else
                OutputText = OutputText & eventArgs.arguments.ItemDescriptor.itemId & " *** Failure: " & eventArgs.ErrorMessageBrief & vbCrLf
            End If
        End If
    Wend
    
    OutputText = OutputText & "Unsubscribing item changes..." & vbCrLf
    client.UnsubscribeAllItems

    OutputText = OutputText & "Finished." & vbCrLf
End Sub
REM #endregion Example PullItemChanged.MultipleItems

REM #region Example ReadItem.Main
REM This example shows how to read a single item, and display its value, timestamp and quality.
REM
REM Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Private Sub ReadItem_Main_Command_Click()
    OutputText = ""
    
    ' Instantiate the client object
    Dim client As New EasyDAClient

    On Error Resume Next
    Dim vtq As DAVtq
    Set vtq = client.ReadItem("", "OPCLabs.KitServer.2", "Simulation.Random")
    If Err.Number <> 0 Then
        OutputText = OutputText & "*** Failure: " & Err.Source & ": " & Err.Description & vbCrLf
        Exit Sub
    End If
    On Error GoTo 0

    ' Display results
    OutputText = OutputText & "Vtq: " & vtq & vbCrLf
End Sub
REM #endregion Example ReadItem.Main

REM #region Example ReadItemValue.Main
REM This example shows how to read value of a single item, and display it.
REM
REM Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Private Sub ReadItemValue_Main_Command_Click()
    OutputText = ""
    
    ' Instantiate the client object
    Dim client As New EasyDAClient

    On Error Resume Next
    Dim value
    value = client.ReadItemValue("", "OPCLabs.KitServer.2", "Simulation.Random")
    If Err.Number <> 0 Then
        OutputText = OutputText & "*** Failure: " & Err.Source & ": " & Err.Description & vbCrLf
        Exit Sub
    End If
    On Error GoTo 0

    ' Display results
    OutputText = OutputText & "Value: " & value & vbCrLf
End Sub
REM #endregion Example ReadItemValue.Main

REM #region Example ReadMultipleItems.Main
REM This example shows how to read 4 items at once, and display their values, timestamps and qualities.
REM
REM Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Private Sub ReadMultipleItems_Main_Command_Click()
    OutputText = ""
    
    Dim readArguments1 As New DAReadItemArguments
    readArguments1.serverDescriptor.ServerClass = "OPCLabs.KitServer.2"
    readArguments1.ItemDescriptor.itemId = "Simulation.Random"
    
    Dim readArguments2 As New DAReadItemArguments
    readArguments2.serverDescriptor.ServerClass = "OPCLabs.KitServer.2"
    readArguments2.ItemDescriptor.itemId = "Trends.Ramp (1 min)"
    
    Dim readArguments3 As New DAReadItemArguments
    readArguments3.serverDescriptor.ServerClass = "OPCLabs.KitServer.2"
    readArguments3.ItemDescriptor.itemId = "Trends.Sine (1 min)"
    
    Dim readArguments4 As New DAReadItemArguments
    readArguments4.serverDescriptor.ServerClass = "OPCLabs.KitServer.2"
    readArguments4.ItemDescriptor.itemId = "Simulation.Register_I4"
    
    Dim arguments(3) As Variant
    Set arguments(0) = readArguments1
    Set arguments(1) = readArguments2
    Set arguments(2) = readArguments3
    Set arguments(3) = readArguments4

    ' Instantiate the client object
    Dim client As New EasyDAClient

    Dim results() As Variant
    results = client.ReadMultipleItems(arguments)

    ' Display results
    Dim i: For i = LBound(results) To UBound(results)
        Dim vtqResult As DAVtqResult: Set vtqResult = results(i)
        If vtqResult.Succeeded Then
            OutputText = OutputText & "results(" & i & ").Vtq: " & vtqResult.vtq & vbCrLf
        Else
            OutputText = OutputText & "results(" & i & ") *** Failure: " & vtqResult.ErrorMessageBrief & vbCrLf
        End If
    Next
End Sub
REM #endregion Example ReadMultipleItems.Main

REM #region Example ReadMultipleItemValues.Main
REM This example shows how to read values of 4 items at once, and display them.
REM
REM Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Private Sub ReadMultipleItemValues_Main_Command_Click()
    OutputText = ""
    
    Dim readArguments1 As New DAReadItemArguments
    readArguments1.serverDescriptor.ServerClass = "OPCLabs.KitServer.2"
    readArguments1.ItemDescriptor.itemId = "Simulation.Random"
    
    Dim readArguments2 As New DAReadItemArguments
    readArguments2.serverDescriptor.ServerClass = "OPCLabs.KitServer.2"
    readArguments2.ItemDescriptor.itemId = "Trends.Ramp (1 min)"
    
    Dim readArguments3 As New DAReadItemArguments
    readArguments3.serverDescriptor.ServerClass = "OPCLabs.KitServer.2"
    readArguments3.ItemDescriptor.itemId = "Trends.Sine (1 min)"
    
    Dim readArguments4 As New DAReadItemArguments
    readArguments4.serverDescriptor.ServerClass = "OPCLabs.KitServer.2"
    readArguments4.ItemDescriptor.itemId = "Simulation.Register_I4"
    
    Dim arguments(3) As Variant
    Set arguments(0) = readArguments1
    Set arguments(1) = readArguments2
    Set arguments(2) = readArguments3
    Set arguments(3) = readArguments4

    ' Instantiate the client object
    Dim client As New EasyDAClient

    Dim results() As Variant
    results = client.ReadMultipleItemValues(arguments)

    ' Display results
    Dim i: For i = LBound(results) To UBound(results)
        Dim valueResult As valueResult: Set valueResult = results(i)
        If valueResult.Succeeded Then
            OutputText = OutputText & "results(" & i & ").Value: " & valueResult.value & vbCrLf
        Else
            OutputText = OutputText & "results(" & i & ") *** Failure: " & valueResult.ErrorMessageBrief & vbCrLf
        End If
    Next
End Sub
REM #endregion Example ReadMultipleItemValues.Main

REM #region Example SubscribeItem.Main
REM This example shows how to subscribe to changes of a single item and display the value of the item with each change.
REM
REM Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

' The client object, with events
'Public WithEvents Client1 As EasyDAClient

Private Sub SubscribeItem_Main_Command_Click()
    OutputText = ""
    
    ' Instantiate the client object and hook events
    Set Client1 = New EasyDAClient

    OutputText = OutputText & "Subscribing..." & vbCrLf
    Call Client1.SubscribeItem("", "OPCLabs.KitServer.2", "Simulation.Random", 1000)

    OutputText = OutputText & "Processing item changed events for 1 minute..." & vbCrLf
    Pause 60000

    OutputText = OutputText & "Unsubscribing..." & vbCrLf
    Client1.UnsubscribeAllItems

    OutputText = OutputText & "Waiting for 5 seconds..." & vbCrLf
    Pause 5000

    OutputText = OutputText & "Finished." & vbCrLf
    Set Client1 = Nothing
End Sub

Public Sub Client1_ItemChanged(ByVal sender As Variant, ByVal eventArgs As EasyDAItemChangedEventArgs)
    If eventArgs.Succeeded Then
        OutputText = OutputText & eventArgs.vtq & vbCrLf
    Else
        OutputText = OutputText & "*** Failure: " & eventArgs.ErrorMessageBrief & vbCrLf
    End If
End Sub
REM #endregion Example SubscribeItem.Main

REM #region Example SubscribeMultipleItems.Main
REM This example shows how to subscribe to changes of multiple items and display the value of the item with each change.
REM
REM Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

' The client object, with events
'Public WithEvents Client2 As EasyDAClient

Private Sub SubscribeMultipleItems_Main_Command_Click()
    OutputText = ""
    
    Dim itemSubscriptionArguments1 As New EasyDAItemSubscriptionArguments
    itemSubscriptionArguments1.serverDescriptor.ServerClass = "OPCLabs.KitServer.2"
    itemSubscriptionArguments1.ItemDescriptor.itemId = "Simulation.Random"
    itemSubscriptionArguments1.GroupParameters.requestedUpdateRate = 1000

    Dim itemSubscriptionArguments2 As New EasyDAItemSubscriptionArguments
    itemSubscriptionArguments2.serverDescriptor.ServerClass = "OPCLabs.KitServer.2"
    itemSubscriptionArguments2.ItemDescriptor.itemId = "Trends.Ramp (1 min)"
    itemSubscriptionArguments2.GroupParameters.requestedUpdateRate = 1000

    Dim itemSubscriptionArguments3 As New EasyDAItemSubscriptionArguments
    itemSubscriptionArguments3.serverDescriptor.ServerClass = "OPCLabs.KitServer.2"
    itemSubscriptionArguments3.ItemDescriptor.itemId = "Trends.Sine (1 min)"
    itemSubscriptionArguments3.GroupParameters.requestedUpdateRate = 1000

    ' Intentionally specifying an unknown item here, to demonstrate its behavior.
    Dim itemSubscriptionArguments4 As New EasyDAItemSubscriptionArguments
    itemSubscriptionArguments4.serverDescriptor.ServerClass = "OPCLabs.KitServer.2"
    itemSubscriptionArguments4.ItemDescriptor.itemId = "Simulation.Register_I4"
    itemSubscriptionArguments4.GroupParameters.requestedUpdateRate = 1000

    Dim arguments(3) As Variant
    Set arguments(0) = itemSubscriptionArguments1
    Set arguments(1) = itemSubscriptionArguments2
    Set arguments(2) = itemSubscriptionArguments3
    Set arguments(3) = itemSubscriptionArguments4

    ' Instantiate the client object
    Set Client2 = New EasyDAClient

    OutputText = OutputText & "Subscribing item changes..." & vbCrLf
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
        OutputText = OutputText & eventArgs.arguments.ItemDescriptor.itemId & ": " & eventArgs.vtq & vbCrLf
    Else
        OutputText = OutputText & eventArgs.arguments.ItemDescriptor.itemId & " *** Failure: " & eventArgs.ErrorMessageBrief & vbCrLf
    End If
End Sub
REM #endregion Example SubscribeMultipleItems.Main

REM #region Example WriteItem.Main
REM This example shows how to write a value, timestamp and quality into a single item.
REM
REM Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Private Sub WriteItem_Main_Command_Click()
    OutputText = ""
    
    ' Instantiate the client object
    Dim client As New EasyDAClient

    On Error Resume Next
    Call client.WriteItem("", "OPCLabs.KitServer.2", "Simulation.Register_I4", 12345, DateSerial(1980, 1, 1), DAQualities_GoodNonspecific)
    If Err.Number <> 0 Then
        OutputText = OutputText & "*** Failure: " & Err.Source & ": " & Err.Description & vbCrLf
        Exit Sub
    End If
    On Error GoTo 0
End Sub
REM #endregion Example WriteItem.Main

REM #region Example WriteItemValue.Main
REM This example shows how to write a value into a single item.
REM
REM Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Private Sub WriteItemValue_Main_Command_Click()
    OutputText = ""
    
    ' Instantiate the client object
    Dim client As New EasyDAClient

    On Error Resume Next
    Call client.WriteItemValue("", "OPCLabs.KitServer.2", "Simulation.Register_I4", 12345)
    If Err.Number <> 0 Then
        OutputText = OutputText & "*** Failure: " & Err.Source & ": " & Err.Description & vbCrLf
        Exit Sub
    End If
    On Error GoTo 0
End Sub
REM #endregion Example WriteItemValue.Main

REM #region Example WriteMultipleItems.Main
REM This example shows how to write values, timestamps and qualities into 3 items at once.
REM
REM Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Private Sub WriteMultipleItems_Main_Command_Click()
    OutputText = ""
    
    Dim itemVtqArguments1 As New DAItemVtqArguments
    itemVtqArguments1.serverDescriptor.ServerClass = "OPCLabs.KitServer.2"
    itemVtqArguments1.ItemDescriptor.itemId = "Simulation.Register_I4"
    itemVtqArguments1.vtq.SetValue 23456
    itemVtqArguments1.vtq.TimestampLocal = Now()
    itemVtqArguments1.vtq.Quality.NumericalValue = DAQualities_GoodNonspecific
    
    Dim itemVtqArguments2 As New DAItemVtqArguments
    itemVtqArguments2.serverDescriptor.ServerClass = "OPCLabs.KitServer.2"
    itemVtqArguments2.ItemDescriptor.itemId = "Simulation.Register_R8"
    itemVtqArguments2.vtq.SetValue 2.3456789
    itemVtqArguments2.vtq.TimestampLocal = Now()
    itemVtqArguments2.vtq.Quality.NumericalValue = DAQualities_GoodNonspecific
    
    Dim itemVtqArguments3 As New DAItemVtqArguments
    itemVtqArguments3.serverDescriptor.ServerClass = "OPCLabs.KitServer.2"
    itemVtqArguments3.ItemDescriptor.itemId = "Simulation.Register_BSTR"
    itemVtqArguments3.vtq.SetValue "ABC"
    itemVtqArguments3.vtq.TimestampLocal = Now()
    itemVtqArguments3.vtq.Quality.NumericalValue = DAQualities_GoodNonspecific
    
    Dim arguments(2) As Variant
    Set arguments(0) = itemVtqArguments1
    Set arguments(1) = itemVtqArguments2
    Set arguments(2) = itemVtqArguments3

    ' Instantiate the client object
    Dim client As New EasyDAClient

    ' Modify values of nodes
    Dim results() As Variant
    results = client.WriteMultipleItems(arguments)

    ' Display results
    Dim i: For i = LBound(results) To UBound(results)
        Dim operationResult As operationResult: Set operationResult = results(i)
        If operationResult.Succeeded Then
            OutputText = OutputText & "result " & i & " success" & vbCrLf
        Else
            OutputText = OutputText & "result " & i & " *** Failure: " & operationResult.ErrorMessageBrief & vbCrLf
        End If
    Next
End Sub
REM #endregion Example WriteMultipleItems.Main

REM #region Example WriteMultipleItemValues.Main
REM This example shows how to write values into 3 items at once.
REM
REM Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Private Sub WriteMultipleItemValues_Main_Command_Click()
    OutputText = ""
    
    Dim itemValueArguments1 As New DAItemValueArguments
    itemValueArguments1.serverDescriptor.ServerClass = "OPCLabs.KitServer.2"
    itemValueArguments1.ItemDescriptor.itemId = "Simulation.Register_I4"
    itemValueArguments1.SetValue 23456
    
    Dim itemValueArguments2 As New DAItemValueArguments
    itemValueArguments2.serverDescriptor.ServerClass = "OPCLabs.KitServer.2"
    itemValueArguments2.ItemDescriptor.itemId = "Simulation.Register_R8"
    itemValueArguments2.SetValue 2.3456789
    
    Dim itemValueArguments3 As New DAItemValueArguments
    itemValueArguments3.serverDescriptor.ServerClass = "OPCLabs.KitServer.2"
    itemValueArguments3.ItemDescriptor.itemId = "Simulation.Register_BSTR"
    itemValueArguments3.SetValue "ABC"
    
    Dim arguments(2) As Variant
    Set arguments(0) = itemValueArguments1
    Set arguments(1) = itemValueArguments2
    Set arguments(2) = itemValueArguments3

    ' Instantiate the client object
    Dim client As New EasyDAClient

    ' Modify values of nodes
    Dim results() As Variant
    results = client.WriteMultipleItemValues(arguments)

    ' Display results
    Dim i: For i = LBound(results) To UBound(results)
        Dim operationResult As operationResult: Set operationResult = results(i)
        If operationResult.Succeeded Then
            OutputText = OutputText & "result " & i & " success" & vbCrLf
        Else
            OutputText = OutputText & "result " & i & " *** Failure: " & operationResult.ErrorMessageBrief & vbCrLf
        End If
    Next
End Sub
REM #endregion Example WriteMultipleItemValues.Main

