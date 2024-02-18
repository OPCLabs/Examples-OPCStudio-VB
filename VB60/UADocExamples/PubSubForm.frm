VERSION 5.00
Begin VB.Form PubSubForm 
   Caption         =   "PubSub"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton EasyUASubscriber_UnsubscribeDataSet_Main1_Command 
      Caption         =   "EasyUASubscriber_UnsubscribeDataSet.Main1"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   4440
      Width           =   3975
   End
   Begin VB.CommandButton EasyUASubscriber_SubscribeDataSet_Secure_Command 
      Caption         =   "_EasyUASubscriber.SubscribeDataSet.Secure"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3960
      Width           =   3975
   End
   Begin VB.CommandButton EasyUASubscriber_SubscribeDataSet_PublisherId_Command 
      Caption         =   "_EasyUASubscriber.SubscribeDataSet.PublisherId"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3480
      Width           =   3975
   End
   Begin VB.CommandButton EasyUASubscriber_SubscribeDataSet_MqttJsonTcp_Command 
      Caption         =   "_EasyUASubscriber.SubscribeDataSet.MqttJsonTcp"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   3975
   End
   Begin VB.CommandButton EasyUASubscriber_SubscribeDataSet_Metadata_Command 
      Caption         =   "_EasyUASubscriber.SubscribeDataSet.Metadata"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   3975
   End
   Begin VB.CommandButton EasyUASubscriber_SubscribeDataSet_Main1_Command 
      Caption         =   "_EasyUASubscriber.SubscribeDataSet.Main1"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   3975
   End
   Begin VB.CommandButton EasyUASubscriber_SubscribeDataSet_Filter_Command 
      Caption         =   "_EasyUASubscriber.SubscribeDataSet.Filter"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   3975
   End
   Begin VB.CommandButton EasyUASubscriber_SubscribeDataSet_FieldNames_Command 
      Caption         =   "_EasyUASubscriber.SubscribeDataSet.FieldNames"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   3975
   End
   Begin VB.CommandButton EasyUASubscriber_SubscribeDataSet_ExtractField_Command 
      Caption         =   "_EasyUASubscriber.SubscribeDataSet.ExtractField"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   3975
   End
   Begin VB.CommandButton EasyUASubscriber_PullDataSetMessage_Main1_Command 
      Caption         =   "_EasyUASubscriber.PullDataSetMessage.Main1"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
   Begin VB.TextBox OutputText 
      Height          =   7575
      Left            =   4200
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "PubSubForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Option Explicit

' The subscriber object, with events
Public WithEvents Subscriber1 As EasyUASubscriber
Attribute Subscriber1.VB_VarHelpID = -1

' The subscriber object, with events
Public WithEvents Subscriber2 As EasyUASubscriber
Attribute Subscriber2.VB_VarHelpID = -1

' The subscriber object, with events
Public WithEvents Subscriber3 As EasyUASubscriber
Attribute Subscriber3.VB_VarHelpID = -1

' The subscriber object, with events
Public WithEvents Subscriber4 As EasyUASubscriber
Attribute Subscriber4.VB_VarHelpID = -1

' The subscriber object, with events
Public WithEvents Subscriber5 As EasyUASubscriber
Attribute Subscriber5.VB_VarHelpID = -1

' The subscriber object, with events
Public WithEvents Subscriber6 As EasyUASubscriber
Attribute Subscriber6.VB_VarHelpID = -1

' The subscriber object, with events.
Public WithEvents Subscriber7 As EasyUASubscriber
Attribute Subscriber7.VB_VarHelpID = -1

' The subscriber object, with events.
Public WithEvents Subscriber8 As EasyUASubscriber
Attribute Subscriber8.VB_VarHelpID = -1

' The subscriber object, with events.
Public WithEvents Subscriber11 As EasyUASubscriber
Attribute Subscriber11.VB_VarHelpID = -1

' Pause
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Public Sub Pause(Optional milliseconds As Long)
    On Error Resume Next
    Dim endTickCount As Long
    endTickCount = GetTickCount + milliseconds
    While GetTickCount < endTickCount: Sleep 1: DoEvents: Wend
End Sub

REM #region Example _EasyUASubscriber.PullDataSetMessage.Main1
REM This example shows how to subscribe to all dataset messages on an OPC-UA PubSub connection, and pull events, and display
REM the incoming datasets.
REM
REM In order to produce network messages for this example, run the UADemoPublisher tool. For documentation, see
REM https://kb.opclabs.com/UADemoPublisher_Basics . In some cases, you may have to specify the interface name to be used.
REM
REM Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Private Sub EasyUASubscriber_PullDataSetMessage_Main1_Command_Click()
    OutputText = ""

    ' Define the PubSub connection we will work with.
    Dim subscribeDataSetArguments :  Set subscribeDataSetArguments = New EasyUASubscribeDataSetArguments
    Dim ConnectionDescriptor :  Set ConnectionDescriptor = subscribeDataSetArguments.dataSetSubscriptionDescriptor.ConnectionDescriptor
    ConnectionDescriptor.ResourceAddress.ResourceDescriptor.UrlString = "opc.udp://239.0.0.1"
    ' In some cases you may have to set the interface (network adapter) name that needs to be used, similarly to
    ' the statement below. Your actual interface name may differ, of course.
    'ConnectionDescriptor.ResourceAddress.InterfaceName = "Ethernet"

    ' Instantiate the subscriber object.
    Dim Subscriber :  Set Subscriber = New EasyUASubscriber
    ' In order to use event pull, you must set a non-zero queue capacity upfront.
    Subscriber.PullDataSetMessageQueueCapacity = 1000

    OutputText = OutputText & "Subscribing..." & vbCrLf
    Subscriber.SubscribeDataSet subscribeDataSetArguments

    OutputText = OutputText & "Processing dataset message events for 20 seconds..." & vbCrLf
    Dim endTime : endTime = Now() + 20 * (1 / 24 / 60 / 60)
    Do
        Dim eventArgs :  Set eventArgs = Subscriber.PullDataSetMessage(2 * 1000)
        If Not (eventArgs Is Nothing) Then
            ' Display the dataset.
            If eventArgs.Succeeded Then
                ' An event with null DataSetData just indicates a successful connection.
                If Not (eventArgs.DataSetData Is Nothing) Then
                    OutputText = OutputText & vbCrLf
                    OutputText = OutputText & "Dataset data: " & eventArgs.DataSetData & vbCrLf
                    Dim Pair : For Each Pair In eventArgs.DataSetData.FieldDataDictionary
                        OutputText = OutputText & Pair & vbCrLf
                    Next
                End If
            Else
                OutputText = OutputText & vbCrLf
                OutputText = OutputText & "*** Failure: " & eventArgs.ErrorMessageBrief & vbCrLf
            End If
        End If
    Loop While Now() < endTime

    OutputText = OutputText & "Unsubscribing..." & vbCrLf
    Subscriber.UnsubscribeAllDataSets

    OutputText = OutputText & "Waiting for 1 second..." & vbCrLf
    ' Unsubscribe operation is asynchronous, messages may still come for a short while.
    Pause 1000

    Set Subscriber = Nothing

    OutputText = OutputText & "Finished." & vbCrLf
End Sub
' Example output:
'
'Subscribing...
'Processing dataset message events for 20 seconds...
'
''Dataset data: Good; Data; publisher="32", writer=1, class=eae79794-1af7-4f96-8401-4096cd1d8908, fields: 4
'[#0, True {System.Boolean} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#1, 7945 {System.Int32} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#2, 5246 {System.Int32} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#3, 9/30/2019 11:19:14 AM {System.DateTime} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'
'Dataset data: Good; Data; publisher="32", writer=3, class=96976b7b-0db7-46c3-a715-0979884b55ae, fields: 100
'[#0, 45 {System.Int64} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#1, 145 {System.Int64} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#2, 245 {System.Int64} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#3, 345 {System.Int64} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#4, 445 {System.Int64} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#5, 545 {System.Int64} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#6, 645 {System.Int64} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#7, 745 {System.Int64} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#8, 845 {System.Int64} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#9, 945 {System.Int64} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#10, 1045 {System.Int64} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'...
REM #endregion Example _EasyUASubscriber.PullDataSetMessage.Main1

REM #region Example _EasyUASubscriber.SubscribeDataSet.ExtractField
REM This example shows how to subscribe to dataset messages extract data of a specific field.
REM
REM In order to produce network messages for this example, run the UADemoPublisher tool. For documentation, see
REM https://kb.opclabs.com/UADemoPublisher_Basics . In some cases, you may have to specify the interface name to be used.
REM
REM Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

' The subscriber object, with events.
'Public WithEvents Subscriber7 As EasyUASubscriber

Private Sub EasyUASubscriber_SubscribeDataSet_ExtractField_Command_Click()
    OutputText = ""

    Const UAPublisherIdType_UInt64 = 4

    ' Define the PubSub connection we will work with.
    Dim subscribeDataSetArguments :  Set subscribeDataSetArguments = CreateObject("OpcLabs.EasyOpc.UA.PubSub.OperationModel.EasyUASubscribeDataSetArguments")
    Dim ConnectionDescriptor :  Set ConnectionDescriptor = subscribeDataSetArguments.dataSetSubscriptionDescriptor.ConnectionDescriptor
    ConnectionDescriptor.ResourceAddress.ResourceDescriptor.UrlString = "opc.udp://239.0.0.1"
    ' In some cases you may have to set the interface (network adapter) name that needs to be used, similarly to
    ' the statement below. Your actual interface name may differ, of course.
    ' ConnectionDescriptor.ResourceAddress.InterfaceName = "Ethernet"

    ' Define the filter. Publisher Id (unsigned 64-bits) is 31, and the dataset writer Id is 1.
    subscribeDataSetArguments.dataSetSubscriptionDescriptor.Filter.PublisherId.SetIdentifier UAPublisherIdType_UInt64, 31
    subscribeDataSetArguments.dataSetSubscriptionDescriptor.Filter.DataSetWriterDescriptor.DataSetWriterId = 1

    ' Define the metadata. For UADP, the order of field metadata must correspond to the order of fields in the dataset message.
    ' If the field names were contained in the dataset message (such as in JSON), or if we knew the metadata from some other
    ' source, this step would not be needed.
    ' Since the encoding is not RawData, we do not have to specify the type information for the fields.
    Dim metaData :  Set metaData = CreateObject("OpcLabs.EasyOpc.UA.PubSub.Configuration.UADataSetMetaData")
    '
    Dim field1 :  Set field1 = CreateObject("OpcLabs.EasyOpc.UA.PubSub.Configuration.UAFieldMetaData")
    field1.Name = "BoolToggle"
    metaData.Add(field1)
    '
    Dim field2 :  Set field2 = CreateObject("OpcLabs.EasyOpc.UA.PubSub.Configuration.UAFieldMetaData")
    field2.Name = "Int32"
    metaData.Add(field2)
    '
    Dim field3 :  Set field3 = CreateObject("OpcLabs.EasyOpc.UA.PubSub.Configuration.UAFieldMetaData")
    field3.Name = "Int32Fast"
    metaData.Add(field3)
    '
    Dim field4 :  Set field4 = CreateObject("OpcLabs.EasyOpc.UA.PubSub.Configuration.UAFieldMetaData")
    field4.Name = "DateTime"
    metaData.Add(field4)
    '
    Set subscribeDataSetArguments.dataSetSubscriptionDescriptor.DataSetMetaData = metaData
    
    ' Instantiate the subscriber object and hook events.
    Set Subscriber7 = New EasyUASubscriber

    OutputText = OutputText & "Subscribing..." & vbCrLf
    Subscriber7.SubscribeDataSet subscribeDataSetArguments

    OutputText = OutputText & "Processing dataset message events for 20 seconds..." & vbCrLf
    Pause 20 * 1000

    OutputText = OutputText & "Unsubscribing..." & vbCrLf
    Subscriber7.UnsubscribeAllDataSets

    OutputText = OutputText & "Waiting for 1 second..." & vbCrLf
    ' Unsubscribe operation is asynchronous, messages may still come for a short while.
    Pause 1000

    Set Subscriber7 = Nothing

    OutputText = OutputText & "Finished." & vbCrLf
End Sub

Private Sub Subscriber7_DataSetMessage(ByVal sender As Variant, ByVal eventArgs As EasyUADataSetMessageEventArgs)
    ' Display the dataset.
    If eventArgs.Succeeded Then
        ' An event with null DataSetData just indicates a successful connection.
        If Not (eventArgs.DataSetData Is Nothing) Then
            ' Extract field data, looking up the field by its name.
            Dim Int32FastValueData :  Set Int32FastValueData = eventArgs.DataSetData.FieldDataDictionary.Item("Int32Fast")
            OutputText = OutputText & Int32FastValueData & vbCrLf
        End If
    Else
        OutputText = OutputText & vbCrLf
        OutputText = OutputText & "*** Failure: " & eventArgs.ErrorMessageBrief & vbCrLf
    End If
End Sub
' Example output:
'
'Subscribing...
'Processing dataset message events for 20 seconds...
'6502 {System.Int32} @2019-10-06T10:02:01.254,647,600,00; Good
'6538 {System.Int32} @2019-10-06T10:02:01.755,010,700,00; Good
'6615 {System.Int32} @2019-10-06T10:02:02.255,780,200,00; Good
'6687 {System.Int32} @2019-10-06T10:02:02.756,495,900,00; Good
'6769 {System.Int32} @2019-10-06T10:02:03.257,320,200,00; Good
'6804 {System.Int32} @2019-10-06T10:02:03.757,667,300,00; Good
'6877 {System.Int32} @2019-10-06T10:02:04.258,405,000,00; Good
'6990 {System.Int32} @2019-10-06T10:02:04.759,532,900,00; Good
'7063 {System.Int32} @2019-10-06T10:02:05.260,257,200,00; Good
'7163 {System.Int32} @2019-10-06T10:02:05.761,261,800,00; Good
'7255 {System.Int32} @2019-10-06T10:02:06.262,176,800,00; Good
'7321 {System.Int32} @2019-10-06T10:02:06.762,839,800,00; Good
'7397 {System.Int32} @2019-10-06T10:02:07.263,598,900,00; Good
'7454 {System.Int32} @2019-10-06T10:02:07.764,168,900,00; Good
'7472 {System.Int32} @2019-10-06T10:02:08.264,350,400,00; Good
'...
REM #endregion Example _EasyUASubscriber.SubscribeDataSet.ExtractField

REM #region Example _EasyUASubscriber.SubscribeDataSet.FieldNames
REM This example shows how to subscribe to dataset messages and specify field names, without having the full metadata.
REM
REM In order to produce network messages for this example, run the UADemoPublisher tool. For documentation, see
REM https://kb.opclabs.com/UADemoPublisher_Basics . In some cases, you may have to specify the interface name to be used.
REM
REM Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

' The subscriber object, with events.
'Public WithEvents Subscriber8 As EasyUASubscriber

Private Sub EasyUASubscriber_SubscribeDataSet_FieldNames_Command_Click()
    OutputText = ""

    ' Define the PubSub connection we will work with.
    Dim subscribeDataSetArguments :  Set subscribeDataSetArguments = CreateObject("OpcLabs.EasyOpc.UA.PubSub.OperationModel.EasyUASubscribeDataSetArguments")
    Dim ConnectionDescriptor :  Set ConnectionDescriptor = subscribeDataSetArguments.dataSetSubscriptionDescriptor.ConnectionDescriptor
    ConnectionDescriptor.ResourceAddress.ResourceDescriptor.UrlString = "opc.udp://239.0.0.1"
    ' In some cases you may have to set the interface (network adapter) name that needs to be used, similarly to
    ' the statement below. Your actual interface name may differ, of course.
    ' ConnectionDescriptor.ResourceAddress.InterfaceName = "Ethernet"

    ' Define the filter. Publisher Id (unsigned 64-bits) is 31, and the dataset writer Id is 1.
    subscribeDataSetArguments.dataSetSubscriptionDescriptor.Filter.PublisherId.SetIdentifier UAPublisherIdType_UInt64, 31
    subscribeDataSetArguments.dataSetSubscriptionDescriptor.Filter.DataSetWriterDescriptor.DataSetWriterId = 1

    ' Define the metadata. For UADP, the order of field metadata must correspond to the order of fields in the dataset message.
    ' Since the encoding is not RawData, we do not have to specify the type information for the fields.
    Dim metaData :  Set metaData = CreateObject("OpcLabs.EasyOpc.UA.PubSub.Configuration.UADataSetMetaData")
    '
    Dim field1 :  Set field1 = CreateObject("OpcLabs.EasyOpc.UA.PubSub.Configuration.UAFieldMetaData")
    field1.Name = "BoolToggle"
    metaData.Add(field1)
    '
    Dim field2 :  Set field2 = CreateObject("OpcLabs.EasyOpc.UA.PubSub.Configuration.UAFieldMetaData")
    field2.Name = "Int32"
    metaData.Add(field2)
    '
    Dim field3 :  Set field3 = CreateObject("OpcLabs.EasyOpc.UA.PubSub.Configuration.UAFieldMetaData")
    field3.Name = "Int32Fast"
    metaData.Add(field3)
    '
    Dim field4 :  Set field4 = CreateObject("OpcLabs.EasyOpc.UA.PubSub.Configuration.UAFieldMetaData")
    field4.Name = "DateTime"
    metaData.Add(field4)
    '
    Set subscribeDataSetArguments.dataSetSubscriptionDescriptor.DataSetMetaData = metaData

    ' Instantiate the subscriber object and hook events.
    Set Subscriber8 = New EasyUASubscriber

    OutputText = OutputText & "Subscribing..." & vbCrLf
    Subscriber8.SubscribeDataSet subscribeDataSetArguments

    OutputText = OutputText & "Processing dataset message events for 20 seconds..." & vbCrLf
    Pause 20 * 1000

    OutputText = OutputText & "Unsubscribing..." & vbCrLf
    Subscriber8.UnsubscribeAllDataSets

    OutputText = OutputText & "Waiting for 1 second..." & vbCrLf
    ' Unsubscribe operation is asynchronous, messages may still come for a short while.
    Pause 1000

    Set Subscriber8 = Nothing

    OutputText = OutputText & "Finished." & vbCrLf
End Sub

Private Sub Subscriber8_DataSetMessage(ByVal sender As Variant, ByVal eventArgs As EasyUADataSetMessageEventArgs)
    ' Display the dataset.
    If eventArgs.Succeeded Then
        ' An event with null DataSetData just indicates a successful connection.
        If Not (eventArgs.DataSetData Is Nothing) Then
            OutputText = OutputText & vbCrLf
            OutputText = OutputText & "Dataset data: " & eventArgs.DataSetData & vbCrLf
            Dim Pair : For Each Pair In eventArgs.DataSetData.FieldDataDictionary
                OutputText = OutputText & Pair & vbCrLf
            Next
        End If
    Else
        OutputText = OutputText & vbCrLf
        OutputText = OutputText & "*** Failure: " & eventArgs.ErrorMessageBrief & vbCrLf
    End If
End Sub
'
' Example output:
'
'Subscribing...
'Processing dataset message events for 20 seconds...
'
'Dataset data: Good; Data; publisher=(UInt64)31, writer=1, fields: 4
'[BoolToggle, True {System.Boolean}; Good]
'[Int32, 25 {System.Int32}; Good]
'[Int32Fast, 928 {System.Int32}; Good]
'[DateTime, 10/3/2019 10:43:01 AM {System.DateTime}; Good]
'
'Dataset data: Good; Data; publisher=(UInt64)31, writer=1, fields: 4
'[Int32, 26 {System.Int32}; Good]
'[Int32Fast, 1007 {System.Int32}; Good]
'[DateTime, 10/3/2019 10:43:02 AM {System.DateTime}; Good]
'[BoolToggle, True {System.Boolean}; Good]
'
'Dataset data: Good; Data; publisher=(UInt64)31, writer=1, fields: 4
'[Int32Fast, 1113 {System.Int32}; Good]
'[DateTime, 10/3/2019 10:43:02 AM {System.DateTime}; Good]
'[BoolToggle, True {System.Boolean}; Good]
'[Int32, 26 {System.Int32}; Good]
'
'Dataset data: Good; Data; publisher=(UInt64)31, writer=1, fields: 4
'[BoolToggle, False {System.Boolean}; Good]
'[Int32, 27 {System.Int32}; Good]
'[Int32Fast, 1201 {System.Int32}; Good]
'[DateTime, 10/3/2019 10:43:03 AM {System.DateTime}; Good]
'
'Dataset data: Good; Data; publisher=(UInt64)31, writer=1, fields: 4
'[Int32Fast, 1260 {System.Int32}; Good]
'[DateTime, 10/3/2019 10:43:03 AM {System.DateTime}; Good]
'[BoolToggle, False {System.Boolean}; Good]
'[Int32, 27 {System.Int32}; Good]
'
'...
'
REM #endregion Example _EasyUASubscriber.SubscribeDataSet.FieldNames

REM #region Example _EasyUASubscriber.SubscribeDataSet.Filter
REM This example shows how to subscribe to dataset messages and specify a filter, on an OPC-UA PubSub connection with
REM UDP UADP mapping.
REM
REM In order to produce network messages for this example, run the UADemoPublisher tool. For documentation, see
REM https://kb.opclabs.com/UADemoPublisher_Basics . In some cases, you may have to specify the interface name to be used.
REM
REM Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

' The subscriber object, with events
'Public WithEvents Subscriber1 As EasyUASubscriber

Private Sub EasyUASubscriber_SubscribeDataSet_Filter_Command_Click()
    OutputText = ""

    ' Define the PubSub connection we will work with.
    Dim subscribeDataSetArguments As New EasyUASubscribeDataSetArguments
    Dim ConnectionDescriptor As UAPubSubConnectionDescriptor
    Set ConnectionDescriptor = subscribeDataSetArguments.dataSetSubscriptionDescriptor.ConnectionDescriptor
    ConnectionDescriptor.ResourceAddress.ResourceDescriptor.UrlString = "opc.udp://239.0.0.1"
    ' In some cases you may have to set the interface (network adapter) name that needs to be used, similarly to
    ' the statement below. Your actual interface name may differ, of course.
    'ConnectionDescriptor.ResourceAddress.InterfaceName := 'Ethernet';

    ' Define the filter. Publisher Id (unsigned 64-bits) is 31, and the dataset writer Id is 1.
    Call subscribeDataSetArguments.dataSetSubscriptionDescriptor.Filter.PublisherId.SetIdentifier(UAPublisherIdType_UInt64, 31)
    subscribeDataSetArguments.dataSetSubscriptionDescriptor.Filter.DataSetWriterDescriptor.DataSetWriterId = 1
    
    ' Instantiate the subscriber object and hook events.
    Set Subscriber1 = New EasyUASubscriber
    
    OutputText = OutputText & "Subscribing..." & vbCrLf
    Call Subscriber1.SubscribeDataSet(subscribeDataSetArguments)

    OutputText = OutputText & "Processing dataset message for 20 seconds..." & vbCrLf
    Pause 20000

    OutputText = OutputText & "Unsubscribing..." & vbCrLf
    Subscriber1.UnsubscribeAllDataSets

    OutputText = OutputText & "Waiting for 1 second..." & vbCrLf
    ' Unsubscribe operation is asynchronous, messages may still come for a short while.
    Pause 1000

    Set Subscriber1 = Nothing

    OutputText = OutputText & "Finished." & vbCrLf
End Sub

Private Sub Subscriber1_OnDataSetMessage(ByVal sender As Variant, ByVal eventArgs As EasyUADataSetMessageEventArgs)
    ' Display the dataset
    If eventArgs.Succeeded Then
        ' An event with null DataSetData just indicates a successful connection.
        If Not eventArgs.DataSetData Is Nothing Then
            OutputText = OutputText & vbCrLf
            OutputText = OutputText & "Dataset data: " & eventArgs.DataSetData & vbCrLf
            Dim dictionaryEntry2 : For Each dictionaryEntry2 In eventArgs.DataSetData.FieldDataDictionary
                OutputText = OutputText & dictionaryEntry2 & vbCrLf
            Next
        End If
    Else
        OutputText = OutputText & vbCrLf
        OutputText = OutputText & eventArgs.ErrorMessageBrief & vbCrLf
    End If
End Sub
' Example output:
'
'Subscribing...
'Processing dataset message events for 20 seconds...
'
'Dataset data: Good; Data; publisher=(UInt64)31, writer=1, fields: 4
'[BoolToggle, True {System.Boolean}; Good]
'[Int32, 25 {System.Int32}; Good]
'[Int32Fast, 928 {System.Int32}; Good]
'[DateTime, 10/3/2019 10:43:01 AM {System.DateTime}; Good]
'
'Dataset data: Good; Data; publisher=(UInt64)31, writer=1, fields: 4
'[Int32, 26 {System.Int32}; Good]
'[Int32Fast, 1007 {System.Int32}; Good]
'[DateTime, 10/3/2019 10:43:02 AM {System.DateTime}; Good]
'[BoolToggle, True {System.Boolean}; Good]
'
'Dataset data: Good; Data; publisher=(UInt64)31, writer=1, fields: 4
'[Int32Fast, 1113 {System.Int32}; Good]
'[DateTime, 10/3/2019 10:43:02 AM {System.DateTime}; Good]
'[BoolToggle, True {System.Boolean}; Good]
'[Int32, 26 {System.Int32}; Good]
'
'Dataset data: Good; Data; publisher=(UInt64)31, writer=1, fields: 4
'[BoolToggle, False {System.Boolean}; Good]
'[Int32, 27 {System.Int32}; Good]
'[Int32Fast, 1201 {System.Int32}; Good]
'[DateTime, 10/3/2019 10:43:03 AM {System.DateTime}; Good]
'
'Dataset data: Good; Data; publisher=(UInt64)31, writer=1, fields: 4
'[Int32Fast, 1260 {System.Int32}; Good]
'[DateTime, 10/3/2019 10:43:03 AM {System.DateTime}; Good]
'[BoolToggle, False {System.Boolean}; Good]
'[Int32, 27 {System.Int32}; Good]
'
'...
REM #endregion Example _EasyUASubscriber.SubscribeDataSet.Filter

REM #region Example _EasyUASubscriber.SubscribeDataSet.Main1
REM This example shows how to subscribe to all dataset messages on an OPC-UA PubSub connection with UDP UADP mapping.
REM
REM In order to produce network messages for this example, run the UADemoPublisher tool. For documentation, see
REM https://kb.opclabs.com/UADemoPublisher_Basics . In some cases, you may have to specify the interface name to be used.
REM
REM Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

' The subscriber object, with events
'Public WithEvents Subscriber2 As EasyUASubscriber

Private Sub EasyUASubscriber_SubscribeDataSet_Main1_Command_Click()
    OutputText = ""

    ' Define the PubSub connection we will work with.
    Dim subscribeDataSetArguments As New EasyUASubscribeDataSetArguments
    Dim ConnectionDescriptor As UAPubSubConnectionDescriptor
    Set ConnectionDescriptor = subscribeDataSetArguments.dataSetSubscriptionDescriptor.ConnectionDescriptor
    ConnectionDescriptor.ResourceAddress.ResourceDescriptor.UrlString = "opc.udp://239.0.0.1"
    
    ' In some cases you may have to set the interface (network adapter) name that needs to be used, similarly to
    ' the statement below. Your actual interface name may differ, of course.
    'ConnectionDescriptor.ResourceAddress.InterfaceName := 'Ethernet';

    ' Instantiate the subscriber object and hook events.
    Set Subscriber2 = New EasyUASubscriber
    
    OutputText = OutputText & "Subscribing..." & vbCrLf
    Subscriber2.SubscribeDataSet subscribeDataSetArguments

    OutputText = OutputText & "Processing dataset message for 20 seconds..." & vbCrLf
    Pause 20000

    OutputText = OutputText & "Unsubscribing..." & vbCrLf
    Subscriber2.UnsubscribeAllDataSets

    OutputText = OutputText & "Waiting for 1 second..." & vbCrLf
    ' Unsubscribe operation is asynchronous, messages may still come for a short while.
    Pause 1000

    Set Subscriber2 = Nothing

    OutputText = OutputText & "Finished." & vbCrLf
End Sub

Private Sub Subscriber2_DataSetMessage(ByVal sender As Variant, ByVal eventArgs As EasyUADataSetMessageEventArgs)
    ' Display the dataset
    If eventArgs.Succeeded Then
        ' An event with null DataSetData just indicates a successful connection.
        If Not eventArgs.DataSetData Is Nothing Then
            OutputText = OutputText & vbCrLf
            OutputText = OutputText & "Dataset data: " & eventArgs.DataSetData & vbCrLf
            Dim dictionaryEntry2 : For Each dictionaryEntry2 In eventArgs.DataSetData.FieldDataDictionary
                OutputText = OutputText & dictionaryEntry2 & vbCrLf
            Next
        End If
    Else
        OutputText = OutputText & vbCrLf
        OutputText = OutputText & eventArgs.ErrorMessageBrief & vbCrLf
    End If
End Sub
' Example output:
'
'Subscribing...
'Processing dataset message events for 20 seconds...
'
'Dataset data: Good; Data; publisher="32", writer=1, class=eae79794-1af7-4f96-8401-4096cd1d8908, fields: 4
'[#0, True {System.Boolean} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#1, 7945 {System.Int32} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#2, 5246 {System.Int32} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#3, 9/30/2019 11:19:14 AM {System.DateTime} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'
'Dataset data: Good; Data; publisher="32", writer=3, class=96976b7b-0db7-46c3-a715-0979884b55ae, fields: 100
'[#0, 45 {System.Int64} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#1, 145 {System.Int64} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#2, 245 {System.Int64} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#3, 345 {System.Int64} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#4, 445 {System.Int64} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#5, 545 {System.Int64} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#6, 645 {System.Int64} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#7, 745 {System.Int64} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#8, 845 {System.Int64} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#9, 945 {System.Int64} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#10, 1045 {System.Int64} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'...
REM #endregion Example _EasyUASubscriber.SubscribeDataSet.Main1

REM #region Example _EasyUASubscriber.SubscribeDataSet.Metadata
REM This example shows how to subscribe to dataset messages with RawData field encoding, specifying the metadata necessary
REM for their decoding directly in the code.
REM
REM In order to produce network messages for this example, run the UADemoPublisher tool. For documentation, see
REM https://kb.opclabs.com/UADemoPublisher_Basics . In some cases, you may have to specify the interface name to be used.
REM
REM Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

' The subscriber object, with events
'Public WithEvents Subscriber3 As EasyUASubscriber

Private Sub EasyUASubscriber_SubscribeDataSet_Metadata_Command_Click()
    OutputText = ""

    ' Define the PubSub connection we will work with.
    Dim subscribeDataSetArguments As New EasyUASubscribeDataSetArguments
    Dim ConnectionDescriptor As UAPubSubConnectionDescriptor
    Set ConnectionDescriptor = subscribeDataSetArguments.dataSetSubscriptionDescriptor.ConnectionDescriptor
    ConnectionDescriptor.ResourceAddress.ResourceDescriptor.UrlString = "opc.udp://239.0.0.1"
    ' In some cases you may have to set the interface (network adapter) name that needs to be used, similarly to
    ' the statement below. Your actual interface name may differ, of course.
    'ConnectionDescriptor.ResourceAddress.InterfaceName := 'Ethernet';

    ' Define the filter. Publisher Id (unsigned 16-bits) is 30, and the writer group Id is 101.
    ' The dataset writer Id (1) must not be specified in the filter, because it does not appear in the message.
    Call subscribeDataSetArguments.dataSetSubscriptionDescriptor.Filter.PublisherId.SetUInt16Identifier(30)
    subscribeDataSetArguments.dataSetSubscriptionDescriptor.Filter.WriterGroupDescriptor.WriterGroupId = 101

    ' Define the metadata. For UADP, the order of field metadata must correspond to the order of fields in the dataset message.
    Dim metaData As New UADataSetMetaData
    '
    Dim field1 As New UAFieldMetaData
    field1.BuiltInType = UABuiltInType_Boolean
    field1.Name = "BoolToggle"
    metaData.Add field1
    '
    Dim field2 As New UAFieldMetaData
    field2.BuiltInType = UABuiltInType_Int32
    field2.Name = "Int32"
    metaData.Add field2
    '
    Dim field3 As New UAFieldMetaData
    field3.BuiltInType = UABuiltInType_Int32
    field3.Name = "Int32Fast"
    metaData.Add field3
    '
    Dim field4 As New UAFieldMetaData
    field4.BuiltInType = UABuiltInType_DateTime
    field4.Name = "DateTime"
    metaData.Add field4
    '
    Set subscribeDataSetArguments.dataSetSubscriptionDescriptor.DataSetMetaData = metaData
    
    ' Define the specific communication parameters for the dataset subscription.
    ' The dataset offset is needed with messages that do not contain dataset writer Ids and use RawData field
    ' encoding. An exception to this rule is when the dataset is the only or first in the dataset message payload,
    ' which is also the case here, but we are specifying the dataset offset anyway, for illustration.
    subscribeDataSetArguments.dataSetSubscriptionDescriptor.CommunicationParameters.UadpDataSetReaderMessageParameters.DataSetOffset = 15
    
    ' Instantiate the subscriber object and hook events.
    Set Subscriber3 = New EasyUASubscriber
    
    OutputText = OutputText & "Subscribing..." & vbCrLf
    Call Subscriber3.SubscribeDataSet(subscribeDataSetArguments)

    OutputText = OutputText & "Processing dataset message for 20 seconds..." & vbCrLf
    Pause 20000

    OutputText = OutputText & "Unsubscribing..." & vbCrLf
    Subscriber3.UnsubscribeAllDataSets

    OutputText = OutputText & "Waiting for 1 second..." & vbCrLf
    ' Unsubscribe operation is asynchronous, messages may still come for a short while.
    Pause 1000

    Set Subscriber3 = Nothing

    OutputText = OutputText & "Finished." & vbCrLf
End Sub

Private Sub Subscriber3_DataSetMessage(ByVal sender As Variant, ByVal eventArgs As EasyUADataSetMessageEventArgs)
    ' Display the dataset
    If eventArgs.Succeeded Then
        ' An event with null DataSetData just indicates a successful connection.
        If Not eventArgs.DataSetData Is Nothing Then
            OutputText = OutputText & vbCrLf
            OutputText = OutputText & "Dataset data: " & eventArgs.DataSetData & vbCrLf
            Dim dictionaryEntry2 : For Each dictionaryEntry2 In eventArgs.DataSetData.FieldDataDictionary
                OutputText = OutputText & dictionaryEntry2 & vbCrLf
            Next
        End If
    Else
        OutputText = OutputText & vbCrLf
        OutputText = OutputText & "*** Failure: " & eventArgs.ErrorMessageBrief & vbCrLf
    End If
End Sub
' Example output:
'
'Subscribing...
'Processing dataset message events for 20 seconds...
'
'Dataset data: Good; Data; publisher=(UInt16)30, group=101, fields: 4
'[BoolToggle, False {System.Boolean}; Good]
'[Int32, 3072 {System.Int32}; Good]
'[Int32Fast, 894 {System.Int32}; Good]
'[DateTime, 10/1/2019 12:21:14 PM {System.DateTime}; Good]
'
'Dataset data: Good; Data; publisher=(UInt16)30, group=101, fields: 4
'[BoolToggle, False {System.Boolean}; Good]
'[Int32, 3072 {System.Int32}; Good]
'[Int32Fast, 920 {System.Int32}; Good]
'[DateTime, 10/1/2019 12:21:14 PM {System.DateTime}; Good]
'
'Dataset data: Good; Data; publisher=(UInt16)30, group=101, fields: 4
'[BoolToggle, False {System.Boolean}; Good]
'[Int32, 3073 {System.Int32}; Good]
'[Int32Fast, 1003 {System.Int32}; Good]
'[DateTime, 10/1/2019 12:21:15 PM {System.DateTime}; Good]
'
'Dataset data: Good; Data; publisher=(UInt16)30, group=101, fields: 4
'[BoolToggle, False {System.Boolean}; Good]
'[Int32, 3073 {System.Int32}; Good]
'[Int32Fast, 1074 {System.Int32}; Good]
'[DateTime, 10/1/2019 12:21:15 PM {System.DateTime}; Good]
'
'Dataset data: Good; Data; publisher=(UInt16)30, group=101, fields: 4
'[BoolToggle, True {System.Boolean}; Good]
'[Int32, 3074 {System.Int32}; Good]
'[Int32Fast, 1140 {System.Int32}; Good]
'[DateTime, 10/1/2019 12:21:16 PM {System.DateTime}; Good]
'
'...
REM #endregion Example _EasyUASubscriber.SubscribeDataSet.Metadata

REM #region Example _EasyUASubscriber.SubscribeDataSet.MqttJsonTcp
REM The following package needs to be referenced in your project (or otherwise made available) for the MQTT transport to 
REM work.
REM - OpcLabs.MqttNet
REM Refer to the documentation for more information.
REM
REM Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

' The subscriber object, with events
'Public WithEvents Subscriber4 As EasyUASubscriber

Private Sub EasyUASubscriber_SubscribeDataSet_MqttJsonTcp_Command_Click()
    OutputText = ""

    ' Define the PubSub connection we will work with. Uses implicit conversion from a string.
    ' Default port with MQTT is 1883.
    Dim subscribeDataSetArguments As New EasyUASubscribeDataSetArguments
    Dim pubSubConnectionDescriptor As UAPubSubConnectionDescriptor
    Set pubSubConnectionDescriptor = subscribeDataSetArguments.dataSetSubscriptionDescriptor.ConnectionDescriptor
    pubSubConnectionDescriptor.ResourceAddress.ResourceDescriptor.UrlString = "mqtt://opcua-pubsub.demo-this.com:1883"
    ' Specify the transport protocol mapping.
    ' The statement below isn't actually necessary, due to automatic message mapping recognition feature; see
    ' https://kb.opclabs.com/OPC_UA_PubSub_Automatic_Message_Mapping_Recognition for more details.
    pubSubConnectionDescriptor.TransportProfileUriString = "http://opcfoundation.org/UA-Profile/Transport/pubsub-mqtt-json" ' UAPubSubTransportProfileUriStrings.MqttJson

    ' Define the arguments for subscribing to the dataset, specifying the MQTT topic name.
    subscribeDataSetArguments.dataSetSubscriptionDescriptor.CommunicationParameters.BrokerDataSetReaderTransportParameters.QueueName = "opcuademo/json"
    
    ' Instantiate the subscriber object and hook events.
    Set Subscriber4 = New EasyUASubscriber
    
    OutputText = OutputText & "Subscribing..." & vbCrLf
    Call Subscriber4.SubscribeDataSet(subscribeDataSetArguments)

    OutputText = OutputText & "Processing dataset message for 20 seconds..." & vbCrLf
    Pause 20000

    OutputText = OutputText & "Unsubscribing..." & vbCrLf
    Subscriber4.UnsubscribeAllDataSets

    OutputText = OutputText & "Waiting for 1 second..." & vbCrLf
    ' Unsubscribe operation is asynchronous, messages may still come for a short while.
    Pause 1000

    Set Subscriber1 = Nothing

    OutputText = OutputText & "Finished." & vbCrLf
End Sub

Private Sub Subscriber4_DataSetMessage(ByVal sender As Variant, ByVal eventArgs As EasyUADataSetMessageEventArgs)
    ' Display the dataset
    If eventArgs.Succeeded Then
        ' An event with null DataSetData just indicates a successful connection.
        If Not eventArgs.DataSetData Is Nothing Then
            OutputText = OutputText & vbCrLf
            OutputText = OutputText & "Dataset data: " & eventArgs.DataSetData & vbCrLf
            Dim dictionaryEntry2 : For Each dictionaryEntry2 In eventArgs.DataSetData.FieldDataDictionary
                OutputText = OutputText & dictionaryEntry2 & vbCrLf
            Next
        End If
    Else
        OutputText = OutputText & vbCrLf
        OutputText = OutputText & eventArgs.ErrorMessageBrief & vbCrLf
    End If
End Sub
' Example output:
'
'Subscribing...
'Processing dataset message events for 20 seconds...
'
'Dataset data: 2020-01-21T17:07:19.778,836,700,00; Good; Data; publisher=[String]31, class=eae79794-1af7-4f96-8401-4096cd1d8908, fields: 4
'[BoolToggle, True {System.Boolean} @2020-01-21T16:07:19.778,836,700,00; Good]
'[Int32, 482 {System.Int64} @2020-01-21T16:07:19.778,836,700,00; Good]
'[Int32Fast, 2287 {System.Int64} @2020-01-21T16:07:19.778,836,700,00; Good]
'[DateTime, 1/21/2020 5:07:19 PM {System.DateTime} @2020-01-21T16:07:19.778,836,700,00; Good]
'
'Dataset data: Good; Data; publisher=[String]32, fields: 4
'[BoolToggle, True {System.Boolean}; Good]
'[Int32, 482 {System.Int32}; Good]
'[Int32Fast, 2287 {System.Int32}; Good]
'[DateTime, 1/21/2020 5:07:19 PM {System.DateTime}; Good]
'
'Dataset data: Good; Data; publisher=[String]32, fields: 100
'[Mass_0, 82 {System.Int64}; Good]
'[Mass_1, 182 {System.Int64}; Good]
'[Mass_2, 282 {System.Int64}; Good]
'[Mass_3, 382 {System.Int64}; Good]
'[Mass_4, 482 {System.Int64}; Good]
'[Mass_5, 582 {System.Int64}; Good]
'[Mass_6, 682 {System.Int64}; Good]
'[Mass_7, 782 {System.Int64}; Good]
'[Mass_8, 882 {System.Int64}; Good]
'[Mass_9, 982 {System.Int64}; Good]
'[Mass_10, 1082 {System.Int64}; Good]
'...

REM #endregion Example _EasyUASubscriber.SubscribeDataSet.MqttJsonTcp

REM #region Example _EasyUASubscriber.SubscribeDataSet.PublisherId
REM This example shows how to subscribe to all dataset messages with specific publisher Id, on an OPC-UA PubSub connection
REM with UDP UADP mapping.
REM
REM In order to produce network messages for this example, run the UADemoPublisher tool. For documentation, see
REM https://kb.opclabs.com/UADemoPublisher_Basics . In some cases, you may have to specify the interface name to be used.
REM
REM Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

' The subscriber object, with events
'Public WithEvents Subscriber5 As EasyUASubscriber

Private Sub EasyUASubscriber_SubscribeDataSet_PublisherId_Command_Click()
    OutputText = ""
    ' Define the PubSub connection we will work with.
    Dim subscribeDataSetArguments As New EasyUASubscribeDataSetArguments
    Dim ConnectionDescriptor As UAPubSubConnectionDescriptor
    Set ConnectionDescriptor = subscribeDataSetArguments.dataSetSubscriptionDescriptor.ConnectionDescriptor
    ConnectionDescriptor.ResourceAddress.ResourceDescriptor.UrlString = "opc.udp://239.0.0.1"
    ' In some cases you may have to set the interface (network adapter) name that needs to be used, similarly to
    ' the statement below. Your actual interface name may differ, of course.
    'ConnectionDescriptor.ResourceAddress.InterfaceName := 'Ethernet';

    ' Define the arguments for subscribing to the dataset, where the filter is (unsigned 64-bit) publisher Id 31.
    Call subscribeDataSetArguments.dataSetSubscriptionDescriptor.Filter.PublisherId.SetIdentifier(UAPublisherIdType_UInt64, 31)
    
    ' Instantiate the subscriber object and hook events.
    Set Subscriber5 = New EasyUASubscriber
    
    OutputText = OutputText & "Subscribing..." & vbCrLf
    Call Subscriber5.SubscribeDataSet(subscribeDataSetArguments)

    OutputText = OutputText & "Processing dataset message for 20 seconds..." & vbCrLf
    Pause 20000

    OutputText = OutputText & "Unsubscribing..." & vbCrLf
    Subscriber5.UnsubscribeAllDataSets

    OutputText = OutputText & "Waiting for 1 second..." & vbCrLf
    ' Unsubscribe operation is asynchronous, messages may still come for a short while.
    Pause 1000

    Set Subscriber5 = Nothing

    OutputText = OutputText & "Finished." & vbCrLf
End Sub

Private Sub Subscriber5_DataSetMessage(ByVal sender As Variant, ByVal eventArgs As EasyUADataSetMessageEventArgs)
    ' Display the dataset
    If eventArgs.Succeeded Then
        ' An event with null DataSetData just indicates a successful connection.
        If Not eventArgs.DataSetData Is Nothing Then
            OutputText = OutputText & vbCrLf
            OutputText = OutputText & "Dataset data: " & eventArgs.DataSetData & vbCrLf
            Dim dictionaryEntry2 : For Each dictionaryEntry2 In eventArgs.DataSetData.FieldDataDictionary
                OutputText = OutputText & dictionaryEntry2 & vbCrLf
            Next
        End If
    Else
        OutputText = OutputText & vbCrLf
        OutputText = OutputText & "*** Failure: " & eventArgs.ErrorMessageBrief & vbCrLf
    End If
    ' Example output:
    '
    'Subscribing...
    'Processing dataset message events for 20 seconds...
    '
    'Dataset data: Good; Event; publisher=(UInt64)31, writer=51, fields: 4
    '[#0, True {System.Boolean}; Good]
    '[#1, 1237 {System.Int32}; Good]
    '[#2, 2514 {System.Int32}; Good]
    '[#3, 10/1/2019 9:03:59 AM {System.DateTime}; Good]
    '
    'Dataset data: Good; Data; publisher=(UInt64)31, writer=1, fields: 4
    '[#0, False {System.Boolean}; Good]
    '[#1, 1239 {System.Int32}; Good]
    '[#2, 2703 {System.Int32}; Good]
    '[#3, 10/1/2019 9:04:01 AM {System.DateTime}; Good]
    '
    'Dataset data: Good; Data; publisher=(UInt64)31, writer=4, fields: 16
    '[#0, False {System.Boolean}; Good]
    '[#1, 215 {System.Byte}; Good]
    '[#2, 1239 {System.Int16}; Good]
    '[#3, 1239 {System.Int32}; Good]
    '[#4, 1239 {System.Int64}; Good]
    '[#5, 87 {System.Int16}; Good]
    '[#6, 1239 {System.Int32}; Good]
    '[#7, 1239 {System.Int64}; Good]
    '[#8, 1239 {System.Decimal}; Good]
    '[#9, 1239 {System.Single}; Good]
    '[#10, 1239 {System.Double}; Good]
    '[#11, Romeo {System.String}; Good]
    '[#12, [20] {175, 186, 248, 246, 215, ...} {System.Byte[]}; Good]
    '[#13, d4492ca8-35c8-4b98-8edf-6ffa5ca041ca {System.Guid}; Good]
    '[#14, 10/1/2019 9:04:01 AM {System.DateTime}; Good]
    '[#15, [10] {1239, 1240, 1241, 1242, 1243, ...} {System.Int64[]}; Good]
    '
    'Dataset data: Good; Data; publisher=(UInt64)31, writer=1, fields: 4
    '[#2, 2722 {System.Int32}; Good]
    '[#3, 10/1/2019 9:04:01 AM {System.DateTime}; Good]
    '[#0, False {System.Boolean}; Good]
    '[#1, 1239 {System.Int32}; Good]
    '
    'Dataset data: Good; Data; publisher=(UInt64)31, writer=3, fields: 100
    '[#0, 39 {System.Int64}; Good]
    '[#1, 139 {System.Int64}; Good]
    '[#2, 239 {System.Int64}; Good]
    '[#3, 339 {System.Int64}; Good]
    '[#4, 439 {System.Int64}; Good]
    '[#5, 539 {System.Int64}; Good]
    '[#6, 639 {System.Int64}; Good]
    '[#7, 739 {System.Int64}; Good]
    '[#8, 839 {System.Int64}; Good]
    '[#9, 939 {System.Int64}; Good]
    '[#10, 1039 {System.Int64}; Good]
    '...
End Sub

REM #endregion Example _EasyUASubscriber.SubscribeDataSet.PublisherId

REM #region Example _EasyUASubscriber.SubscribeDataSet.Secure
REM This example shows how to securely subscribe to signed and encrypted dataset messages.
REM An external Security Key Service (SKS) is needed (not a part of QuickOPC).
REM
REM The network messages for this example can be published e.g. using the UADemoPublisher tool - see
REM https://kb.opclabs.com/How_to_publish_or_subscribe_to_secure_OPC_UA_PubSub_messages .
REM
REM Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

' The subscriber object, with events
'Public WithEvents Subscriber6 As EasyUASubscriber

Private Sub EasyUASubscriber_SubscribeDataSet_Secure_Command_Click()
    OutputText = ""

    ' Define the PubSub connection we will work with.
    Dim subscribeDataSetArguments As New EasyUASubscribeDataSetArguments
    Dim ConnectionDescriptor As UAPubSubConnectionDescriptor
    Set ConnectionDescriptor = subscribeDataSetArguments.dataSetSubscriptionDescriptor.ConnectionDescriptor
    ConnectionDescriptor.ResourceAddress.ResourceDescriptor.UrlString = "opc.udp://239.0.0.1"
    ' In some cases you may have to set the interface (network adapter) name that needs to be used, similarly to
    ' the statement below. Your actual interface name may differ, of course.
    'ConnectionDescriptor.ResourceAddress.InterfaceName := 'Ethernet';

    ' Define the arguments for subscribing to the dataset.
    Dim comunicationParameters As New UASubscriberCommunicationParameters
    ' Specifies the security mode for the PubSub network messages received. This is a minimum security
    ' mode that you want to accept.
    comunicationParameters.SecurityMode = UAMessageSecurityModes_SecuritySignAndEncrypt
    ' Specifies the URL of the SKS (Security Key Service) endpoint.
    comunicationParameters.SecurityKeyServiceTemplate.UrlString = "opc.tcp://localhost:48010"
    ' Specifies the security mode that will be used to connect to the SKS.
    Dim endpointSelectionPolicy As New UAEndpointSelectionPolicy
    endpointSelectionPolicy.AllowedMessageSecurityModes = UAMessageSecurityModes_SecuritySignAndEncrypt
    Set comunicationParameters.SecurityKeyServiceTemplate.endpointSelectionPolicy = endpointSelectionPolicy ' UAMessageSecurityModes_SecuritySignAndEncrypt
    ' Specifies the user name and password used for "logging in" to the SKS.
    comunicationParameters.SecurityKeyServiceTemplate.UserIdentity.UserNameTokenInfo.UserName = "root"
    comunicationParameters.SecurityKeyServiceTemplate.UserIdentity.UserNameTokenInfo.Password = "secret"
    ' Specifies the Id of the security group in the SKS that will be used (the security group in the
    ' SKS is configured to use certain security policy, and has other parameters detailing how the
    ' security keys are generated).
    comunicationParameters.securityGroupId = "TestGroup"
    
    Set subscribeDataSetArguments.dataSetSubscriptionDescriptor.CommunicationParameters = comunicationParameters
    
    ' Instantiate the subscriber object and hook events.
    Set Subscriber6 = New EasyUASubscriber
        
    OutputText = OutputText & "Subscribing..." & vbCrLf
    Call Subscriber6.SubscribeDataSet(subscribeDataSetArguments)

    OutputText = OutputText & "Processing dataset message for 20 seconds..." & vbCrLf
    Pause 20000

    OutputText = OutputText & "Unsubscribing..." & vbCrLf
    Subscriber6.UnsubscribeAllDataSets

    OutputText = OutputText & "Waiting for 1 second..." & vbCrLf
    ' Unsubscribe operation is asynchronous, messages may still come for a short while.
    Pause 1000

    Set Subscriber6 = Nothing

    OutputText = OutputText & "Finished." & vbCrLf
End Sub

Private Sub Subscriber6_DataSetMessage(ByVal sender As Variant, ByVal eventArgs As EasyUADataSetMessageEventArgs)
    ' Display the dataset
    If eventArgs.Succeeded Then
        ' An event with null DataSetData just indicates a successful connection.
        If Not eventArgs.DataSetData Is Nothing Then
            OutputText = OutputText & vbCrLf
            OutputText = OutputText & "Dataset data: " & eventArgs.DataSetData & vbCrLf
            Dim dictionaryEntry2 : For Each dictionaryEntry2 In eventArgs.DataSetData.FieldDataDictionary
                OutputText = OutputText & dictionaryEntry2 & vbCrLf
            Next
        End If
    Else
        OutputText = OutputText & vbCrLf
        OutputText = OutputText & eventArgs.ErrorMessageBrief & vbCrLf
    End If
End Sub
REM #endregion Example _EasyUASubscriber.SubscribeDataSet.Secure

REM #region Example _EasyUASubscriber.UnsubscribeDataSet.Main1
REM This example shows how to subscribe to dataset messages on an OPC-UA PubSub connection, and then unsubscribe from that
REM dataset.
REM
REM In order to produce network messages for this example, run the UADemoPublisher tool. For documentation, see
REM https://kb.opclabs.com/UADemoPublisher_Basics . In some cases, you may have to specify the interface name to be used.
REM
REM Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

' The subscriber object, with events.
'Public WithEvents Subscriber11 As EasyUASubscriber

Private Sub EasyUASubscriber_UnsubscribeDataSet_Main1_Command_Click()
    OutputText = ""
    
    ' Define the PubSub connection we will work with.
    Dim subscribeDataSetArguments: Set subscribeDataSetArguments = CreateObject("OpcLabs.EasyOpc.UA.PubSub.OperationModel.EasyUASubscribeDataSetArguments")
    Dim ConnectionDescriptor: Set ConnectionDescriptor = subscribeDataSetArguments.dataSetSubscriptionDescriptor.ConnectionDescriptor
    ConnectionDescriptor.ResourceAddress.ResourceDescriptor.UrlString = "opc.udp://239.0.0.1"
    ' In some cases you may have to set the interface (network adapter) name that needs to be used, similarly to
    ' the statement below. Your actual interface name may differ, of course.
    ' ConnectionDescriptor.ResourceAddress.InterfaceName = "Ethernet"

    ' Define the filter. Publisher Id (unsigned 64-bits) is 31, and the dataset writer Id is 1.
    subscribeDataSetArguments.dataSetSubscriptionDescriptor.Filter.PublisherId.SetIdentifier UAPublisherIdType_UInt64, 31
    subscribeDataSetArguments.dataSetSubscriptionDescriptor.Filter.DataSetWriterDescriptor.DataSetWriterId = 1

    ' Instantiate the subscriber object and hook events.
    Set Subscriber11 = New EasyUASubscriber

    OutputText = OutputText & "Subscribing..." & vbCrLf
    Subscriber11.SubscribeDataSet subscribeDataSetArguments

    OutputText = OutputText & "Processing dataset message events for 20 seconds..." & vbCrLf
    Pause 20 * 1000

    OutputText = OutputText & "Unsubscribing..." & vbCrLf
    Subscriber11.UnsubscribeAllDataSets

    OutputText = OutputText & "Waiting for 1 second..." & vbCrLf
    ' Unsubscribe operation is asynchronous, messages may still come for a short while.
    Pause 1000

    Set Subscriber11 = Nothing

    OutputText = OutputText & "Finished." & vbCrLf
End Sub

Private Sub Subscriber11_DataSetMessage(ByVal sender As Variant, ByVal eventArgs As EasyUADataSetMessageEventArgs)
    ' Display the dataset.
    If eventArgs.Succeeded Then
        ' An event with null DataSetData just indicates a successful connection.
        If Not (eventArgs.DataSetData Is Nothing) Then
            OutputText = OutputText & vbCrLf
            OutputText = OutputText & "Dataset data: " & eventArgs.DataSetData & vbCrLf
            Dim Pair: For Each Pair In eventArgs.DataSetData.FieldDataDictionary
                OutputText = OutputText & Pair & vbCrLf
            Next
        End If
    Else
        OutputText = OutputText & vbCrLf
        OutputText = OutputText & "*** Failure: " & eventArgs.ErrorMessageBrief & vbCrLf
    End If
End Sub
'
' Example output:
'
'Subscribing...
'Processing dataset message events for 20 seconds...
'
'Dataset data: Good; Data; publisher=(UInt64)31, writer=1, fields: 4
'[#0, True {System.Boolean}; Good]
'[#1, 7134 {System.Int32}; Good]
'[#2, 7364 {System.Int32}; Good]
'[#3, 10/1/2019 10:42:16 AM {System.DateTime}; Good]
'
'Dataset data: Good; Data; publisher=(UInt64)31, writer=1, fields: 4
'[#1, 7135 {System.Int32}; Good]
'[#2, 7429 {System.Int32}; Good]
'[#3, 10/1/2019 10:42:17 AM {System.DateTime}; Good]
'[#0, True {System.Boolean}; Good]
'
'Dataset data: Good; Data; publisher=(UInt64)31, writer=1, fields: 4
'[#2, 7495 {System.Int32}; Good]
'[#3, 10/1/2019 10:42:17 AM {System.DateTime}; Good]
'[#0, True {System.Boolean}; Good]
'[#1, 7135 {System.Int32}; Good]
'
'Dataset data: Good; Data; publisher=(UInt64)31, writer=1, fields: 4
'[#1, 7136 {System.Int32}; Good]
'[#2, 7560 {System.Int32}; Good]
'[#3, 10/1/2019 10:42:18 AM {System.DateTime}; Good]
'[#0, True {System.Boolean}; Good]
'
'Dataset data: Good; Data; publisher=(UInt64)31, writer=1, fields: 4
'[#2, 7626 {System.Int32}; Good]
'[#3, 10/1/2019 10:42:18 AM {System.DateTime}; Good]
'[#0, True {System.Boolean}; Good]
'[#1, 7136 {System.Int32}; Good]
'
'...
'
REM #endregion Example _EasyUASubscriber.UnsubscribeDataSet.Main1
