VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents Client As EasyDAClient
Attribute Client.VB_VarHelpID = -1

REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Private Sub Form_Load()
    Set Client = New EasyDAClient

    Dim ItemSubscriptionArguments1: Set ItemSubscriptionArguments1 = New EasyDAItemSubscriptionArguments
    ItemSubscriptionArguments1.ServerDescriptor.ServerClass = "OPCLabs.KitServer.2"
    ItemSubscriptionArguments1.ItemDescriptor.ItemId = "Simulation.Random"
    ItemSubscriptionArguments1.GroupParameters.RequestedUpdateRate = 1000
    
    Dim ItemSubscriptionArguments2: Set ItemSubscriptionArguments2 = New EasyDAItemSubscriptionArguments
    ItemSubscriptionArguments2.ServerDescriptor.ServerClass = "OPCLabs.KitServer.2"
    ItemSubscriptionArguments2.ItemDescriptor.ItemId = "Trends.Ramp (1 min)"
    ItemSubscriptionArguments2.GroupParameters.RequestedUpdateRate = 1000
    
    Dim ItemSubscriptionArguments3: Set ItemSubscriptionArguments3 = New EasyDAItemSubscriptionArguments
    ItemSubscriptionArguments3.ServerDescriptor.ServerClass = "OPCLabs.KitServer.2"
    ItemSubscriptionArguments3.ItemDescriptor.ItemId = "Trends.Sine (1 min)"
    ItemSubscriptionArguments3.GroupParameters.RequestedUpdateRate = 1000
    
    Dim ItemSubscriptionArguments4: Set ItemSubscriptionArguments4 = New EasyDAItemSubscriptionArguments
    ItemSubscriptionArguments4.ServerDescriptor.ServerClass = "OPCLabs.KitServer.2"
    ItemSubscriptionArguments4.ItemDescriptor.ItemId = "Simulation.Register_I4"
    ItemSubscriptionArguments4.GroupParameters.RequestedUpdateRate = 1000
    
    Dim arguments(3)
    Set arguments(0) = ItemSubscriptionArguments1
    Set arguments(1) = ItemSubscriptionArguments2
    Set arguments(2) = ItemSubscriptionArguments3
    Set arguments(3) = ItemSubscriptionArguments4

    Client.SubscribeMultipleItems arguments
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Client.UnsubscribeAllItems
    Set Client = Nothing
End Sub

Private Sub Client_ItemChanged(ByVal varSender As Variant, ByVal varE As EasyDAItemChangedEventArgs)
    List1.AddItem varE.arguments.ItemDescriptor.ItemId & ": " & varE.Vtq.ToString()
End Sub

