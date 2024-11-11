VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Private Sub Form_Load()

    ' Create EasyOPC-DA component
    Dim Client As New EasyDAClient

    ' Write values into 3 items at once

    Dim ItemValueArguments1 As New DAItemValueArguments
    ItemValueArguments1.ServerDescriptor.ServerClass = "OPCLabs.KitServer.2"
    ItemValueArguments1.ItemDescriptor.ItemId = "Simulation.Register_I4"
    ItemValueArguments1.SetValue(23456)

    Dim ItemValueArguments2 As New DAItemValueArguments
    ItemValueArguments2.ServerDescriptor.ServerClass = "OPCLabs.KitServer.2"
    ItemValueArguments2.ItemDescriptor.ItemId = "Simulation.Register_R8"
    ItemValueArguments2.SetValue(2.3456789)

    Dim ItemValueArguments3 As New DAItemValueArguments
    ItemValueArguments3.ServerDescriptor.ServerClass = "OPCLabs.KitServer.2"
    ItemValueArguments3.ItemDescriptor.ItemId = "Simulation.Register_BSTR"
    ItemValueArguments3.SetValue("ABC")

    Dim arguments(2) As Variant
    Set arguments(0) = ItemValueArguments1
    Set arguments(1) = ItemValueArguments2
    Set arguments(2) = ItemValueArguments3
    
    Call Client.WriteMultipleItemValues(arguments)

End Sub
