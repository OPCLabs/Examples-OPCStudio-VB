VERSION 5.00
Begin VB.Form MainForm 
   Caption         =   "Main"
   ClientHeight    =   3945
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3225
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   3225
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton DataAccessXmlCommand 
      Caption         =   "DataAccess.Xml"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   3015
   End
   Begin VB.CommandButton OpcServerDialogCommand 
      Caption         =   "_OpcServerDialog"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   3015
   End
   Begin VB.CommandButton OpcBrowseDialogCommand 
      Caption         =   "_OpcBrowseDialog"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   3015
   End
   Begin VB.CommandButton DAItemDialogCommand 
      Caption         =   "_DAItemDialog"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   3015
   End
   Begin VB.CommandButton ComputerBrowserDialogCommand 
      Caption         =   "_ComputerBrowserDialog"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   3015
   End
   Begin VB.CommandButton DataAccess_EasyDAClientCommand 
      Caption         =   "DataAccess._EasyDAClient"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
   Begin VB.CommandButton AlarmsAndEvents_EasyAEClientCommand 
      Caption         =   "AlarmsAndEvents._EasyAEClient"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Rem
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Private Sub AlarmsAndEvents_EasyAEClientCommand_Click()
    AlarmsAndEvents_EasyAEClientForm.Show vbModal
End Sub

Private Sub ComputerBrowserDialogCommand_Click()
    ComputerBrowserDialogForm.Show vbModal
End Sub

Private Sub DataAccess_EasyDAClientCommand_Click()
    DataAccess_EasyDAClientForm.Show vbModal
End Sub

Private Sub DataAccessXmlCommand_Click()
    DataAccessXmlForm.Show vbModal
End Sub

Private Sub DAItemDialogCommand_Click()
  DAItemDialogForm.Show vbModal
End Sub

Private Sub OpcBrowseDialogCommand_Click()
  OpcBrowseDialogForm.Show vbModal
End Sub

Private Sub OpcServerDialogCommand_Click()
  OpcServerDialogForm.Show vbModal
End Sub
