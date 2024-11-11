VERSION 5.00
Begin VB.Form MainForm 
   Caption         =   "Main"
   ClientHeight    =   8880
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   8880
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton UAHostAndEndpointCommand 
      Caption         =   "_UAHostAndEndpointDialog"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   8380
      Width           =   3015
   End
   Begin VB.CommandButton UAEndpointDialogCommand 
      Caption         =   "_UAEndpointDialog"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   7900
      Width           =   3015
   End
   Begin VB.CommandButton UADataDialogCommand 
      Caption         =   "_UADataDialog"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   7420
      Width           =   3015
   End
   Begin VB.CommandButton UABrowseDialogCommand 
      Caption         =   "_UABrowseDialog"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   6940
      Width           =   3015
   End
   Begin VB.CommandButton ComputerBrowserDialogCommand 
      Caption         =   "_ComputerBrowserDialog"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   6460
      Width           =   3015
   End
   Begin VB.CommandButton PubSubCommand 
      Caption         =   "PubSub"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   5400
      Width           =   3015
   End
   Begin VB.CommandButton LicensingCommand 
      Caption         =   "Licensing"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   4920
      Width           =   3015
   End
   Begin VB.CommandButton GdsCommand 
      Caption         =   "Gds"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   4440
      Width           =   3015
   End
   Begin VB.CommandButton ComplexDataCommand 
      Caption         =   "ComplexData"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3960
      Width           =   3015
   End
   Begin VB.CommandButton ApplicationCommand 
      Caption         =   "Application"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   3015
   End
   Begin VB.CommandButton AlarmsAndConditionsCommand 
      Caption         =   "AlarmsAndConditions"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   3015
   End
   Begin VB.CommandButton UANodeIdCommand 
      Caption         =   "_UANodeId"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   3015
   End
   Begin VB.CommandButton UAIndexRangeListCommand 
      Caption         =   "_UAIndexRangeList"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   3015
   End
   Begin VB.CommandButton UABrowsePathParserCommand 
      Caption         =   "_UABrowsePathParser"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   3015
   End
   Begin VB.CommandButton UAApplicationManifestCommand 
      Caption         =   "_UAApplicationManifest"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   3015
   End
   Begin VB.CommandButton EasyUAClientManagementCommand 
      Caption         =   "_EasyUAClientManagement"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
   Begin VB.CommandButton EasyUAClientCommand 
      Caption         =   "_EasyUAClient"
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

REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Private Sub ComputerBrowserDialogCommand_Click()
    ComputerBrowserDialogForm.Show vbModal
End Sub

Private Sub EasyUAClientCommand_Click()
    EasyUAClientForm.Show vbModal
End Sub

Private Sub EasyUAClientManagementCommand_Click()
    EasyUAClientManagementForm.Show vbModal
End Sub

Private Sub EasyUASubscriberCommand_Click()
  EasyUASubscriberForm.Show vbModal
End Sub

Private Sub UABrowseDialogCommand_Click()
  UABrowseDialogForm.Show vbModal
End Sub

Private Sub UADataDialogCommand_Click()
  UADataDialogForm.Show vbModal
End Sub

Private Sub UAEndpointDialogCommand_Click()
  UAEndpointDialogForm.Show vbModal
End Sub

Private Sub UAHostAndEndpointCommand_Click()
  UAHostAndEndpointDialogForm.Show vbModal
End Sub

Private Sub UAReadOnlyPubSubConfigurationCommand_Click()
 UAReadOnlyPubSubConfigurationForm.Show vbModal
End Sub

Private Sub UAApplicationManifestCommand_Click()
  UAApplicationManifestForm.Show vbModal
End Sub

Private Sub UABrowsePathParserCommand_Click()
  UABrowsePathParserForm.Show vbModal
End Sub

Private Sub UAIndexRangeListCommand_Click()
  UAIndexRangeListForm.Show vbModal
End Sub

Private Sub UANodeIdCommand_Click()
  UANodeIdForm.Show vbModal
End Sub

Private Sub AlarmsAndConditionsCommand_Click()
  AlarmsAndConditionsForm.Show vbModal
End Sub

Private Sub ApplicationCommand_Click()
  ApplicationForm.Show vbModal
End Sub

Private Sub ComplexDataCommand_Click()
  ComplexDataForm.Show vbModal
End Sub

Private Sub GdsCommand_Click()
  GdsForm.Show vbModal
End Sub

Private Sub LicensingCommand_Click()
  LicensingForm.Show vbModal
End Sub

Private Sub PubSubCommand_Click()
  PubSubForm.Show vbModal
End Sub
