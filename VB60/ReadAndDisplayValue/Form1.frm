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
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   4455
   End
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

    ' Read item value and display it
    Me.Text1 = Client.ReadItemValue("", "OPCLabs.KitServer.2", "Demo.Single")
End Sub
