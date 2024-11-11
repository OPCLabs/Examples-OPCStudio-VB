VERSION 5.00
Begin VB.Form UABrowsePathParserForm 
   Caption         =   "UABrowsePathParser"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton TryParseRelative_Main_Command 
      Caption         =   "TryParseRelative.Main"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   3015
   End
   Begin VB.CommandButton TryParse_Main_Command 
      Caption         =   "TryParse.Main"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   3015
   End
   Begin VB.CommandButton ParseRelative_Main_Command 
      Caption         =   "ParseRelative.Main"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   3015
   End
   Begin VB.CommandButton Parse_Main_Command 
      Caption         =   "Parse.Main"
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
Attribute VB_Name = "UABrowsePathParserForm"
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

REM #region Example Parse.Main
REM Parses an absolute  OPC-UA browse path and displays its starting node and elements.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Private Sub Parse_Main_Command_Click()
    OutputText = ""
    
    Dim BrowsePathParser As New UABrowsePathParser

    On Error Resume Next
    Dim browsePath As UABrowsePath
    Set browsePath = BrowsePathParser.Parse("[ObjectsFolder]/Data/Static/UserScalar")
    If Err.Number <> 0 Then
        OutputText = OutputText & "*** Failure: " & Err.Source & ": " & Err.Description & vbCrLf
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Display results
    OutputText = OutputText & "StartingNodeId: " & browsePath.StartingNodeId & vbCrLf
    
    OutputText = OutputText & "Elements:" & vbCrLf
    Dim BrowsePathElement: For Each BrowsePathElement In browsePath.Elements
        OutputText = OutputText & BrowsePathElement & vbCrLf
    Next

    ' Example output:
    'StartingNodeId: ObjectsFolder
    'Elements:
    '/Data
    '/Static
    '/UserScalar
End Sub

REM #endregion Example Parse.Main

REM #region Example ParseRelative.Main
REM Parses a relative OPC-UA browse path and displays its elements.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Private Sub ParseRelative_Main_Command_Click()
    OutputText = ""
    Dim BrowsePathParser As New UABrowsePathParser

    On Error Resume Next
    Dim BrowsePathElements As UABrowsePathElementCollection
    Set BrowsePathElements = BrowsePathParser.ParseRelative("/Data.Dynamic.Scalar.CycleComplete")
    If Err.Number <> 0 Then
        OutputText = OutputText & "*** Failure: " & Err.Source & ": " & Err.Description & vbCrLf
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Display results
    Dim BrowsePathElement: For Each BrowsePathElement In BrowsePathElements
        OutputText = OutputText & BrowsePathElement & vbCrLf
    Next

    ' Example output:
    '/Data
    '.Dynamic
    '.Scalar
    '.CycleComplete
End Sub
REM #endregion Example ParseRelative.Main

REM #region Example TryParse.Main
REM Attempts to parses an absolute  OPC-UA browse path and displays its starting node and elements.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Private Sub TryParse_Main_Command_Click()
    OutputText = ""
    
    Dim BrowsePathParser As New UABrowsePathParser

    Dim stringParsingError As stringParsingError
    Dim browsePath As Variant
    Set stringParsingError = BrowsePathParser.TryParse("[ObjectsFolder]/Data/Static/UserScalar", browsePath)
    
    ' Display results
    If Not stringParsingError Is Nothing Then
        OutputText = OutputText & "*** Error: " & stringParsingError & vbCrLf
        Exit Sub
    End If
    
    OutputText = OutputText & "StartingNodeId: " & browsePath.StartingNodeId & vbCrLf
    
    OutputText = OutputText & "Elements:" & vbCrLf
    Dim BrowsePathElement: For Each BrowsePathElement In browsePath.Elements
        OutputText = OutputText & BrowsePathElement & vbCrLf
    Next

    ' Example output:
    'StartingNodeId: ObjectsFolder
    'Elements:
    '/Data
    '/Static
    '/UserScalar
End Sub

REM #endregion Example TryParse.Main

REM #region Example TryParseRelative.Main
REM Attempts to parse a relative OPC-UA browse path and displays its elements.
REM
REM Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
REM OPC client and subscriber examples in Visual Basic on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VB .
REM Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
REM a commercial license in order to use Online Forums, and we reply to every post.

Private Sub TryParseRelative_Main_Command_Click()
    OutputText = ""
    
    Dim BrowsePathElements As New UABrowsePathElementCollection
    
    Dim BrowsePathParser As New UABrowsePathParser

    Dim stringParsingError As stringParsingError
    Set stringParsingError = BrowsePathParser.TryParseRelative("/Data.Dynamic.Scalar.CycleComplete", BrowsePathElements)
    
    ' Display results
    If Not stringParsingError Is Nothing Then
        OutputText = OutputText & "*** Error: " & stringParsingError & vbCrLf
        Exit Sub
    End If
    
    Dim BrowsePathElement: For Each BrowsePathElement In BrowsePathElements
        OutputText = OutputText & BrowsePathElement & vbCrLf
    Next

    ' Example output:
    '/Data
    '.Dynamic
    '.Scalar
    '.CycleComplete
End Sub
REM #endregion Example TryParseRelative.Main

