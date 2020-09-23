VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPrettyPrintTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "XML Pretty - Print"
   ClientHeight    =   10605
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11070
   Icon            =   "frmPrettyPrintTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10605
   ScaleWidth      =   11070
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   615
      Top             =   -225
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Output: Indented XML String"
      Height          =   4770
      Left            =   240
      TabIndex        =   3
      Top             =   5580
      Width           =   10590
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   4320
         Left            =   210
         TabIndex        =   4
         Top             =   285
         Width           =   10170
         _ExtentX        =   17939
         _ExtentY        =   7620
         _Version        =   393217
         ScrollBars      =   3
         RightMargin     =   24000
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmPrettyPrintTest.frx":030A
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Input: Valid XML String"
      Height          =   5160
      Left            =   240
      TabIndex        =   0
      Top             =   225
      Width           =   10590
      Begin VB.CommandButton btnXMLPPrint 
         Caption         =   "Execute"
         Height          =   345
         Left            =   9195
         TabIndex        =   2
         Top             =   4680
         Width           =   1170
      End
      Begin VB.TextBox Text1 
         Height          =   4260
         Left            =   225
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Text            =   "frmPrettyPrintTest.frx":038C
         Top             =   300
         Width           =   10140
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSave 
         Caption         =   "Save Output To:"
         Begin VB.Menu mnuSaveTextFile 
            Caption         =   "Text File"
         End
         Begin VB.Menu mnuSaveXMLFile 
            Caption         =   "XML File"
         End
      End
      Begin VB.Menu separator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmPrettyPrintTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub btnXMLPPrint_Click()

Dim oXMLPP As New XMLPrettyPrint
Dim lResult As Long
Dim sXMLOutPut As String
Dim xmlDOMDoc As New DOMDocument

On Error GoTo Errorhandler

lResult = oXMLPP.XMLPPrint(sXMLOutPut, Text1.Text)

    If lResult < 0 Then
        MsgBox "Error returned: " & lResult
        GoTo CleanUpBlock
    Else
        RichTextBox1.Text = sXMLOutPut
        GoTo CleanUpBlock
    End If

Exit Sub

Errorhandler:
    MsgBox Err.Number & ":" & Err.Description, vbCritical
    
CleanUpBlock:

    Set oXMLPP = Nothing
    Set xmlDOMDoc = Nothing
    
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuSaveTextFile_Click()

Dim oFS As New FileSystemObject
Dim oFile As TextStream

On Error GoTo Errorhandler
    
    If Len(RichTextBox1.Text) = 0 Then
        MsgBox "No Output to save", vbCritical
        GoTo CleanUpBlock
    End If

    With CommonDialog1
        .Filter = "Text (*.txt)|*.txt"
        .DialogTitle = "Save Output to: "
        .Orientation = cdlLandscape
        .Action = 2
    End With

    If Len(CommonDialog1.FileName) = 0 Then
        MsgBox "Not a valid file name", vbCritical
        GoTo CleanUpBlock
    End If

    Set oFile = oFS.CreateTextFile(CommonDialog1.FileName, True)
    oFile.Write RichTextBox1.Text
    MsgBox "Output saved to:" & oFS.GetParentFolderName(CommonDialog1.FileName), vbInformation

Exit Sub
Errorhandler:
    MsgBox Err.Number & ":" & Err.Description, vbCritical
    
CleanUpBlock:

    Set oFS = Nothing
    Set oFile = Nothing

End Sub

Private Sub mnuSaveXMLFile_Click()

Dim xmlDOMDoc As New DOMDocument
Dim bWellformed As Boolean
Dim lFileExtension As Long
Dim sTempFilePath As String

On Error GoTo Errorhandler
    
     If Len(RichTextBox1.Text) = 0 Then
        MsgBox "No Output to save", vbCritical
        GoTo CleanUpBlock
    End If

    bWellformed = xmlDOMDoc.loadXML(RichTextBox1.Text)
    
    If bWellformed = False Then
        MsgBox "Output xml string not well formed", vbCritical
        GoTo CleanUpBlock
    End If
    
    With CommonDialog1
        .Filter = "XML (*.xml)|*.xml"
        .DialogTitle = "Save Output to: "
        .Orientation = cdlLandscape
        .Action = 2
    End With

    If Len(CommonDialog1.FileName) = 0 Then
        MsgBox "Not a valid file name", vbCritical
        GoTo CleanUpBlock
    End If
          
    xmlDOMDoc.save CommonDialog1.FileName
    MsgBox "Output as :" & CommonDialog1.FileName, vbInformation

Exit Sub

Errorhandler:
    MsgBox Err.Number & ":" & Err.Description, vbCritical

CleanUpBlock:

    Set xmlDOMDoc = Nothing
    
End Sub
