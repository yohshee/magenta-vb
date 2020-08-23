VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPrivate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Private Messaging"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   Icon            =   "frmPrivate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   2160
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save..."
      Height          =   375
      Left            =   3068
      TabIndex        =   4
      Top             =   3480
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox txtPanel 
      Height          =   3015
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   5318
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      OLEDragMode     =   0
      OLEDropMode     =   1
      TextRTF         =   $"frmPrivate.frx":030A
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   1748
      TabIndex        =   2
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Default         =   -1  'True
      Height          =   375
      Left            =   428
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txtMessage 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   3120
      Width           =   4695
   End
End
Attribute VB_Name = "frmPrivate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mReceiver As String

Private Sub cmdClose_Click()
' Not much to do, since it's not that closely associated.
Unload Me
End Sub

Private Sub cmdSave_Click()
Dim szFilename As String
Dim hFile As Integer

On Error GoTo err_handler:

' First, prep and show the dialog..
With dlgOpen
    .DialogTitle = "Save Private Chat Log"
    .Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
    .Filter = "Log Files (*.log)|*.log|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    .ShowSave
    szFilename = .FileName
End With

' Then save it up...
hFile = FreeFile
Open szFilename For Output As #hFile

Print #hFile, "[Magenta Private Chat Log (" & mReceiver & ") saved at " & Now & "]"
Print #hFile,
Print #hFile, txtPanel.Text

Close #hFile

Exit Sub
err_handler:
If Err.Number = cdlCancel Then
    ' Do nothing
Else
    ReportCrash Err, "frmPrivate", "cmdSave_Click", V(szFilename, "szFilename"), _
        V(hFile, "hFile")
End If
End Sub

Private Sub cmdSend_Click()

On Error GoTo err_handle:

' Simple enough, just invoke the method on the chat form
' that sends privately.
frmChat.SendPrivate mReceiver, txtMessage.Text, False

With txtPanel
    .SelStart = Len(txtPanel.Text) + 1
    .SelLength = 0
    .SelColor = vbMagenta
    .SelText = vbCrLf & "[" & Time$ & "] " & frmChat.txtName.Text & "> " & txtMessage.Text
    .SelStart = Len(txtPanel.Text) + 1
End With

txtMessage.Text = ""
Exit Sub
err_handle:
    ReportCrash Err, "frmPrivate", "cmdSend_Click", "no variables"
End Sub

' Properties for setting up the private chat window
Public Property Get Receiver() As String
Receiver = mReceiver
End Property

Public Property Let Receiver(ByVal szNewValue As String)
mReceiver = szNewValue
Me.Caption = "Private Chat - " & mReceiver
End Property

' Prints a message from the other end.
Public Sub PrintMessage(szMessage As String)
On Error GoTo crash

With txtPanel
    .SelStart = Len(txtPanel.Text) + 1
    .SelLength = 0
    .SelColor = vbCyan
    .SelText = vbCrLf & "[" & Time$ & "] " & mReceiver & "> " & szMessage
End With

Exit Sub
crash:
    ReportCrash Err, "frmPrivate", "PrintMessage", "no variables"
End Sub

Private Sub Form_Unload(Cancel As Integer)
' We need to dissociate this form from the chat form
' when we go.
frmChat.RemovePrivateChat mReceiver
End Sub
