VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmConsole 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Magenta - Chat Console"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   Icon            =   "frmConsole.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   5910
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   2760
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save..."
      Height          =   375
      Left            =   3904
      TabIndex        =   3
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   375
      Left            =   2112
      TabIndex        =   2
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   312
      TabIndex        =   1
      Top             =   2520
      Width           =   1695
   End
   Begin RichTextLib.RichTextBox rtfConsole 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4260
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      OLEDragMode     =   0
      OLEDropMode     =   1
      TextRTF         =   $"frmConsole.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ConsolePrint(ByVal szMessage As String)
Dim szEntry As String

On Error GoTo err_handle:

' Timestamp the entry, and add it.
szEntry = "[" & Now & "] " & szMessage
rtfConsole.Text = rtfConsole.Text & vbCrLf & vbCrLf & szEntry
rtfConsole.SelStart = Len(rtfConsole.Text) + 1
Exit Sub

err_handle:
    ReportCrash Err, "frmConsole", "ConsolePrint", V(szMessage, "szMessage"), _
        V(szEntry, "szEntry")
End Sub

Private Sub cmdClear_Click()
' Simple enough...just print where it left off.
rtfConsole.Text = "[Magenta Console restarted at " & Now & "]"
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Dim szFilename As String
Dim hFile As Integer

On Error GoTo err_handler:

' First, prep and show the dialog..
With dlgOpen
    .DialogTitle = "Save Console Log"
    .Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
    .Filter = "Log Files (*.log)|*.log|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    .ShowSave
    szFilename = .FileName
End With

' Then save it up...
hFile = FreeFile
Open szFilename For Output As #hFile

Print #hFile, "[Magenta Console Log saved at " & Now & "]"
Print #hFile,
Print #hFile, rtfConsole.Text

Close #hFile

Exit Sub
err_handler:
If Err.Number = cdlCancel Then
    ' Do nothing
Else
    ReportCrash Err, "frmConsole", "cmdSave_Click", V(szFilename, "szFilename"), _
        V(hFile, "hFile")
End If
End Sub

Private Sub Form_Load()
rtfConsole.Text = "[Magenta Console started at " & Now & "]"
End Sub
