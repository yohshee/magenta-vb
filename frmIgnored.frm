VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIgnored 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ignored Chatters"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3870
   Icon            =   "frmIgnored.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvwIgnored 
      Height          =   3015
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "IP"
         Text            =   "Address"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "name"
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "connected"
         Text            =   "Connected"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdUnignore 
      Caption         =   "&Unignore"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   3120
      Width           =   1215
   End
End
Attribute VB_Name = "frmIgnored"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdUnignore_Click()
Dim szIP As String

On Error GoTo err_handle:

If Not lvwIgnored.SelectedItem Is Nothing Then
    szIP = lvwIgnored.SelectedItem.Key
    ' Remove it from the list...
    lvwIgnored.ListItems.Remove szIP
    ' ...then remove it from the master ignored list.
    frmChat.RemoveIgnored szIP
Else
    MsgBox "You must select a user from the list.", vbExclamation
End If

Exit Sub
err_handle:
    ReportCrash Err, "frmIgnored", "cmdUnignore_Click", V(szIP, "szIP")
End Sub
