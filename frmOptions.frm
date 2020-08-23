VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Magenta Options"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkShowIgnore 
      Caption         =   "Display a message when someone else ignores you"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   4335
   End
   Begin VB.CheckBox chkTwoSide 
      Caption         =   "Two-Sided Ignore"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   3135
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CheckBox chkNotify 
      Caption         =   "Notify about events when minimized"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   3015
   End
   Begin VB.CheckBox chkMutIgnore 
      Caption         =   "Mutual Ignore"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Frame fraHost 
      Caption         =   "Host Information"
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   720
         Width           =   3855
      End
      Begin VB.TextBox txtHostname 
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label lblPort 
         Caption         =   "Port:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblHost 
         Caption         =   "Hostname:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbDirty As Boolean  ' Dirty flag set when the host or port are changed.

Private Sub cmdCancel_Click()
' Do nothing...
Unload Me
End Sub

Private Sub cmdOK_Click()

On Error GoTo err_handle:
' Okay, we just write the new settings into the registry, and unload.
SaveSetting App.Title, "Settings", "MutualIgnore", CStr(chkMutIgnore.Value)
SaveSetting App.Title, "Settings", "Notify", CStr(chkNotify.Value)
SaveSetting App.Title, "Settings", "TwoSidedIgnore", CStr(chkTwoSide.Value)
SaveSetting App.Title, "Settings", "ShowIgnore", CStr(chkShowIgnore.Value)
SaveSetting App.Title, "Settings", "Host", txtHostname.Text
SaveSetting App.Title, "Settings", "Port", txtPort.Text

If mbDirty Then
    MsgBox "Changes to the hostname or port settings will not take effect" & vbCrLf & _
        "until you disconnect and reconnect Magenta.", vbInformation
End If

Unload Me
Exit Sub ' Redundant, probably.

err_handle:
    ReportCrash Err, "frmOptions", "cmdOK_Click", "no variables"
End Sub

Private Sub Form_Load()

On Error GoTo err_handle:

' First, we aren't using the hostname information just yet,
' so we take proper precautions and lock those down.
'With txtHostname
'    .Text = mszHostname
'    .Enabled = False
'    .BackColor = vbButtonFace
'End With
'
'With txtPort
'    .Text = CStr(REMOTE_PORT)
'    .Enabled = False
'    .BackColor = vbButtonFace
'End With

mbDirty = False

' Read the old settings out of the registry to load the dialog.
txtHostname.Text = GetSetting(App.Title, "Settings", "Host", REMOTE_HOST)
txtPort.Text = GetSetting(App.Title, "Settings", "Port", CStr(REMOTE_PORT))
chkMutIgnore.Value = CInt(GetSetting(App.Title, "Settings", "MutualIgnore", Unchecked))
chkNotify.Value = CInt(GetSetting(App.Title, "Settings", "Notify", Unchecked))
chkTwoSide.Value = CInt(GetSetting(App.Title, "Settings", "TwoSidedIgnore", Checked))
chkShowIgnore.Value = CInt(GetSetting(App.Title, "Settings", "ShowIgnore", Checked))
Exit Sub

err_handle:
    ReportCrash Err, "frmOptions", "Form_Load", "no variables"
End Sub

Private Sub txtHostname_Change()
mbDirty = True
End Sub

Private Sub txtPort_Change()
mbDirty = True
End Sub
