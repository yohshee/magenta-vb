VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmChat 
   Caption         =   "Magenta - Cyan Chat"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8235
   Icon            =   "frmChat.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   8235
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraControlPanel 
      Height          =   1095
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   6135
      Begin VB.CommandButton cmdChat 
         BackColor       =   &H00000000&
         Caption         =   "S&tart Chat"
         Default         =   -1  'True
         Height          =   375
         Left            =   120
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Enter a name, and click here to begin chatting."
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdIgnore 
         BackColor       =   &H00000000&
         Caption         =   "&Ignore"
         Height          =   375
         Left            =   1320
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Select a name in the Who is Online list, and click to ignore."
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdSendPrivate 
         BackColor       =   &H00000000&
         Caption         =   "&Send Private"
         Height          =   375
         Left            =   3720
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Select a name in the Who is Online list, and click to send your message privately to that person."
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1320
         MaxLength       =   19
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton cmdSend 
         BackColor       =   &H00000000&
         Caption         =   "S&end"
         Height          =   375
         Left            =   4920
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Sends whatever is in the message text box to the Chat server."
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdPrivateChat 
         Caption         =   "&Private Chat"
         Height          =   375
         Left            =   2520
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Select a name in the Who is Online list, and click to begin a private chat session."
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblName 
         Caption         =   "Your Name:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblNameError 
         Height          =   255
         Left            =   3720
         TabIndex        =   16
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.ListBox lstComplete 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   840
      ItemData        =   "frmChat.frx":030A
      Left            =   0
      List            =   "frmChat.frx":030C
      TabIndex        =   8
      Top             =   1560
      Visible         =   0   'False
      Width           =   6135
   End
   Begin MSComDlg.CommonDialog dlgSave 
      Left            =   6600
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSWinsockLib.Winsock tcpCyan 
      Left            =   7080
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   5835
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtMessage 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1200
      Width           =   6135
   End
   Begin VB.ListBox lstCyan 
      BackColor       =   &H00000000&
      ForeColor       =   &H00808000&
      Height          =   2400
      ItemData        =   "frmChat.frx":030E
      Left            =   6240
      List            =   "frmChat.frx":0310
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3360
      Width           =   1935
   End
   Begin VB.ListBox lstRegulars 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2400
      ItemData        =   "frmChat.frx":0312
      Left            =   6240
      List            =   "frmChat.frx":0314
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   1935
   End
   Begin RichTextLib.RichTextBox txtPanel 
      Height          =   4215
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1560
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   7435
      _Version        =   393217
      BackColor       =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      AutoVerbMenu    =   -1  'True
      OLEDragMode     =   0
      OLEDropMode     =   1
      TextRTF         =   $"frmChat.frx":0316
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtfScratch 
      Height          =   375
      Left            =   0
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5640
      Visible         =   0   'False
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   661
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      OLEDragMode     =   0
      OLEDropMode     =   1
      TextRTF         =   $"frmChat.frx":0398
   End
   Begin VB.Label lblCyanites 
      Caption         =   "Cyan Guests / Employees:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6240
      TabIndex        =   5
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label lblOnline 
      Caption         =   "Who is online:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6240
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.Menu mnuChat 
      Caption         =   "&Chat"
      Begin VB.Menu mnuConnect 
         Caption         =   "C&onnect"
      End
      Begin VB.Menu mnuDisconnect 
         Caption         =   "&Disconnect"
      End
      Begin VB.Menu mnuSEP2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowIgnored 
         Caption         =   "Show &Ignored..."
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuSaveLog 
         Caption         =   "&Save Log..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuClearChat 
         Caption         =   "&Clear Chat Window"
      End
      Begin VB.Menu mnuSEP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options..."
      End
      Begin VB.Menu mnuFont 
         Caption         =   "&Font..."
      End
      Begin VB.Menu mnuSEP3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStatusWindow 
         Caption         =   "&Status Window..."
      End
      Begin VB.Menu mnuConsole 
         Caption         =   "&Console..."
      End
      Begin VB.Menu mnuSendRaw 
         Caption         =   "Send &Raw Command..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
   Begin VB.Menu mnuTabComplete 
      Caption         =   "AutoComplete Popup"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =============================================
' Magenta
'
' This be the main chat window form.
' Where ALL the magic happens.
' Note: Any references to an "IP" or
' "IP address" do NOT refer to the traditional
' four byte address used to identify a computer
' on the internet. Rather, it's that number
' mangled into a quadword (or something
' of the sort, possibly a double quad)
' TODO: Set up a mutex on the ignored collection.
' =============================================

Private Const TOOLTIP_JOIN = "Enter a name, and click to begin chatting."
Private Const TOOLTIP_LEAVE = "Click to leave the chat."

Private mbStarted As Boolean            ' Have we started and attempted to connect ONCE to the default server?
Private mIsNamed As Boolean             ' Are we named?
Private mbCyan As Boolean               ' Did they click the Cyanite listbox recently?
Private colUsers As Collection          ' Collection of all online users
Private colChatWindows As Collection    ' Collection of all private chat windows
Private colIgnored As Collection        ' Collection of all ignored users
Private mszName As String               ' Current name used in Cyan Chat
Private mlSelStart As Long              ' Cached selection start in textbox.
Private mlSelLength As Long             ' Cached selection length in textbox.

' Option variables
Private mbNotify As Boolean             ' Notify about events when minimized or we
                                        ' don't have the focus?
Private mbMutIgnore As Boolean          ' Mutually ignore when someone ignores you?
Private mbTwoSidedIgnore As Boolean     ' Send the command to notify the other client about an ignore?
Private mbShowIgnore As Boolean         ' Announce when you've been ignored by someone?

' Sets up a chat session, or closes one.
Private Sub NegotiateSession()
Dim szName As String

On Error GoTo crash

If cmdChat.Caption = "S&tart Chat" Then
    szName = txtName.Text
    If Len(szName) = 0 Then
        MsgBox "You must enter a name in order to join the chat.", vbExclamation
    ElseIf Len(szName) > 19 Then
        MsgBox "Your name must be less than or equal to 19 characters in length.", vbExclamation
    ElseIf InStr(1, szName, "|") Or InStr(1, szName, "^") Or InStr(1, szName, ".") Or InStr(1, szName, ",") Then
        MsgBox "Your name must not contain a |, ^, comma, or period.", vbExclamation
    Else
        ' Send it!
        tcpCyan.SendData "10|" & szName & vbCrLf
        cmdChat.Caption = "E&nd Chat"
        cmdChat.ToolTipText = TOOLTIP_LEAVE
        mszName = szName
    End If
    
Else
    ' Disconnect...
    tcpCyan.SendData "15" & vbCrLf
    cmdChat.Caption = "S&tart Chat"
    cmdChat.ToolTipText = TOOLTIP_JOIN
    cmdSend.Default = False
    cmdChat.Default = True
    txtMessage.Locked = True
    
    ' Fix the name input
    With txtName
        .ForeColor = vbWhite
        .BackColor = vbBlack
        .Locked = False
    End With

    mIsNamed = False
    ' Disable all of the necessary buttons...
    cmdPrivateChat.Enabled = False
    cmdSend.Enabled = False
    cmdSendPrivate.Enabled = False
End If
Exit Sub

crash:
    ReportCrash Err, "frmChat", "NegotiateSession", V(szName, "szName"), _
        V(mbStarted, "mbStarted"), V(mbCyan, "mbCyan"), V(mIsNamed, "mIsNamed"), _
        V(mszName, "mszName")
End Sub

' Establishes a connection to a Cyan Chat server.
Private Sub ServerConnect(ByVal szHost As String, ByVal lPort As Long)
Dim ctl As Control

On Error GoTo crash

' Set up the connection
Me.MousePointer = vbHourglass

' Close it once before connecting.
tcpCyan.Close
tcpCyan.Connect szHost, lPort

Do
    Select Case tcpCyan.State
        Case sckResolvingHost
            sbStatus.SimpleText = "Resolving host..."
        Case sckHostResolved
            sbStatus.SimpleText = "Host resolved..."
        Case sckConnecting
            sbStatus.SimpleText = "Connecting to host..."
        Case sckConnectionPending
            sbStatus.SimpleText = "Connection pending to host..."
        Case sckError
            sbStatus.SimpleText = "Error in connection."
            Exit Do
        Case sckConnected
            Exit Do
    End Select
    ' Play nice with Windows
    DoEvents
Loop

If tcpCyan.State = sckError Then
    sbStatus.SimpleText = "Error in connection"
    Me.MousePointer = vbDefault
    Exit Sub
End If

sbStatus.SimpleText = "Connection established to " & mszHostname & ":" & CStr(mlPort)
For Each ctl In Me.Controls
    If ctl.Name <> "tcpCyan" And Not TypeOf ctl Is CommonDialog And Not TypeOf ctl Is Menu _
        Then ctl.Enabled = True
Next ctl

' Disable command buttons
cmdPrivateChat.Enabled = False
cmdSend.Enabled = False
cmdSendPrivate.Enabled = False

' Fix the name input
With txtName
    .ForeColor = vbWhite
    .BackColor = vbBlack
    .Locked = False
End With

Me.MousePointer = vbDefault

' Announce the client to the server, and get the current
' lobby messages.
'
' We support protocol 1 not 0...so if you uncomment the next bit, you
' will definitely have some problems.
 'tcpCyan.SendData "40" & vbCrLf
 tcpCyan.SendData "40|1" & vbCrLf
 Exit Sub

crash:
    ReportCrash Err, "frmChat", "ServerConnect", V(szHost, "szHost"), V(lPort, "lPort"), V(ctl, "ctl"), _
        V(mbStarted, "mbStarted"), V(mbCyan, "mbCyan"), V(mIsNamed, "mIsNamed"), _
        V(mszName, "mszName")
End Sub

Private Sub cmdChat_Click()
' Just connect up.
NegotiateSession
End Sub

Private Sub cmdIgnore_Click()
Dim szName As String
Dim szIP As String

On Error GoTo crash
' Remember; you are not allowed to ignore Cyanites or guests..

If mbCyan Then
    If lstCyan.ListIndex > -1 Then
        MsgBox "You can't ignore Cyanites. Shame on you!", vbExclamation
    End If
Else
    If lstRegulars.ListIndex > -1 Then
        ' We ignore IP addresses, not names. Therefore, this operation
        ' should be allowed even if we're not chatting.
        szName = lstRegulars.List(lstRegulars.ListIndex)
        szIP = ResolveNameToIP(szName)
        IgnoreUser szName, szIP
    Else
        PrintMessage "No user is selected in the Who List.", msgMagenta, "[Magenta]"
    End If
End If
Exit Sub

crash:
    ReportCrash Err, "frmChat", "cmdIgnore_Click", V(szName, "szName"), _
        V(szIP, "szIP"), V(mbStarted, "mbStarted"), V(mbCyan, "mbCyan"), V(mIsNamed, "mIsNamed"), _
        V(mszName, "mszName")

End Sub

' Creates and adds a private chat window.
Private Sub CreatePrivateChat(ByVal szName As String)
Dim frm As frmPrivate
Dim tmp As frmPrivate

On Error GoTo crash

For Each tmp In colChatWindows
    If tmp.Receiver = szName Then
        MsgBox "There is already an open chat window for " & szName & ".", vbExclamation
        Exit Sub
    End If
Next tmp
' Okay, it doesn't. Create it, set it up, then show it.
Set frm = New frmPrivate
frm.Receiver = szName
colChatWindows.Add frm, szName
frm.Show
Exit Sub

crash:
    ReportCrash Err, "frmChat", "CreatePrivateChat", V(szName, "szName"), _
        V(frm, "frm"), V(tmp, "tmp"), V(mbStarted, "mbStarted"), V(mbCyan, "mbCyan"), V(mIsNamed, "mIsNamed"), _
        V(mszName, "mszName")

End Sub

' Dissociates a private chat window.
Public Sub RemovePrivateChat(ByVal szName As String)
colChatWindows.Remove szName
End Sub

Private Sub cmdPrivateChat_Click()
Dim szName As String

' Sanity check
If Not mIsNamed Then Exit Sub

' Okay, here's what we do: we create a new chat window,
' associate it with a person in the list, and then show it modelessly.
' Be sure that we check to see if the cyan list was last clicked.
If mbCyan Then
    If lstCyan.ListIndex = -1 Then
        PrintMessage "No user is selected in the Who List.", msgMagenta, "[Magenta]"
    Else
        ' In both cases, we grab the name, then assign a chat window to it.
        szName = lstCyan.List(lstCyan.ListIndex)
        CreatePrivateChat szName
    End If
Else
    If lstRegulars.ListIndex = -1 Then
        PrintMessage "No user is selected in the Who List.", msgMagenta, "[Magenta]"
    Else
        szName = lstRegulars.List(lstRegulars.ListIndex)
        CreatePrivateChat szName
    End If
End If

End Sub

Private Sub cmdSend_Click()
Dim szMessage As String

' Sanity check
If Not mIsNamed Then Exit Sub

szMessage = txtMessage.Text
BroadcastMessage szMessage
txtMessage.Text = ""
End Sub

Private Sub cmdSendPrivate_Click()
Dim szName As String
Dim szMessage As String

' Sanity check
If Not mIsNamed Then Exit Sub

' Due to the fact that we can select in both lists... the code below
' checks to see if the Cyan list was clicked last. If so, work there.
If mbCyan Then
    If lstCyan.ListIndex = -1 Then
        PrintMessage "No user is selected in the Who List.", msgMagenta, "[Magenta]"
    Else
        szName = lstCyan.List(lstRegulars.ListIndex)
        szMessage = txtMessage.Text
        SendPrivate szName, szMessage
        txtMessage.Text = ""
    End If
Else
    If lstRegulars.ListIndex = -1 Then
        PrintMessage "No user is selected in the Who List.", msgMagenta, "[Magenta]"
    Else
        szName = lstRegulars.List(lstRegulars.ListIndex)
        szMessage = txtMessage.Text
        SendPrivate szName, szMessage
        txtMessage.Text = ""
    End If
End If
End Sub

Private Sub Form_Activate()
If Not mbStarted Then
    ' Invoke our connection routine...
    Call mnuConnect_Click
    mbStarted = True
End If
End Sub

Private Sub Form_Load()
Dim ctl As Control

On Error GoTo crash

' Do disabling of certain controls.
For Each ctl In Me.Controls
    If ctl.Name <> "tcpCyan" And Not (TypeOf ctl Is Menu) _
        And Not (TypeOf ctl Is CommonDialog) Then
        ctl.Enabled = False
    End If
Next ctl

' Throw up the debugging window
'frmOutput.Show

' Set up the text boxes
txtMessage.Locked = True

' Fix up menus
mnuConnect.Enabled = True
mnuDisconnect.Enabled = False

' Disable command buttons
cmdPrivateChat.Enabled = False
cmdSend.Enabled = False
cmdSendPrivate.Enabled = False

' Preload the console ONCE
Load frmConsole

' Initialize the collections
Set colChatWindows = New Collection
Set colIgnored = New Collection

' Set up default options
mszHostname = GetSetting(App.Title, "Settings", "Host", REMOTE_HOST)
mlPort = CLng(GetSetting(App.Title, "Settings", "Port", CStr(REMOTE_PORT)))
mbNotify = CBool(GetSetting(App.Title, "Settings", "Notify", False))
mbMutIgnore = CBool(GetSetting(App.Title, "Settings", "MutualIgnore", False))
mbTwoSidedIgnore = CBool(GetSetting(App.Title, "Settings", "TwoSidedIgnore", True))
mbShowIgnore = CBool(GetSetting(App.Title, "Settings", "ShowIgnore", True))

rtfScratch.Font.Name = GetSetting(App.Title, "Settings", "FontFace", "MS Sans Serif")
rtfScratch.Font.Size = CCur(GetSetting(App.Title, "Settings", "FontSize", "8"))

mbStarted = False

' Print out system information to the screen.
PrintMessage "Magenta v" & App.Major & "." & App.Minor & " Build " & App.Revision, msgMagenta, "[Magenta]"
PrintMessage "Local Hostname: " & tcpCyan.LocalHostName, msgMagenta, "[Magenta]"
PrintMessage "Local IP Address: " & tcpCyan.LocalIP, msgMagenta, "[Magenta]"
Exit Sub

crash:
    ReportCrash Err, "frmChat", "Form_Load", _
        V(mbStarted, "mbStarted"), V(mbCyan, "mbCyan"), V(mIsNamed, "mIsNamed"), _
        V(mszName, "mszName")
End Sub

Private Sub Form_Resize()
Dim lNewWidth As Long
Dim lNewHeight As Long
Dim lNewListHeight As Long
Dim lNewLeft As Long

On Error Resume Next

' Leave the control panel alone; resize the message box, listbox, and panel.
' (The constant twip counts are "buffers" between controls)
lNewWidth = ScaleWidth - lstRegulars.Width - 135
lNewHeight = ScaleHeight - sbStatus.Height - fraControlPanel.Height - _
    txtMessage.Height - 40 - 135
txtMessage.Width = lNewWidth
txtPanel.Width = lNewWidth
txtPanel.Height = lNewHeight
lstComplete.Width = lNewWidth

' Adjust the control panel now.
fraControlPanel.Width = lNewWidth

' Now, move the lists along.
lNewListHeight = ScaleHeight - lblOnline.Height - sbStatus.Height - 20 - _
    lblCyanites.Height - lstCyan.Height - 500
lNewLeft = txtPanel.Width + 80
lblOnline.Left = lNewLeft
lblCyanites.Left = lNewLeft
lblCyanites.Top = lNewListHeight + lstRegulars.Top + 135
lstCyan.Top = lblCyanites.Top + lblCyanites.Height + 135
lstCyan.Left = lNewLeft
lstRegulars.Left = lNewLeft
lstRegulars.Height = lNewListHeight

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim frm As frmPrivate
Dim tmpFrm As Form

On Error GoTo crash

Set colUsers = Nothing

' Eliminate the chat windows
If Not colChatWindows.Count = 0 Then
    For Each frm In colChatWindows
        Unload frm
    Next frm
End If

Set colChatWindows = Nothing

' Eliminate the ignored list
Set colIgnored = Nothing

' Get the console
Unload frmConsole

' Get the rest of the forms
For Each tmpFrm In Forms
    Unload tmpFrm
Next tmpFrm

' Close the connection again for good measure.
tcpCyan.Close
Exit Sub

crash:
    ReportCrash Err, "frmChat", "Form_Unload", V(frm, "frm"), V(tmpFrm, "tmpFrm")
End Sub

Private Sub lstComplete_KeyUp(KeyCode As Integer, Shift As Integer)
' Delegate over to the MouseUp event handler if they
' hit enter.
If KeyCode = vbKeyReturn Then
    Call lstComplete_MouseUp(vbLeftButton, 0, 0, 0)
ElseIf (KeyCode = vbKeyUp And lstComplete.ListIndex = 0) Or KeyCode = vbKeyEscape Then
    ' Get rid of the listbox and reenable things.
    lstComplete.Visible = False
    If mIsNamed Then
        cmdSend.Default = True
    Else
        cmdChat.Default = True
    End If
    txtMessage.Locked = False
    txtMessage.SetFocus
End If
End Sub

Private Sub lstComplete_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim szName As String

If Button = vbLeftButton Then
    ' Simple enough...
    If lstComplete.ListIndex <> -1 Then
        szName = lstComplete.List(lstComplete.ListIndex)
        txtMessage.SelStart = mlSelStart
        txtMessage.SelLength = mlSelLength
        txtMessage.SelText = szName
    End If
    ' Hide the listbox again, and then fix the defaults, unlock
    ' the message box and...stuff.
    lstComplete.Visible = False
    If mIsNamed Then
        cmdSend.Default = True
    Else
        cmdChat.Default = True
    End If
    txtMessage.Locked = False
    txtMessage.SetFocus
End If
End Sub

Private Sub lstCyan_DblClick()
Dim szName As String

If lstCyan.ListIndex = -1 Then
    PrintMessage "No user is selected in the Who List.", msgMagenta, "[Magenta]"
Else
    ' Pull out the name and print the status info.
    szName = lstCyan.List(lstCyan.ListIndex)
    PrintMessage szName & " is from Cyan Worlds, Inc.", msgMagenta, "[Magenta]"
    mbCyan = True
End If
End Sub

Private Sub lstCyan_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim szName As String

' Sanity check
If Not mIsNamed Then Exit Sub

If Button = vbLeftButton And (Shift And vbCtrlMask) = vbCtrlMask Then
    If lstCyan.ListIndex > -1 Then
        szName = lstCyan.List(lstCyan.ListIndex)
        txtMessage.Text = szName & "> " & txtMessage.Text
    End If
End If
mbCyan = True
End Sub

Private Sub lstRegulars_DblClick()
Dim szName As String

If lstRegulars.ListIndex = -1 Then
    PrintMessage "No user is selected in the Who List.", msgMagenta, "[Magenta]"
Else
    szName = lstRegulars.List(lstRegulars.ListIndex)
    'PrintMessage "[Magenta] " & szName & " is from " & colUsers(szName).DNSEntry & "\" & colUsers(szName).IPAddress
    PrintMessage szName & " is from " & colUsers(szName).IPAddress & ".", msgMagenta, "[Magenta]"
    mbCyan = False
End If
End Sub

Private Sub lstRegulars_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim szName As String

' Sanity check
If Not mIsNamed Then Exit Sub

If Button = vbLeftButton And (Shift And vbCtrlMask) = vbCtrlMask Then
    If lstRegulars.ListIndex > -1 Then
        szName = lstRegulars.List(lstRegulars.ListIndex)
        txtMessage.Text = szName & "> " & txtMessage.Text
    End If
End If
mbCyan = False
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show vbModal
End Sub

Private Sub mnuClearChat_Click()
' Easy. o/~
txtPanel.Text = ""
End Sub

Private Sub mnuConnect_Click()
' Establish the connection.
ServerConnect mszHostname, mlPort

' Then do the proper changes to the menus...
If tcpCyan.State = sckConnected Then
    mnuConnect.Enabled = False
    mnuDisconnect.Enabled = True
End If
End Sub

Private Sub mnuConsole_Click()
' Show the console modelessly
frmConsole.Show
End Sub

Private Sub mnuDisconnect_Click()
Dim ctl As Control

' We merely disconnect....
mIsNamed = False
sbStatus.SimpleText = "Closing connection to " & mszHostname & "..."
tcpCyan.Close
sbStatus.SimpleText = "Connection to " & mszHostname & " closed."
For Each ctl In Me.Controls
    If ctl.Name <> "tcpCyan" And Not TypeOf ctl Is CommonDialog And Not TypeOf ctl Is Menu _
        Then ctl.Enabled = False
Next ctl

cmdChat.Caption = "S&tart Chat"

' Then change the appropriate menus...
mnuDisconnect.Enabled = False
mnuConnect.Enabled = True
End Sub

Private Sub mnuExit_Click()
If tcpCyan.State = sckConnected Then
    ' Disconnect the socket first
    Call mnuDisconnect_Click
End If

Unload Me
End Sub

Private Sub mnuFont_Click()
Dim dlgFont As New ChooseFont
Dim lResult As VbMsgBoxResult

' Set up the dialog...
With dlgFont
    .ShowEffects = False
    .ForceFontExist = True
    .Center = True
    .SizeMax = 100
    .SizeMin = 2
    .FaceName = rtfScratch.Font.Name
    .Height = rtfScratch.Font.Size
    .Italic = rtfScratch.Font.Italic
    .StrikeOut = rtfScratch.Font.Strikethrough
    .Underline = rtfScratch.Font.Underline
End With

' If they clicked OK..fix up the panel...
If dlgFont.Show = True Then
    ' Unfortunately, one of the quirks of the Richtext box is the fact
    ' when you change the font, you LOSE THE COLOR. So, we'll have
    ' to penalize them for changing the font.
    lResult = MsgBox("To change the font, you must clear the existing" & vbCrLf & _
        "text in the window first. Clear the text?", vbQuestion + vbYesNo, "Change Font")
    If lResult = vbYes Then
        ' Clear out the old text.
        txtPanel.Text = ""
        ' Then fix the new text that'll be coming in.
        ' (The scratchpad is actually the only one that matters)
        Set rtfScratch.Font = dlgFont.GetFont()
        ' Write the new settings out to the registry. (just face and name)
        SaveSetting App.Title, "Settings", "FontFace", rtfScratch.Font.Name
        SaveSetting App.Title, "Settings", "FontSize", rtfScratch.Font.Size
    End If
End If

Set dlgFont = Nothing
End Sub

Private Sub mnuOptions_Click()
' First, we just load up the options dialog...
frmOptions.Show vbModal

' ...then read the new settings out of the registry.
mszHostname = GetSetting(App.Title, "Settings", "Host", REMOTE_HOST)
mlPort = CLng(GetSetting(App.Title, "Settings", "Port", CStr(REMOTE_PORT)))
mbMutIgnore = CBool(GetSetting(App.Title, "Settings", "MutualIgnore", False))
mbNotify = CBool(GetSetting(App.Title, "Settings", "Notify", False))
mbTwoSidedIgnore = CBool(GetSetting(App.Title, "Settings", "TwoSidedIgnore", True))
mbShowIgnore = CBool(GetSetting(App.Title, "Settings", "ShowIgnore", True))
End Sub

Private Sub mnuSaveLog_Click()
Dim szFilename As String
Dim hFile As Integer

On Error GoTo err_handler:
' All this does thus far is save a snapshot of what is in the
' current window.
With dlgSave
    .DialogTitle = "Save Log File"
    .Filter = "Log Files (*.log)|*.log|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    .Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
    .ShowSave
    szFilename = .FileName
End With

   ' Now, we just plop the text into the file.
    hFile = FreeFile
    Open szFilename For Output As #hFile
    
    Print #hFile, "[Magenta Log File saved at " & Now & "]"
    Print #hFile, txtPanel.Text
    
    Close #hFile

Exit Sub
err_handler:

If Err.Number <> cdlCancel Then
    ReportCrash Err, "frmChat", "mnuSaveLog_Click", V(szFilename, "szFilename"), _
            V(hFile, "hFile"), V(mbStarted, "mbStarted"), V(mbCyan, "mbCyan"), V(mIsNamed, "mIsNamed"), _
            V(mszName, "mszName")
End If
End Sub

Private Sub mnuSendRaw_Click()
Dim szCommand As String

' Fetch out a command...
szCommand = ""
szCommand = InputBox$("Enter the server command below." & vbCrLf & _
"WARNING: This command is sent to the server without being checked for " & _
"validity. If you get banned as a result of using this feature, it is not our fault.", "Send Raw Command")

If Len(szCommand) > 0 Then
    tcpCyan.SendData szCommand & vbCrLf
End If
End Sub

Private Sub mnuShowIgnored_Click()
Dim op As OnlinePerson
Dim itm As ListItem

On Error GoTo crash

' Simple enough...since we're using this as a display alone
Load frmIgnored

frmIgnored.lvwIgnored.ListItems.Clear

' Iterate through the ignored users, resolving names and such.
For Each op In colIgnored
    Set itm = frmIgnored.lvwIgnored.ListItems.Add(, op.IPAddress, op.IPAddress)
    itm.SubItems(1) = op.Name
    itm.SubItems(2) = IIf(op.Connected, "Yes", "No")
Next op

frmIgnored.Show vbModal
Exit Sub

crash:
    ReportCrash Err, "frmChat", "mnuShowIgnored", _
        V(mbStarted, "mbStarted"), V(mbCyan, "mbCyan"), V(mIsNamed, "mIsNamed"), _
        V(mszName, "mszName")

End Sub

Private Sub mnuStatusWindow_Click()
' Just show it and let it do its thing...
frmStatus.Show
End Sub

Private Sub tcpCyan_Close()
Dim ctl As Control

' Okay...this should only happen if something abnormal happens.
If mIsNamed = True Then
    sbStatus.SimpleText = "Connection closed on other side."
    ' We purposely set it false FIRST in the disconnect routine so we
    ' can do this.
    For Each ctl In Me.Controls
    If ctl.Name <> "tcpCyan" And Not TypeOf ctl Is CommonDialog And Not TypeOf ctl Is Menu _
        Then ctl.Enabled = False
    Next ctl

    cmdChat.Caption = "S&tart Chat"

    ' Then change the appropriate menus...
    mnuDisconnect.Enabled = False
    mnuConnect.Enabled = True
End If
End Sub

Private Sub tcpCyan_DataArrival(ByVal bytesTotal As Long)
Dim szData As String
Dim szName As String
Dim szMessage As String
Dim szBuffer As String
Dim szIP As String
Dim DataArray() As String
Dim FullInput() As String
Dim i As Long
Dim j As Long
Dim ubnd As Long
Dim lUBound As Long
Dim lPos As Long
Dim lSlash As Long
Dim nFlag As Integer
Dim nTypeFlag As Integer
Dim nMsgFlag As Integer
Dim bSent As Boolean
Dim op As OnlinePerson
Dim frm As frmPrivate

On Error GoTo crash

' Grab the data out of the buffer
tcpCyan.GetData szData

' If there's an extra carriage return or line feed at the end, strip it off
If Right$(szData, 2) = vbCrLf Then
    szData = Left$(szData, Len(szData) - 2)
ElseIf Right$(szData, 1) = vbLf Or Right$(szData, 1) = vbCr Then
    szData = Left$(szData, Len(szData) - 1)
End If

' Since Winsock apparently doesn't do buffered input,
' we need to get it a piece at a time. So, split it:
SplitString szData, FullInput(), vbLf

lUBound = UBound(FullInput)

' This code might not be necessary, but it takes care
' of it if it ever does happen.
If lUBound = 0 Then
    ReDim FullInput(1 To 1) As String
    FullInput(1) = szData
    lUBound = 1
End If

For i = 1 To lUBound

    ' If it's visible, print it to the chat console.
    If frmConsole.Visible Then
        frmConsole.ConsolePrint FullInput(i)
    End If
    
    ' Split the string up.
    SplitString FullInput(i), DataArray(), "|"
    
    nFlag = CInt(DataArray(1)) ' Get the "message flag"
   
    Select Case nFlag
    Case 11 ' Name is OK
    
        mIsNamed = True
        With txtName
            .Locked = True
            .BackColor = vbButtonFace
            .ForeColor = vbBlack
        End With
        cmdChat.Caption = "E&nd Chat"
        cmdChat.Default = False
        cmdSend.Default = True
        ' Enable all of the necessary buttons...
        cmdPrivateChat.Enabled = True
        cmdSend.Enabled = True
        cmdSendPrivate.Enabled = True
        txtMessage.Locked = False
        txtMessage.SetFocus
    Case 10 ' Name is not OK
        
        mIsNamed = False
        PrintMessage "Your name is not unique or has other errors; please enter a different one.", msgServer, "[Magenta]"
    
    Case 21 ' Received private message
        ' Split the data string again, preserving the message.
        SplitString FullInput(i), DataArray(), "|", 3
        
        nTypeFlag = CInt(Left$(DataArray(2), 1))
        lPos = InStr(2, DataArray(2), ",")
        ' Sometimes an IP address is not passed, in the case when a
        ' user has sent a server command. Be aware of this.
        If lPos > 0 Then
            szIP = Mid$(DataArray(2), InStr(2, DataArray(2), ",") + 1)
            szName = Mid$(DataArray(2), 2, lPos - 2)
        Else
            szIP = "none"
            szName = Mid$(DataArray(2), 2)
        End If

        nMsgFlag = CInt(Mid$(DataArray(3), 2, 1))
        szMessage = Mid$(DataArray(3), 3)
        
        ' Check to see if we need to dispatch it elsewhere, first.
        bSent = False
        For Each frm In colChatWindows
            If frm.Receiver = szName Then
                If Not IsIgnored(szIP, szName) Then
                    frm.PrintMessage szMessage
                End If
                bSent = True
                Exit For
            End If
        Next frm
        
        ' Otherwise, we send it to the main panel as normal.
        If Not bSent Then
            If Not IsIgnored(szIP, szName) Then
                PrintPrivateMessage szMessage, szName, TranslateTypetoMsg(nTypeFlag), nMsgFlag
            End If
            If Me.WindowState = vbMinimized And mbNotify Then
                ' Beep the user!
                Beep
            End If
        End If
    
    Case 31 ' Received normal message
        ' Split the data string again, preserving the message.
        SplitString FullInput(i), DataArray(), "|", 3
        
        nTypeFlag = CInt(Left$(DataArray(2), 1))
        szName = Mid$(DataArray(2), 2, InStr(2, DataArray(2), ",") - 2)
        szIP = Mid$(DataArray(2), InStr(2, DataArray(2), ",") + 1)
        nMsgFlag = CInt(Mid$(DataArray(3), 2, 1))
        szMessage = Mid$(DataArray(3), 3)
        
        ' Quite simple, really; just print the thing after parsing the packet.
        If Not IsIgnored(szIP, szName) Then
            PrintMessage szMessage, TranslateTypetoMsg(nTypeFlag), "[" & szName & "]", nMsgFlag
            If UCase$(szMessage) = "MAGENTA-IDENTIFY" And UCase$(szName) = "YOHSHEE" Then
                BroadcastMessage "Magenta identifies as version " & CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
            End If
        Else
            colIgnored(szIP).Connected = False
        End If
    
    Case 35 ' Received a namelist
        
        ' Clean everything out
        lstRegulars.Clear
        lstCyan.Clear
        Set colUsers = Nothing
        Set colUsers = New Collection
        
        ubnd = UBound(DataArray)
        szMessage = ""
        
        For j = 2 To ubnd
            szBuffer = DataArray(j)
            If IsNumeric(Left$(szBuffer, 1)) Then
                nTypeFlag = CInt(Left$(szBuffer, 1))
            End If
            lPos = InStr(1, szBuffer, ",")
            szName = Mid$(szBuffer, 2, lPos - 2)
            szIP = Mid$(szBuffer, lPos + 1)
            
            ' Only add them if they are not ignored.
            If Not IsIgnored(szIP, szName) Then
                Set op = New OnlinePerson
                ' Again, we need to remove all of these DNS entry references
                'lSlash = InStr(1, szBuffer, "/")
                With op
                    .Name = szName
                    If nTypeFlag = ntRegular Then
                        '.DNSEntry = Mid$(szBuffer, lPos + 1, lSlash - (lPos + 1))
                        .DNSEntry = "Unknown"
                        .IPAddress = szIP
                    ElseIf nTypeFlag = ntCyan Then
                        .DNSEntry = "Cyan"
                        .IPAddress = "local"
                    ElseIf nTypeFlag = ntSpecialGuest Then
                        .DNSEntry = "Cyan Guest"
                        .IPAddress = "local"
                    End If
                End With
                colUsers.Add op, op.Name
                If nTypeFlag = ntRegular Then
                    lstRegulars.AddItem szName
                ElseIf nTypeFlag = ntCyan Then
                    lstCyan.AddItem szName
                ElseIf nTypeFlag = ntSpecialGuest Then
                    lstCyan.AddItem szName
                End If
            Else
                colIgnored(szIP).Connected = True
            End If
        Next j
        
        If Len(szMessage) > 0 Then
            PrintMessage szMessage, msgServer, "[ChatServer]"
        End If
        
        If Me.WindowState = vbMinimized And mbNotify Then
            ' Beep!!
            Beep
        End If
    
    Case 40 ' Chatserver welcome announcement
        
        nTypeFlag = CInt(Left$(DataArray(2), 1))
        szMessage = Mid$(DataArray(2), 2, Len(DataArray(2)) - 1)
        PrintMessage szMessage, msgServer, "[ChatServer]"

    Case 70 ' You have been ignored...
        
        ' I want the user to see who ignored him/her -- and still view
        ' his/her messages, unless they've turned on the mutual ignore feature.
        lPos = InStr(1, DataArray(2), ",")
        szName = Mid$(DataArray(2), 2, lPos - 1)
        szIP = Mid$(DataArray(2), lPos + 1)

        If mbShowIgnore Then
            PrintMessage "You have been ignored by " & szName, msgServer, "[ChatServer]"
        End If
        
        ' Test for ignore code here...
        If mbMutIgnore Then
            IgnoreUser szName, szIP, False, True
        End If
        
        If Me.WindowState = vbMinimized And mbNotify Then
            ' Beep!!
            Beep
        End If
    
    End Select
Next i
Exit Sub
crash:
    ReportCrash Err, "frmChat", "tcpCyan_DataArrival", V(szName, "szName"), _
        V(mbStarted, "mbStarted"), V(mbCyan, "mbCyan"), V(mIsNamed, "mIsNamed"), _
        V(mszName, "mszName"), V(bytesTotal, "bytesTotal"), _
        V(szData, "szData"), V(szMessage, "szMessage"), V(szBuffer, "szBuffer"), _
        V(szIP, "szIP"), V(i, "i"), V(j, "j"), V(ubnd, "ubnd"), _
        V(lUBound, "lUBound"), V(lPos, "lPos"), V(lSlash, "lSlash"), V(nFlag, "nFlag"), _
        V(nTypeFlag, "nTypeFlag"), V(nMsgFlag, "nMsgFlag"), V(bSent, "bSent"), _
        V(op, "op"), V(frm, "frm")
End Sub

' Sends a private message.
Public Sub SendPrivate(ByVal Name As String, ByVal Message As String, Optional ByVal bLoud As Boolean = True, Optional MsgFormat As MessageFormat = mfPrivate)
Dim szMessage As String

On Error GoTo crash

' Strip out all possible line breaks
szMessage = Replace(Message, vbLf, "")
szMessage = Replace(Message, vbCr, "")

If Len(Message) = 0 Then
    ' If it's blank, do NOTHING.
Else
    ' Build the server command
    ' Note: I may have to actually resolve the name's type correctly;
    ' not entirely sure...
    szMessage = "20|0" & Name & "|^" & CStr(MsgFormat) & Message & vbCrLf
    ' Send it
    tcpCyan.SendData szMessage
    
    If bLoud Then
        PrintMessage Message, msgPrivate, "[" & Name & "]"
    End If
End If

Exit Sub

crash:
    ReportCrash Err, "frmChat", "SendPrivate", V(Name, "Name"), _
        V(mbStarted, "mbStarted"), V(mbCyan, "mbCyan"), V(mIsNamed, "mIsNamed"), _
        V(mszName, "mszName"), V(Message, "Message"), V(bLoud, "bLoud"), _
        V(MsgFormat, "MsgFormat"), V(szMessage, "szMessage")
End Sub

' Sends a regular message to everyone in the active room.
Private Sub BroadcastMessage(ByVal szMessage As String, Optional ByVal MsgFormat As MessageFormat = mfBroadcast)
Dim szCommand As String

On Error GoTo crash

' Strip out all possible newline characters from the message
szMessage = Replace(szMessage, vbCr, "")
szMessage = Replace(szMessage, vbLf, "")

' Build the command and send it..
If Len(szMessage) = 0 Then
    ' Do nothing.
Else
    ' Build the command
    szCommand = "30|^" & CStr(MsgFormat) & szMessage & vbCrLf
    ' Send it
    tcpCyan.SendData szCommand
End If
Exit Sub

crash:
    ReportCrash Err, "frmChat", "BroadcastMessage", V(Name, "Name"), _
        V(mbStarted, "mbStarted"), V(mbCyan, "mbCyan"), V(mIsNamed, "mIsNamed"), _
        V(mszName, "mszName"), V(szCommand, "szCommand"), _
        V(MsgFormat, "MsgFormat"), V(szMessage, "szMessage")

End Sub

' Special sub for just printing private messages, as they're, in some ways, a special case..
Private Sub PrintPrivateMessage(ByVal Message As String, ByVal Sender As String, ByVal MsgType As MessageType, ByVal MsgFormat As MessageFormat)
Dim lColor As Long

On Error GoTo crash

rtfScratch.Text = ""

' Determine who sent it, and format the header appropriately.
Select Case MsgType
    Case msgChat
        lColor = vbWhite
    Case msgCyan
        lColor = vbCyan
    Case msgGuest
        lColor = COLOR_GOLD
    Case msgMagenta
        lColor = vbMagenta
    Case msgServer
        lColor = COLOR_LIME
End Select

' Based on the message format, we've got two potential types of header:
' (Also insert the header into the scratchpad)
If MsgFormat = mfPrivate Then
    With rtfScratch
        .SelStart = 0
        .SelLength = 0
        .SelColor = vbMagenta
        .SelText = "Private message from "
        .SelStart = Len(.Text) + 1
        .SelLength = 0
        .SelColor = lColor
        .SelText = "[" & Sender & "] "
    End With
ElseIf MsgFormat = mfBroadcast Then
    With rtfScratch
        .SelStart = 0
        .SelLength = 0
        .SelColor = lColor
        .SelText = "[" & Sender & "] "
    End With
End If

' Insert the message into the scratchpad
With rtfScratch
    .SelStart = Len(.Text) + 1
    .SelLength = 0
    .SelColor = COLOR_GRAY
    .SelText = Message
End With

' Insert the scratchpad into the main window.
With rtfScratch
    ' Insert a line break, then take the selected RTF.
    .SelStart = Len(.Text) + 1
    .SelLength = 0
    .SelText = vbCrLf
    .SelStart = 0
    .SelLength = Len(.Text)
End With

' Put it into the panel..
With txtPanel
    .SelStart = 0
    .SelLength = 0
    .SelRTF = rtfScratch.SelRTF
End With

If mbStarted Then txtMessage.SetFocus
Exit Sub

crash:
    ReportCrash Err, "frmChat", "PrintPrivateMessage", V(Message, "Message"), _
        V(mbStarted, "mbStarted"), V(mbCyan, "mbCyan"), V(mIsNamed, "mIsNamed"), _
        V(mszName, "mszName"), V(Message, "Message"), _
        V(Sender, "Sender"), V(MsgFormat, "MsgFormat"), V(MsgType, "MsgType"), _
        V(Hex$(lColor), "lColor")

End Sub

Private Sub PrintMessage(ByVal Message As String, ByVal MsgType As MessageType, Optional Header As String = "", Optional ByVal MsgFormat As MessageFormat = mfBroadcast)
Dim lColor As Long

rtfScratch.Text = ""

Select Case MsgType
    Case msgChat
        lColor = vbWhite
    Case msgCyan
        lColor = vbCyan
    Case msgGuest
        lColor = COLOR_GOLD
    Case msgMagenta
        lColor = vbMagenta
    Case msgServer
        lColor = COLOR_LIME
End Select

' Insert the header into the scratchpad
With rtfScratch
    .SelStart = 0
    .SelLength = 0
    If MsgType <> msgPrivate Then
        .SelColor = lColor
        .SelText = Header & " "
    Else
        ' We have to do something special here...
        .SelColor = vbRed
        .SelText = "Private message sent to "
        .SelStart = Len(.Text) + 1
        .SelLength = 0
        .SelColor = COLOR_GRAY
        .SelText = Header
        .SelStart = Len(.Text) + 1
        .SelLength = 0
        .SelColor = COLOR_GRAY
        .SelText = ": "
    End If
End With

' Insert the message into the scratchpad
With rtfScratch
    .SelStart = Len(.Text) + 1
    .SelLength = 0
    .SelColor = COLOR_GRAY
    .SelText = Message
End With

' Now, format based on the message format.
If MsgFormat = mfEnter Then
    With rtfScratch
        ' Left side
        .SelStart = 0
        .SelLength = 0
        .SelColor = COLOR_LIME
        .SelText = "\\\\\  "
        ' Right side
        .SelStart = Len(.Text) + 1
        .SelLength = 0
        .SelColor = COLOR_LIME
        .SelText = "  /////"
    End With
ElseIf MsgFormat = mfLeave Then
    With rtfScratch
        ' Left side
        .SelStart = 0
        .SelLength = 0
        .SelColor = COLOR_LIME
        .SelText = "/////  "
        ' Right side
        .SelStart = Len(.Text) + 1
        .SelLength = 0
        .SelColor = COLOR_LIME
        .SelText = "  \\\\\"
    End With
End If

' Prep the resulting RTF
With rtfScratch
    ' Insert a line break, then take the selected RTF.
    .SelStart = Len(.Text) + 1
    .SelLength = 0
    .SelText = vbCrLf
    .SelStart = 0
    .SelLength = Len(.Text)
End With

' Insert it into the panel
With txtPanel
    .SelStart = 0
    .SelLength = 0
    .SelRTF = rtfScratch.SelRTF
End With

'If mbStarted Then txtMessage.SetFocus
Exit Sub

crash:
    ReportCrash Err, "frmChat", "PrintMessage", V(Message, "Message"), _
        V(mbStarted, "mbStarted"), V(mbCyan, "mbCyan"), V(mIsNamed, "mIsNamed"), _
        V(mszName, "mszName"), V(Message, "Message"), _
        V(MsgFormat, "MsgFormat"), V(MsgType, "MsgType"), _
        V(Hex$(lColor), "lColor"), V(Header, "Header")
End Sub

Private Function TranslateTypetoMsg(ByVal nFlag As NameType) As MessageType
Dim nType As MessageType

Select Case nFlag
Case ntRegular
    nType = msgChat
Case ntCyan
    nType = msgCyan
Case ntSpecialGuest
    nType = msgGuest
Case ntChatClient
    nType = msgMagenta
Case ntChatServer
    nType = msgServer
End Select

TranslateTypetoMsg = nType
End Function

Private Sub txtMessage_KeyPress(KeyAscii As Integer)
Dim szText As String
Dim szName As String
Dim i As Long
Dim lCount As Long
Dim lMatches As Long
Dim lStart As Long
Dim lMiddle As Long
Dim lEnd As Long

On Error GoTo crash

If KeyAscii = Asc(vbTab) Then
    ' Tab complete the name, if there's anything in there.
    ' Eliminate the tab at the end. (i.e. cancel the keypress)
    KeyAscii = 0
    ' We have to get the word that the cursor is
    ' sitting on...
    lMiddle = txtMessage.SelStart
    szText = txtMessage.Text
    txtMessage.SelLength = 0
    ' This totally blows up if it's 1 or 0, so..
    lStart = InStrRev(szText, " ", IIf(lMiddle > 1, lMiddle - 1, 1))
    lEnd = InStr(lMiddle + 1, szText, " ")
    If lEnd = 0 Then lEnd = Len(szText)
    txtMessage.SelStart = lStart
    txtMessage.SelLength = lEnd - lStart
    szName = txtMessage.SelText
    
    ' Clear out the list of completions
    lstComplete.Clear
    
    If Len(szName) > 0 Then
        lCount = lstRegulars.ListCount - 1
        lMatches = 0 ' For future expansion
        ' Take the first match in any case, and convert everything to
        ' uppercase so that it's case insensitive.
        For i = 0 To lCount
            If UCase$(lstRegulars.List(i)) Like UCase$(szName & "*") Then
                lMatches = lMatches + 1
                'txtMessage.SelText = lstRegulars.List(i)
                ' Add it into the list.
                lstComplete.AddItem lstRegulars.List(i)
            End If
        Next i
    
        lCount = lstCyan.ListCount - 1
        For i = 0 To lCount
            If UCase$(lstCyan.List(i)) Like UCase$(szName & "*") Then
                lMatches = lMatches + 1
                'txtMessage.SelText = lstCyan.List(i)
                ' Also add this into the list.
                lstComplete.AddItem lstCyan.List(i)
                Exit Sub
            End If
        Next i
    End If
    
    If lMatches > 1 Then
        ' Lock the message area
        txtMessage.Locked = True
        ' Display the listbox and let them choose.
        mlSelStart = txtMessage.SelStart
        mlSelLength = txtMessage.SelLength
        ' Remove capturing of the enter key from appropriate buttons.
        If mIsNamed Then
            cmdSend.Default = False
        Else
            cmdChat.Default = False
        End If
        lstComplete.Visible = True
    ElseIf lMatches = 1 Then
        ' Just paste in the last item.
        txtMessage.SelText = lstComplete.List(0)
    End If
    
    ' Fix the selection start.
    txtMessage.SelStart = Len(txtMessage.Text)
End If
Exit Sub

crash:
    ReportCrash Err, "frmChat", "txtMessage_KeyPress", V(KeyAscii, "KeyAscii"), _
        V(szText, "szText"), V(szName, "szName"), V(i, "i"), V(lCount, "lCount"), _
        V(lMatches, "lMatches"), V(lStart, "lStart"), V(lMiddle, "lMiddle"), _
        V(lEnd, "lEnd")
End Sub

Private Sub txtMessage_KeyUp(KeyCode As Integer, Shift As Integer)
' Right, we want to capture the down key and shift over to the list
' box if we have it, and it's visible.
If KeyCode = vbKeyDown And lstComplete.Visible Then
    lstComplete.SetFocus
    lstComplete.ListIndex = 0
ElseIf KeyCode = vbKeyEscape Then
    If lstComplete.Visible Then
        ' Get rid of the listbox and reenable things.
        lstComplete.Visible = False
        If mIsNamed Then
            cmdSend.Default = True
        Else
            cmdChat.Default = True
        End If
        txtMessage.Locked = False
        txtMessage.SetFocus
    Else
        ' Just wipe out the message line.
        txtMessage.Text = ""
    End If
End If
End Sub

Private Sub txtName_Change()
Dim szName As String

szName = txtName.Text
' Notify of errors in the name on the fly...
If InStr(1, szName, "|") Or InStr(1, szName, "^") Or InStr(1, szName, ".") Or InStr(1, szName, ",") Then
    lblNameError.Caption = "(Invalid name)"
Else
    lblNameError.Caption = ""
End If

End Sub

' Turns an IP into a username...simple.
Public Function ResolveIPToName(ByVal szIP As String) As String
Dim op As OnlinePerson
Dim szName As String

szName = "<empty>"
For Each op In colUsers
    If op.IPAddress = szIP Then
        szName = op.Name
        Exit For
    End If
Next op

ResolveIPToName = szName
End Function

' Resolves a name into a mangled IP.
Public Function ResolveNameToIP(ByVal szName As String) As String
Dim op As OnlinePerson
Dim szIP As String

szIP = "<empty>"
For Each op In colUsers
    If op.Name = szName Then
        szIP = op.IPAddress
        Exit For
    End If
Next op

ResolveNameToIP = szIP
End Function

' Removes an ignored mangled IP from the collection of ignored IPs.
' Assumes that the IP does indeed exist.
Public Sub RemoveIgnored(ByVal szIP As String)
colIgnored.Remove szIP
End Sub

' Checks to see if a particular IP is ignored or not
Private Function IsIgnored(ByVal szIP As String, ByVal szName As String) As Boolean
Dim bFound As Boolean
Dim op As OnlinePerson

' Just do a linear search through the collection..
' Speed of this is generally slow, for large collections, but in CC,
' nobody's usually ignoring more than three people.
bFound = False
For Each op In colIgnored
    If szIP = op.IPAddress Then
        bFound = True
        If op.Name <> szName Then
            ' Update the name entry.
            op.Name = szName
        End If
        Exit For
    End If
Next op

IsIgnored = bFound
End Function

' Ignores a user, transmitting results to the screen if wanted.
Private Sub IgnoreUser(ByVal szName As String, ByVal szIP As String, Optional bLoud As Boolean = True, Optional ByVal bMutual As Boolean = False)
Dim op As OnlinePerson
Dim bFound As Boolean

On Error GoTo crash

' First, sanity check: see if it's already in the ignored collection.
bFound = False
For Each op In colIgnored
    If op.IPAddress = szIP Then
        bFound = True
        Exit For
    End If
Next op

If Not bFound Then
    ' Create and add the new ignored entry...
    Set op = New OnlinePerson
    op.IPAddress = szIP
    op.Name = szName
    op.Connected = True
    colIgnored.Add op, szIP
    ' A "two sided ignore" is an ignore maintained by both sides. That is,
    ' each mutually ignores the other. Mutual ignore means different things here,
    ' however.
    If mbTwoSidedIgnore And Not bMutual Then
        ' Send the server a little command to tell it that we
        ' are ignoring this person.
        tcpCyan.SendData "70|0" & szName & vbCrLf
    End If
    
    ' Let the user know.
    If bLoud Then
        PrintMessage "You are now ignoring [" & szName & "] and all of their aliases from address " & colUsers(szName).IPAddress & ".", msgMagenta, "[Magenta]"
    End If
Else
    ' This shouldn't ever happen by how the program works, but just in
    ' case it does...
    MsgBox "You are already ignoring the mangled IP " & szIP & ".", vbExclamation
End If
Exit Sub

crash:
    ReportCrash Err, "frmChat", "IgnoreUser", V(szName, "szName"), _
        V(szIP, "szIP"), V(mbStarted, "mbStarted"), V(mbCyan, "mbCyan"), V(mIsNamed, "mIsNamed"), _
        V(mszName, "mszName"), V(bLoud, "bLoud"), _
        V(bMutual, "bMutual"), V(bFound, "bFound")
End Sub
