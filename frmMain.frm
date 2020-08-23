VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmStatus 
   Caption         =   "Magenta - Cyan Chat Status"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgSave 
      Left            =   2760
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSWinsockLib.Winsock tcpCyanChat 
      Left            =   2880
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   2040
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0626
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0942
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwData 
      Height          =   2895
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   5106
      Arrange         =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "imlIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar sbConnection 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2940
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuLog 
         Caption         =   "&Log All Incoming Packets"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuSEP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "C&lose"
      End
   End
   Begin VB.Menu mnuConnection 
      Caption         =   "&Connection"
      Begin VB.Menu mnuConnect 
         Caption         =   "&Connect"
      End
      Begin VB.Menu mnuDisconnect 
         Caption         =   "&Disconnect"
      End
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ==============================================
' Magenta <Source>
' A program that can work with the Cyan Chat server as well as
' the Java applet can. Well, hopefully, anyway.
'
' Developed by Rick Coogle
'
' This form: A glorified Who's online window.
' ==============================================

Private mCancel As Boolean
Private mFilenum As Integer
Private mLogging As Boolean
Private mStarted As Boolean
Private colNormal As Collection

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Form_Activate()
If Not mStarted Then
    ' Automatically try to connect to cho on our own.
    Call mnuConnect_Click
    mStarted = True
End If
End Sub

Private Sub Form_Load()
mnuDisconnect.Enabled = False
mStarted = False
End Sub

Private Sub Form_Resize()
Dim lNewHeight As Long
Dim lNewWidth As Long

lNewHeight = ScaleHeight - sbConnection.Height - 20
lNewWidth = ScaleWidth - 15
lvwData.Width = lNewWidth
lvwData.Height = lNewHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
' Check the connection first
If tcpCyanChat.State <> sckClosed Then
    ' Close it as well as the log
    tcpCyanChat.Close
    If mLogging Then Close #mFilenum
End If
End Sub

Private Sub lvwData_DblClick()
With lvwData.SelectedItem
    frmProperties.txtName.Text = colNormal(.Text).Name
    frmProperties.txtDNS.Text = colNormal(.Text).DNSEntry
    frmProperties.txtIP.Text = colNormal(.Text).IPAddress
End With

frmProperties.Show vbModal
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show vbModal
End Sub

Private Sub mnuChat_Click()
frmChat.Show
End Sub

Private Sub mnuConnect_Click()

Me.MousePointer = vbHourglass

tcpCyanChat.Connect mszHostname, mlPort

mnuConnection.Enabled = False

Do
    Select Case tcpCyanChat.State
        Case sckResolvingHost
            sbConnection.SimpleText = "Resolving host..."
        Case sckHostResolved
            sbConnection.SimpleText = "Host resolved..."
        Case sckConnecting
            sbConnection.SimpleText = "Connecting to host..."
        Case sckConnectionPending
            sbConnection.SimpleText = "Connection pending to host..."
        Case sckError
            sbConnection.SimpleText = "Error in connection."
            Exit Do
        Case sckConnected
            Exit Do
    End Select
    ' Play nice with Windows
    DoEvents
Loop

If tcpCyanChat.State = sckError Then
    Call mnuDisconnect_Click
    sbConnection.SimpleText = "Error in connection"
    mnuFile.Enabled = True
    mnuConnection.Enabled = True
    Me.MousePointer = vbDefault
    Exit Sub
End If

sbConnection.SimpleText = "Connection established to " & mszHostname & _
    ":" & CStr(mlPort) & "."
mnuConnect.Enabled = False
mnuDisconnect.Enabled = True
mnuConnection.Enabled = True
mnuFile.Enabled = True
mFilenum = FreeFile
If mLogging Then
    Open App.Path & "\Magenta.log" For Output As #mFilenum
    Print #mFilenum, "MAGENTA: Magenta log begun at: " & Time$ & " on " & Date$ & "."
End If
Me.MousePointer = vbDefault
End Sub

Private Sub mnuDisconnect_Click()
' Close the log and the socket
mCancel = True
sbConnection.SimpleText = "Connection to cho.cyan.com closed"
tcpCyanChat.Close

If mLogging Then
    On Error Resume Next
    Print #mFilenum, "MAGENTA: Magenta log stopped at: " & Time$ & " on " & Date$ & "."
    Close #mFilenum
    On Error GoTo 0
End If

mnuDisconnect.Enabled = False
mnuConnect.Enabled = True
End Sub

Private Sub mnuClose_Click()
Unload Me
End Sub

Private Sub mnuLog_Click()
mnuLog.Checked = Not mnuLog.Checked

If mnuLog.Checked Then
    mFilenum = FreeFile
    If Len(Dir$(App.Path & "\Magenta.log")) > 0 Then
        Open App.Path & "\Magenta.log" For Append As #mFilenum
        Print #mFilenum, "MAGENTA: Magenta log restarted at: " & Time$ & " on " & Date$ & "."
    Else
        Open App.Path & "\Magenta.log" For Output As #mFilenum
        Print #mFilenum, "MAGENTA: Magenta log started at: " & Time$ & " on " & Date$ & "."
    End If
    mLogging = True
Else
    ' Close it
    Print #mFilenum, "MAGENTA: Magenta log stopped at: " & Time$ & " on " & Date$ & "."
    Close #mFilenum
End If
End Sub

Private Sub mnuSave_Click()
Dim FileName As String
Dim FileNum As Integer
Dim op As OnlinePerson

On Error GoTo err_handler:

' This routine iterates through the collection, grabbing info
' and saving it to a INI-type file. Not as if you can actually DO
' anything with this result, but it's there.

With dlgSave
    .DialogTitle = "Save Who's Online List"
    .Filter = "Online List Files (*.olf)|*.olf|Text Files (*.txt)|(*.txt)|All Files (*.*)|*.*|"
    .ShowSave
    FileName = .FileName
End With

FileNum = FreeFile

Open FileName For Output As #FileNum

Print #FileNum, "; Online list log generated at " & Now
Print #FileNum, ";"

For Each op In colNormal
    Print #FileNum, "[" & op.Name & "]"
    Print #FileNum, "DNSEntry = " & op.DNSEntry
    Print #FileNum, "IPAddress = " & op.IPAddress
    ' Even though this is totally irrelevant...
    Print #FileNum, "Ignored = " & CStr(op.Ignored)
    Print #FileNum,
Next op

Close #FileNum

Exit Sub

err_handler:
    If Err.Number = cdlCancel Then
        ' Ignore it
    Else
        MsgBox "An error occurred while trying to save a file." & vbCrLf & _
            Err.Description, vbExclamation, "Error"
    End If
End Sub

Private Sub tcpCyanChat_DataArrival(ByVal bytesTotal As Long)
Dim szData As String
Dim szBuffer As String
Dim szName As String
Dim nTypeFlag As Integer
Dim i As Long
Dim ubnd As Long
Dim lPos As Long
Dim lSlash As Long
Dim DataArray() As String
Dim op As OnlinePerson
Dim itm As ListItem

' Simple little command loop.
tcpCyanChat.GetData szData

If mLogging Then
    Print #mFilenum, "[" & Now & "]" & szData
End If

' Clear it all out.
lvwData.ListItems.Clear
Set colNormal = Nothing
Set colNormal = New Collection

' It is a namelist, so we don't have to do flag-checking.
SplitString szData, DataArray(), "|"
ubnd = UBound(DataArray)
For i = 2 To ubnd
    szBuffer = DataArray(i)
    If IsNumeric(Left$(szBuffer, 1)) Then
        nTypeFlag = CInt(Left$(szBuffer, 1))
    End If
    lPos = InStr(1, szBuffer, ",")
    szName = Mid$(szBuffer, 2, lPos - 2)
    Set op = New OnlinePerson
    'Certain things were removed from the CyanChat protocol by MarkD,
    ' such as DNS resolution and actual IP addresses.
    ' -RC, 12/6/2001
    
    'lSlash = InStr(1, szBuffer, "/")
    With op
        .Name = szName
        If nTypeFlag = ntRegular Then
            '.DNSEntry = Mid$(szBuffer, lPos + 1, lSlash - (lPos + 1))
            .DNSEntry = "Unknown"
            ' The IP address now is a mangled long integer.
            .IPAddress = Mid$(szBuffer, lPos + 1)
        ElseIf nTypeFlag = ntCyan Then
            .DNSEntry = "Cyan"
            .IPAddress = "local"
        ElseIf nTypeFlag = ntSpecialGuest Then
            .DNSEntry = "Cyan Guest"
            .IPAddress = "local"
        End If
    End With
    colNormal.Add op, op.Name
    If nTypeFlag = ntRegular Then
        Set itm = lvwData.ListItems.Add(, , szName, 1)
    ElseIf nTypeFlag = ntCyan Then
        Set itm = lvwData.ListItems.Add(, , szName, 2)
    ElseIf nTypeFlag = ntSpecialGuest Then
        Set itm = lvwData.ListItems.Add(, , szName, 3)
    End If
Next i
Set itm = Nothing
Set op = Nothing
End Sub

Private Sub tcpCyanChat_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Dim szMessage As String

If Number <> sckConnectionReset Then
    szMessage = "Error #" & Number & vbCrLf & Description & vbCrLf & vbCrLf
    szMessage = szMessage & "Correct any issues with your Internet connection (if need be) and attempt to " & vbCrLf
    szMessage = szMessage & "connect again using the options under the Connection menu." & vbCrLf
    szMessage = szMessage & "If that does not fix the problem," & vbCrLf
    szMessage = szMessage & mszHostname & " may be down at the moment, or a router might be" & vbCrLf
    szMessage = szMessage & "may be down on your path. So, be patient, and hang on."
        
    MsgBox szMessage, vbExclamation, "Connection Error"
    sbConnection.SimpleText = "Error in connection: " & Description
Else
    ' Attempt to reconnect.
    sbConnection.SimpleText = "Connection reset by peer. Attempting to reconnect automatically in 1 second..."
    Sleep 1000
    Call mnuDisconnect_Click
    Sleep 10
    Call mnuConnect_Click
End If
End Sub
