Attribute VB_Name = "modGlobals"
Option Explicit

' Constants for packets

' All possible user types
Public Enum NameType
    ntRegular
    ntCyan
    ntChatServer
    ntChatClient
    ntSpecialGuest
End Enum

' The types of messages that can be displayed
Public Enum MessageType
    msgServer
    msgChat
    msgMagenta
    msgGuest
    msgCyan
    msgPrivate ' Protocol "extension" - for printing sent private message text.
End Enum

' Message formats, as based on the numeric code sent with a message.
Public Enum MessageFormat
    mfPrivate
    mfBroadcast
    mfEnter
    mfLeave
End Enum

' These will eventually be phased out as constants.
Public Const REMOTE_HOST = "cho.cyan.com"
Public Const REMOTE_PORT = 1812

' Colors
Public Const COLOR_LIME = &HFF00&
Public Const COLOR_GOLD = &H80C0FF
Public Const COLOR_GRAY = &HC0C0C0

' Variables everyone needs to be able to get at
Public mszHostname As String           ' Cyan Chat host
Public mlPort As Long                   ' Cyan Chat port

Sub Main()

' Right, we only want one instance of Magenta running at a time,
' soooo...
If App.PrevInstance Then
    MsgBox "A previous instance of Magenta is already running.", vbInformation
Else
    ' VB dun terminate until the last form is gone...
    frmChat.Show
End If
End Sub

' Generates and prints a crash report.
Public Sub ReportCrash(errInfo As VBA.ErrObject, ByVal szObject As String, ByVal szProcedure As String, ParamArray arrVariables() As Variant)
Dim szMessage As String
Dim i As Long
Dim ubnd As Long
Dim lbnd As Long

' If the socket on the forms are open, shut them now to stop subsequent
' errors, also...since we've crashed. Why should we be receiving anything else?
If frmChat.tcpCyan.State = sckConnected Then
    frmChat.tcpCyan.Close
End If

If frmStatus.tcpCyanChat.State = sckConnected Then
    frmStatus.tcpCyanChat.Close
End If

' First, we print out the requisite blather.
szMessage = "[" & Now & "]" & vbCrLf
szMessage = szMessage & "A program error has occurred. Please report this to the author by" & vbCrLf
szMessage = szMessage & "copying and pasting this report into an email, and sending it to" & vbCrLf
szMessage = szMessage & "yohshee@hotmail.com with the subject " & Chr$(34) & "Magenta Bug Report" & Chr$(34) & "." & vbCrLf & vbCrLf

' Now comes the good stuff...
szMessage = szMessage & "Error Object Information:" & vbCrLf
szMessage = szMessage & "Number: " & errInfo.Number & " (0x" & Hex$(errInfo.Number) & ")" & vbCrLf
szMessage = szMessage & "Description: " & errInfo.Description & vbCrLf
szMessage = szMessage & "Apparent Source: " & errInfo.Source & vbCrLf
szMessage = szMessage & "Last DLL Error code: " & Hex$(errInfo.LastDllError) & vbCrLf & vbCrLf

szMessage = szMessage & "Crash originated in " & szObject & "." & szProcedure & "." & vbCrLf & vbCrLf
szMessage = szMessage & "Variables in " & szObject & "." & szProcedure & ": " & vbCrLf

ubnd = UBound(arrVariables)
lbnd = LBound(arrVariables)
For i = lbnd To ubnd
    szMessage = szMessage & arrVariables(i) & vbCrLf
Next i

szMessage = szMessage & vbCrLf & "End of crash report."
frmCrash.txtPanel.Text = szMessage
frmCrash.Show vbModal

' Just exit without doing a LICK of cleaning up.
End
End Sub

' Formats a variable for ya.
Public Function V(ByVal var As Variant, ByVal szName As String) As String
' Unfortunately, because object references that are Nothing screw with everything,
' we have to put in more error checking here.
If TypeName(var) = "Nothing" Then
    V = "Object " & szName & " = Nothing"
ElseIf TypeOf var Is Object  Then
    V = "Object " & szName & " = [Object Reference (" & CStr(ObjPtr(var)) & ")]"
ElseIf TypeName(var) = "String" Then
    V = "String " & szName & " = " & Chr$(34) & var & Chr$(34)
Else
    V = TypeName(var) & " " & szName & " = " & CStr(var)
End If
End Function

