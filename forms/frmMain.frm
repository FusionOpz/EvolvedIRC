VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "EvolvedIRC Codename ""Grasshopper"", BuildID:20040212-devel"
   ClientHeight    =   7155
   ClientLeft      =   3660
   ClientTop       =   3225
   ClientWidth     =   9855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   9855
   Begin MSComctlLib.TreeView tvUsers 
      Height          =   6015
      Left            =   7920
      TabIndex        =   3
      Top             =   1080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   10610
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "iUsers"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin RichTextLib.RichTextBox rtfBuffer 
      Height          =   5655
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   9975
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":29C12
   End
   Begin MSWinsockLib.Winsock sckIRC 
      Left            =   9240
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtChat 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   6840
      Width           =   7815
   End
   Begin MSComctlLib.ImageList iUsers 
      Left            =   8520
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":29C8D
            Key             =   "owner"
            Object.Tag             =   "owner"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A227
            Key             =   "user"
            Object.Tag             =   "user"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A7C1
            Key             =   "host"
            Object.Tag             =   "host"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2AD5B
            Key             =   "halfop"
            Object.Tag             =   "halfop"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B2F5
            Key             =   "voice"
            Object.Tag             =   "voice"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B88F
            Key             =   "admin"
            Object.Tag             =   "admin"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTopic 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Channel Topic."
      ForeColor       =   &H80000009&
      Height          =   195
      Left            =   1680
      TabIndex        =   7
      Top             =   720
      Width           =   1065
   End
   Begin VB.Label lblTpc 
      BackStyle       =   0  'Transparent
      Caption         =   "Topic:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblChannel 
      BackStyle       =   0  'Transparent
      Caption         =   "#channel"
      ForeColor       =   &H80000009&
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label lblChn 
      BackStyle       =   0  'Transparent
      Caption         =   "Channel:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "EvolvedIRC"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   615
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   4575
   End
   Begin VB.Image imgTop 
      DragIcon        =   "frmMain.frx":2BE29
      Height          =   720
      Left            =   120
      Picture         =   "frmMain.frx":55A3B
      Top             =   120
      Width           =   720
   End
   Begin VB.Shape shTop 
      BorderColor     =   &H00000000&
      FillColor       =   &H80000002&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   9855
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuConnect 
         Caption         =   "Connect"
      End
      Begin VB.Menu mnuDiscon 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu m 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuServ 
      Caption         =   "Servers"
      Begin VB.Menu mnuFnn 
         Caption         =   "FreeNode.Net"
      End
      Begin VB.Menu mnuWbo 
         Caption         =   "WinBeta.Org"
      End
   End
   Begin VB.Menu mnuChan 
      Caption         =   "Channels"
      Begin VB.Menu mnuE2G 
         Caption         =   "#Evolved2Go"
      End
      Begin VB.Menu mnuEIRC 
         Caption         =   "#EvolvedIRC"
      End
      Begin VB.Menu mnuIP 
         Caption         =   "#Ignition-Project"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuOpt 
         Caption         =   "&Options"
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About EvolvedIRC"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'EvolvedIRC is (C)  Matthew Sporich.
'----------------------------------------------------------
'You must include this notice in any modifications you make. You must additionally
'follow the GPL's provisions for sourcecode distribution and binary distribution.
'If you are not familiar with the GPL, please read LICENSE.TXT.
'(you are welcome to add a "Based On" line above this notice, but this notice must
'remain intact!)
'Released under the GNU General Public License
'Contact information: Matthew Sporich (DJ_Dark) <djdark@gmail.com>
'                     Evolved2Go Support (Support) <support.evolved2go@gmail.com>
'                     Website <http://evolved2go.ws4f.us/>
'
' $Id: frmMain.frm,v 1.8 2005/03/02 00:55:02 dj_dark Exp $
'
'
'This program is free software.
'You can redistribute it and/or modify it under the terms of the
'GNU General Public License as published by the Free Software Foundation; either version 2 of the License,
'or (at your option) any later version.
'
'This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY.
'Without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
'See the GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License along with this program.
'if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA

'TODO:
'Start formattig the recived code - In Progress,
'Start coding Options System - In Progress,
'Make use of the User list box - Almost done,

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hWnd As Long, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long
'Implements ISubclass

Option Explicit

'INI Stuff
Dim SectionName As String
Dim KeyName As String
Dim Value As String
Private inifile As String

Dim Version As String
Dim User As String
Dim Nick As String
Dim Channel As String
Dim RName As String
Dim Server As String
Dim Port As Integer
Dim Topic As String
Dim IRCX As Integer
Dim isIRCX As String


'TODO: Write Options dialog to set Server address, server port, nickname and username,
'Also get it to either save the optios to a Var or into a File.
'TODO: now get to coding the dialog so it'll remember the settings.

Private Sub Form_Load()
    Debug.Print "EvolvedIRC Loaded"
    frmMain.Caption = App.Title & " (BuildID: " & BuildID & ")"
    Version = App.Major & "." & App.Minor & "." & App.Revision & " (Build ID " & BuildID & ")"
    Server_AuthNotices = "false"
    'Nick = "EvolvedIRC_PR"
    'User = "EvolvedIRC User"
    'Channel = "#Lobby"
    'Server = "localhost"
    'Port = "6667"
    Topic = ""

    RTF_AutoURLDetect True
    RTF_SetViewRect 10, 10
    'Call Connect

Call INI_Load
    'INI Load
        'User Options
       ' Nick = ReadINI("userinfo", "nickname", App.Path + "\options.ini")
       ' User = ReadINI("userinfo", "username", App.Path + "\options.ini")
       ' RName = ReadINI("userinfo", "realname", App.Path + "\options.ini")
    
        'Server Options
       ' Server = ReadINI("server", "address", App.Path + "\options.ini")
       ' Port = ReadINI("server", "port", App.Path + "\options.ini")
       ' Channel = ReadINI("server", "defaultchan", App.Path + "\options.ini")
       ' IRCX = ReadINI("server", "IRCX", App.Path + "\options.ini")
       '
       ' If IRCX = "1" Then
       '     isIRCX = "IRCX"
       ' ElseIf IRCX = "0" Then
       '     isIRCX = "IRC"
       ' End If
End Sub

Private Sub INI_Load()
Debug.Print "Options Loaded"
'INI Load
        'User Options
        Nick = ReadINI("userinfo", "nickname", App.Path + "\options.ini")
        User = ReadINI("userinfo", "username", App.Path + "\options.ini")
        RName = ReadINI("userinfo", "realname", App.Path + "\options.ini")
    
        'Server Options
        Server = ReadINI("server", "address", App.Path + "\options.ini")
        Port = ReadINI("server", "port", App.Path + "\options.ini")
        Channel = ReadINI("server", "defaultchan", App.Path + "\options.ini")
        IRCX = ReadINI("server", "IRCX", App.Path + "\options.ini")
        
        If IRCX = "1" Then
            isIRCX = "IRCX"
        ElseIf IRCX = "0" Then
            isIRCX = "IRC"
        End If
End Sub

Public Function ReadINI(SectionHeader As String, _
    VariableName As String, _
    FileName As String) As String
    
    Dim strReturn As String
    strReturn = String(255, Chr(0))
    ReadINI = Left$(strReturn, GetPrivateProfileString(SectionHeader, _
        ByVal VariableName, "", strReturn, Len(strReturn), FileName))
End Function

Public Function WriteINI(SectionHeader As String, _
    VariableName As String, _
    Value As String, _
    FileName As String)

    WriteINI = WritePrivateProfileString(SectionHeader, _
        VariableName, _
        Value, _
        FileName)
    End Function

Private Sub Connect()
On Error Resume Next
    With sckIRC
        .RemoteHost = Server
        .RemotePort = Port
        .Connect
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    sckIRC.SendData "QUIT : Time for me to go l8r" & vbCrLf
    sckIRC.Close
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuConnect_Click()
    Call Connect
End Sub

Private Sub mnuDiscon_Click()
    sckIRC.SendData "QUIT : Time for me to go l8r" & vbCrLf
    sckIRC.Close
End Sub

Private Sub mnuExit_Click()
On Error Resume Next
    sckIRC.Close
    Unload Me
    Unload frmOptions
    Unload frmAbout
End Sub

Private Sub mnuFnn_Click()
On Error Resume Next
    sckIRC.Close
    Server = "irc.freenode.net" 'The IRC server
    Port = "6667" 'Connect on port 6667
    Debug.Print "Server address changed to:" & Server & " And port changed to" & Port
    Call Connect
End Sub

Private Sub mnuOpt_Click()
    frmOptions.Show
End Sub

Private Sub mnuWbo_Click()
On Error Resume Next
    sckIRC.Close
    Server = "irc.winbeta.org" 'The IRC server
    Port = "6667" 'Connect on port 6667
    Debug.Print "Server address changed to: " & Server & " And port changed to: " & Port
    Call Connect
End Sub

'TODO: Write Options dialog to set Server address, server port, nickname and username,
'Also get it to either save the optios to a Var or into a File.
'TODO: now get to coding the dialog so it'll remember the settings.
Private Sub sckIRC_Connect()
    With sckIRC
        .SendData "NICK " & Nick & vbCrLf
        sckIRC.SendData "USER " & User & " """ & sckIRC.LocalHostName & """ """ & Server & """ :" & RName & vbCrLf
        
        '.SendData "USER " & ": " & User & sckIRC.LocalHostName & " " & _
         'UCase(sckIRC.LocalHostName & ":" & sckIRC.LocalPort & "/0") & _
         '" :EvolvedIRC Pre-Alpha Client" & vbCrLf
        
        '.SendData "USER " & User & sckIRC.LocalHostName & " " & _
        '    UCase(sckIRC.LocalHostName & ":" & sckIRC.LocalPort & "/0") & " :EvolvedIRC Pre-Alpha Client" & vbCrLf
        '.SendData "JOIN " & Channel & vbCrLf
    End With
End Sub

Function RightOf(strData As String, strDelim As String) As String
    Dim tmpData As String
    tmpData = strData
    If Left(tmpData, 1) = ":" Then tmpData = Right(tmpData, Len(tmpData) - 1)
    Dim intPos As Integer
    intPos = InStr(tmpData, strDelim)
    
    If intPos Then
        RightOf = Mid(tmpData, intPos + 1, Len(tmpData) - intPos)
    Else
        RightOf = tmpData
    End If
End Function

Private Sub sckIRC_DataArrival(ByVal bytesTotal As Long)
    'On Error Resume Next
    'Dim sRecv As String
    '
    'sckIRC.GetData sRecv 'Put the data recieved into the string
    'InternalDebug sRecv
'
'
    'Play ping pong with the server
    'If Split(sRecv, " ")(0) = "PING" Then
    '    sckIRC.SendData "PONG " & Split(sRecv, " ")(1)
    'End If
'
    'Update the buffer
    'rtfBuffer.Text = rtfBuffer.Text & sRecv & vbCrLf
    'rtfBuffer.SelStart = Len(rtfBuffer.Text)
'
'
Dim tmpString As String
Dim tmpSplitLF() As String
Dim tmpSplit() As String
Dim tmpPrefix As String
Dim tmpNamesList As String
Dim tmpNamesArray() As String
Dim tmpRand As Long
Dim tmpPartMsg As String
Dim a As Long
Dim B As Long

sckIRC.GetData tmpString, vbString
InternalDebug tmpString
tmpString = Replace(tmpString, vbCrLf, vbLf)
tmpString = Replace(tmpString, vbCr, vbLf)
tmpSplitLF = Split(tmpString, vbLf)

For a = LBound(tmpSplitLF) To UBound(tmpSplitLF)
  tmpSplit = Split(tmpSplitLF(a), " ")
  If Len(tmpSplitLF(a)) = 0 Then GoTo NextLine
 
  If UCase$(tmpSplit(0)) = "PING" Then
    sckIRC.SendData "PONG " & tmpSplit(1) & vbCrLf
  ElseIf UCase$(tmpSplit(0)) = "ERROR" Then
    RTF_SetColor QBColor(12)
    RTF_AddText ilIndent & RightOf(tmpSplitLF(a), ":") & vbCrLf
    sckIRC.Close
  End If
  
  tmpPrefix = tmpSplit(0)
     
  Select Case UCase$(tmpSplit(1))
    Case "001"
      sckIRC.SendData "USERHOST " & User & vbCrLf
    'Case "002"
    'Case "003"
    'Case "004"
    Case "005"
      If UCase$(Left(tmpSplit(3), 4)) = "IRCX" Then
        sckIRC.SendData "IRCX" & vbCrLf
      'ElseIf UCase$(Left(tmpSplit(3), 4)) = "IRC" Then
      '  sckIRC.SendData "IRC" & vbCrLf
      End If
    Case "251"
      'we send this here, because some servers don't send a 005
      'but all servers are supposed to send this...
      sckIRC.SendData "JOIN " & Channel & vbCrLf
    Case "332"
      Topic = RightOf(tmpSplitLF(a), ":")
      RTF_SetBold False
      RTF_SetColor QBColor(0)
      RTF_AddText ilIndent
      RTF_SetBold True
      'RTF_AddText Replace(Split(tmpPrefix, "!")(0), ":", "") & ": "
      'RTF_SetBold False
      RTF_AddText "Topic: " & Topic & vbCrLf
      lblTopic.Caption = Topic
    Case "353"
      'this code is licensed under the GNU General Public
      'License. please do not copy this code unless you intend on
      'releasing the source under the GNU General Public License (or
      'a compatible license)
      '(this notice is here because a lot of people would love to have
      '100% working names code in their non-GPL'ed client)
      
      ':localhost 353 Ziggy = #Ziggy :.Ziggy
      'on with the show!
      tmpNamesList = RightOf(tmpSplitLF(a), ":")
      tmpNamesArray = Split(tmpNamesList, " ")
      
      For B = 0 To UBound(tmpNamesArray)
        If Left(tmpNamesArray(B), 1) = "." Then
          tvUsers.Nodes.Add , , Mid(tmpNamesArray(B), 2), Mid(tmpNamesArray(B), 2) & " (Owner)", "owner", "owner"
        ElseIf Left(tmpNamesArray(B), 1) = "!" Then
          tvUsers.Nodes.Add , , Mid(tmpNamesArray(B), 2), Mid(tmpNamesArray(B), 2) & " (Owner)", "owner", "owner"
        ElseIf Left(tmpNamesArray(B), 1) = "@" Then
          tvUsers.Nodes.Add , , Mid(tmpNamesArray(B), 2), Mid(tmpNamesArray(B), 2) & " (Host)", "host", "host"
        ElseIf Left(tmpNamesArray(B), 1) = "%" Then
          tvUsers.Nodes.Add , , Mid(tmpNamesArray(B), 2), Mid(tmpNamesArray(B), 2) & " (HalfOp)", "halfop", "halfop"
        ElseIf Left(tmpNamesArray(B), 1) = "+" Then
          tvUsers.Nodes.Add , , Mid(tmpNamesArray(B), 2), Mid(tmpNamesArray(B), 2) & " (Voice)", "voice", "voice"
        Else
          tvUsers.Nodes.Add , , tmpNamesArray(B), tmpNamesArray(B) & "", "user", "user"
        End If
      Next B
    Case "372"
        RTF_SetBold False
        RTF_SetColor QBColor(0)
        RTF_AddText ilIndent
        RTF_SetBold True
        RTF_AddText Replace(Split(tmpPrefix, "!")(0), ":", "") & ": "
        RTF_SetBold False
        RTF_AddText RightOf(tmpSplitLF(a), ":") & vbCrLf
    Case "433"
      ':localhost 433 Anonymous Ziggy :Nickname is already in use
      RTF_SetColor QBColor(6)
      Randomize Timer
      tmpRand = Int(Rnd * 100)
      RTF_AddText ilIndent & "The nickname " & Nick & " is already in use. Trying " & Nick & tmpRand & "." & vbCrLf
      Nick = Nick & tmpRand
      sckIRC.SendData "NICK :" & Nick & vbCrLf
    Case "NOTICE"
      Select Case UCase$(tmpSplit(2))
        Case "AUTH"
          'If Server_AuthNotices Then
            RTF_SetColor QBColor(3)
            RTF_AddText ilIndent & RightOf(tmpSplitLF(a), ":") & vbCrLf
          'End If
      End Select
      'If Replace(Split(tmpPrefix, "!")(0), ":", "") = NickServ Then
        
    Case "JOIN"
      If UCase$(Replace(Split(tmpPrefix, "!")(0), ":", "")) = UCase$(Nick) Then
        'I joined
        RTF_AddText vbCrLf
        RTF_SetColor QBColor(1)
        RTF_AddText ilIndent & "You are now chatting on " & Channel & vbCrLf
        tvUsers.Nodes.Clear
        lblChannel.Caption = Channel
      Else
        RTF_Indent
        RTF_AddBullet Arrow, QBColor(8)
        RTF_SetColor QBColor(8)
        RTF_AddText " " & Replace(Split(tmpPrefix, "!")(0), ":", "") & " has joined the conversation." & vbCrLf
        tvUsers.Nodes.Add , , Replace(Split(tmpPrefix, "!")(0), ":", ""), Replace(Split(tmpPrefix, "!")(0), ":", "")
      End If
    Case "PING"
        RTF_Indent
        RTF_AddBullet Arrow, QBColor(8)
        RTF_SetColor QBColor(8)
        RTF_AddText "Ping" & vbCrLf
        RTF_Indent
        RTF_AddBullet Arrow, QBColor(8)
        RTF_SetColor QBColor(8)
        RTF_AddText "Pong" & vbCrLf
    Case "PART"
      If UCase$(Replace(Split(tmpPrefix, "!")(0), ":", "")) = UCase$(Nick) Then
        'I parted
        RTF_AddText vbCrLf
        RTF_SetColor QBColor(1)
        RTF_AddText ilIndent & "You no longer are chatting on " & Channel & vbCrLf
        tvUsers.Nodes.Clear
      Else
        RTF_Indent
        RTF_AddBullet Arrow, QBColor(8)
        RTF_SetColor QBColor(8)
        If UBound(tmpSplit) = 2 Then
          RTF_AddText " " & Replace(Split(tmpPrefix, "!")(0), ":", "") & " has left the conversation." & vbCrLf
        Else
          RTF_AddText " " & Replace(Split(tmpPrefix, "!")(0), ":", "") & " has left the conversation (" & RightOf(tmpSplitLF(a), ":") & ")." & vbCrLf
        End If
        RemoveKey Replace(Split(tmpPrefix, "!")(0), ":", "")
      End If
    Case "QUIT"
      If UCase$(Replace(Split(tmpPrefix, "!")(0), ":", "")) = UCase$(Nick) Then
        'I Quit
        'Unload Me
      Else
        RTF_Indent
        RTF_AddBullet Arrow, QBColor(8)
        RTF_SetColor QBColor(8)
        If UBound(tmpSplit) = 2 Then
          RTF_AddText " " & Replace(Split(tmpPrefix, "!")(0), ":", "") & " has left the server." & vbCrLf
        Else
          RTF_AddText " " & Replace(Split(tmpPrefix, "!")(0), ":", "") & " has left the server (" & RightOf(tmpSplitLF(a), ":") & ")." & vbCrLf
        End If
        RemoveKey Replace(Split(tmpPrefix, "!")(0), ":", "")
      End If
    'Case "MODE"
    '    If UCase$(Replace(Split(tmpPrefix, "!")(0), ":", "")) = UCase$(Nick) Then
    '    'My Mode Was Changed
    '  Else
    '    RTF_Indent
    '    RTF_AddBullet Arrow, QBColor(8)
    '    RTF_SetColor QBColor(8)
    '    If UBound(tmpSplit) = 2 Then
    '      RTF_AddText " " & Replace(Split(tmpPrefix, "!")(0), ":", "") & " has left the server." & vbCrLf
    '    Else
    '      RTF_AddText " " & Replace(Split(tmpPrefix, "!")(0), ":", "") & " has left the server (" & RightOf(tmpSplitLF(a), ":") & ")." & vbCrLf
    '    End If
    '    RemoveKey Replace(Split(tmpPrefix, "!")(0), ":", "")
    '  End If
    Case "PRIVMSG"
      Select Case UCase$(tmpSplit(2))
        Case UCase$(Channel)
          'RTF_SetColor QBColor(2)
          'RTF_SetBold True
          'RTF_AddText ilIndent & Replace(Split(tmpPrefix, "!")(0), ":", "") & ": " & RightOf(tmpSplitLF(a), ":") & vbCrLf
          RTF_SetBold False
          RTF_SetColor QBColor(0)
          RTF_AddText ilIndent
          RTF_SetBold True
          RTF_AddText Replace(Split(tmpPrefix, "!")(0), ":", "") & ": "
          RTF_SetBold False
          RTF_AddText RightOf(tmpSplitLF(a), ":") & vbCrLf
        Case UCase$(Nick)
          
          'RTF_SetColor QBColor(2)
          'RTF_SetBold True
          'RTF_AddText ilIndent & Replace(Split(tmpPrefix, "!")(0), ":", "") & ": " & RightOf(tmpSplitLF(a), ":") & vbCrLf
          RTF_SetBold False
          RTF_SetColor QBColor(0)
          RTF_AddText ilIndent
          RTF_SetBold True
          RTF_AddText Replace(Split(tmpPrefix, "!")(0), ":", "") & ": "
          RTF_SetBold False
          RTF_AddText RightOf(tmpSplitLF(a), ":") & vbCrLf
        'Case UCase$(RightOf(Nick, ":"))
          'If Chr(1) & "VERSION" & Chr(1) Then
            'RTF_Indent
            'RTF_AddBullet Arrow, QBColor(8)
            'RTF_SetColor QBColor(8)
            'RTF_AddText ilIndent & Replace(Split(tmpPrefix, "!")(0), ":", "") & " has asked for your version" & vbCrLf
            'sckIRC.SendData "PRIVMSG " & Channel & " :" & txtChat.Text & vbCrLf
            'RTF_SetBold False
            'RTF_SetColor QBColor(0)
            'RTF_AddText ilIndent
            'RTF_SetBold True
            'RTF_AddText Nick & ": "
            'RTF_SetBold False
            'RTF_AddText txtChat.Text & vbCrLf
          'End If
      End Select
    
  End Select
NextLine:
Next a
    
End Sub
Public Sub RemoveKey(Key As String)
Dim a As Long
For a = 1 To tvUsers.Nodes.Count
  If tvUsers.Nodes.Item(a).Key = Key Then
    tvUsers.Nodes.Remove a
    Exit For
  End If
Next a
End Sub

Private Function ISubClass_WindowProc(ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim NH As NMHDR
Dim MF As MSGFILTER
Debug.Print "WindowProc!"
On Error Resume Next
'If hwnd <> rtfBuffer.hwnd Then Exit Function
' Get the NMHDR structure

'MoveMemory NH, ByVal lParam, Len(NH)

Debug.Print "MSG: " & MSG & " lParam: " & lParam & " wParam: " & wParam
Select Case NH.code
  Case EN_LINK
     ProcessLinkEvent lParam
End Select
End Function
Private Sub ProcessLinkEvent(ByVal lParam As Long)
Debug.Print "ProcessLinkEvent!"
Dim tLink As ENLINK
Dim tPoint As POINTL
Dim iButton As Integer
Dim iShift As Integer

   ' Get the ENLINK structure
'   MoveMemory tLink, ByVal lParam, LenB(tLink)

   With tLink

      ' Set the iButton parameter
      'If (.wParam And MK_LBUTTON) Then iButton = iButton Or vbLeftButton
      'If (.wParam And MK_MBUTTON) Then iButton = iButton Or vbMiddleButton
      'If (.wParam And MK_RBUTTON) Then iButton = iButton Or vbRightButton

      ' Set the iShift parameter
      'iShift = GetShiftMask()

   End With

   ' Get the mouse position
   GetCursorPos tPoint

   ' Convert the position to
   ' client coordinates
   ScreenToClient rtfBuffer.hWnd, tPoint

   ' Get the link range
   'Set oLinkRange = Range(tLink.chrg.cpMin, tLink.chrg.cpMax)
  Debug.Print "Start: " & tLink.chrg.cpMin & " End: " & tLink.chrg.cpMax
   ' Raise the event
   Select Case tLink.MSG

      Case WM_LBUTTONDBLCLK, WM_RBUTTONDBLCLK, WM_MBUTTONDBLCLK
         'not really needed

      Case WM_LBUTTONDOWN, WM_RBUTTONDOWN, WM_MBUTTONDOWN
         'shell the link!

      Case WM_LBUTTONUP, WM_RBUTTONUP, WM_MBUTTONUP
         'not really needed
      Case WM_MOUSEMOVE
         'not really needed
   End Select

End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    'TODO: Add more client commands
    If KeyAscii = 13 Then
        'If the text is not a command (prefixed with '/'), then just speak the text
        'normally. Otherwise, see which command it is, and execute it accordingly.
        Debug.Print "Enter Pressed"
        Debug.Print txtChat.Text
        txtChat.Text = Replace(txtChat.Text, vbCr, vbNullString)
        txtChat.Text = Replace(txtChat.Text, vbLf, vbNullString)
        txtChat.Text = Replace(txtChat.Text, vbCrLf, vbNullString)
        If Left$(txtChat.Text, 1) <> "/" Then
            'sckIRC.SendData "PRIVMSG " & Channel & ":" & txtChat.Text & vbCrLf
            sckIRC.SendData "PRIVMSG " & Channel & " :" & txtChat.Text & vbCrLf
            RTF_SetBold False
            RTF_SetColor QBColor(0)
            RTF_AddText ilIndent
            RTF_SetBold True
            RTF_AddText Nick & ": "
            RTF_SetBold False
            RTF_AddText txtChat.Text & vbCrLf
        Else
            If UCase$(Left$(txtChat.Text, 5)) = "/DBG1" Then
            Debug.Print rtfBuffer.TextRTF
            InternalDebug rtfBuffer.Text
            End If
            If LCase$(Left$(txtChat.Text, 4)) = "/me " Then 'It's an action
                txtChat.Text = Right$(txtChat.Text, Len(txtChat.Text) - 4)
                sckIRC.SendData "PRIVMSG " & Channel & " :" & Chr(1) & "ACTION " & txtChat.Text & Chr(1) & vbCrLf
            End If
            
            If LCase$(Left$(txtChat.Text, 6)) = "/join " Then 'JOIN
                txtChat.Text = Right$(txtChat.Text, Len(txtChat.Text) - 6)
                sckIRC.SendData "JOIN :" & txtChat.Text & vbCrLf
                
                Channel = txtChat.Text
            End If
            
            If LCase$(Left$(txtChat.Text, 6)) = "/nick " Then 'NICK
                txtChat.Text = Right$(txtChat.Text, Len(txtChat.Text) - 6)
                sckIRC.SendData "NICK :" & txtChat.Text & vbCrLf
                
                Nick = txtChat.Text
            End If
            
            If LCase$(Left$(txtChat.Text, 6)) = "/quit " Then 'QUIT
                If Right$(txtChat.Text, Len(txtChat.Text) - 6) = Null Then
                    sckIRC.SendData "Quit :" & vbCrLf
                Else
                    txtChat.Text = Right$(txtChat.Text, Len(txtChat.Text) - 6)
                    sckIRC.SendData "QUIT :" & txtChat.Text & vbCrLf
                End If
            End If
            
            If LCase$(Left$(txtChat.Text, 10)) = "/nickserv " Then 'NICKSERV
                txtChat.Text = Right$(txtChat.Text, Len(txtChat.Text) - 10)
                sckIRC.SendData "PRIVMSG nickserv :" & txtChat.Text & vbCrLf
            End If
            
            If LCase$(Left$(txtChat.Text, 10)) = "/chanserv " Then 'CHANSERV
                txtChat.Text = Right$(txtChat.Text, Len(txtChat.Text) - 10)
                sckIRC.SendData "PRIVMSG chanserv :" & txtChat.Text & vbCrLf
            End If
            
            If LCase$(Left$(txtChat.Text, 5)) = "/msg " Then 'MSG
                txtChat.Text = Right$(txtChat.Text, Len(txtChat.Text) - 5)
                Dim Count, SendTo, Sendtxt
                Count = 1
                While (Sendtxt <> Chr(32))
                    Sendtxt = Right(Left(txtChat.Text, Count), 1)
                    Count = Count + 1
                Wend
                SendTo = Left(txtChat.Text, Count - 1)
                sckIRC.SendData "PRIVMSG " & SendTo & ":" & Right(txtChat.Text, Len(txtChat.Text) - (Count - 1)) & vbCrLf
            End If
            
            If UCase$(Left$(txtChat.Text, 8)) = "/VERSION" Then
                RTF_SetColor QBColor(9)
                RTF_AddText ilIndent & "EvolvedIRC Version " & Version & vbCrLf
                RTF_SetColor QBColor(9)
                RTF_AddText ilIndent & "© 2004 Matthew Sporich and Contributors. For product information, please see http://evolved2go.ws4f.us/" & vbCrLf
                RTF_SetColor QBColor(9)
                RTF_AddText ilIndent & "EvolvedIRC is protected by the GNU General Public License. For more information, type /gpl." & vbCrLf
            End If
            
            If LCase$(Left$(txtChat.Text, 6)) = "/oper " Then 'OPER
                txtChat.Text = Right$(txtChat.Text, Len(txtChat.Text) - 6)
                sckIRC.SendData "OPER " & txtChat.Text & vbCrLf
            End If
        End If
        
        'txtChat.Text = "" 'Clear the textbox
        txtChat.Text = vbNullString
        KeyAscii = 0
        DoEvents
    End If
End Sub


Private Sub Form_Resize()
On Error Resume Next
    'width stuff
    shTop.Width = ScaleWidth
    tvUsers.Left = ScaleWidth - 0 - tvUsers.Width
    rtfBuffer.Width = ScaleWidth - tvUsers.Width - 100
    txtChat.Width = ScaleWidth - tvUsers.Width - 100

    'now adjust the height
    txtChat.Top = ScaleHeight - 0 - txtChat.Height
    rtfBuffer.Height = txtChat.Top - rtfBuffer.Top - 120
    tvUsers.Height = ScaleHeight - 0 - tvUsers.Top
    
End Sub
