VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "EvolvedIRC Codename ""Grasshopper"", Build: 0003"
   ClientHeight    =   6525
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
   ScaleHeight     =   6525
   ScaleWidth      =   9855
   Begin VB.ListBox lstUsers 
      Appearance      =   0  'Flat
      Height          =   6465
      ItemData        =   "frmMain.frx":2B82
      Left            =   8280
      List            =   "frmMain.frx":2B84
      TabIndex        =   2
      Top             =   0
      Width           =   1575
   End
   Begin RichTextLib.RichTextBox txtBuffer 
      Height          =   6135
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   10821
      _Version        =   393217
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":2B86
   End
   Begin MSWinsockLib.Winsock sckIRC 
      Left            =   8760
      Top             =   6480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtChat 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   6240
      Width           =   8175
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
'                     Website <http://myth.ws4f.us/>
'
' $Id: frmMain.frm,v 1.4 2004/10/22 03:56:44 dj_dark Exp $
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
'Start formattig the recived code,
'Start coding Options System,
'Make use of the User list box,
'Commit To CVS!

Option Explicit

'TODO: Write Options dialog to set Server address, server port, nickname and username,
'Also get it to either save the optios to a Var or into a File.
'TODO: now get to coding the dialog so it'll remember the settings.

Private Sub Form_Load()
'    Call Connect
End Sub

Private Sub Connect()
    With sckIRC
        .RemoteHost = "irc.freenode.net" 'The IRC server
        .RemotePort = 6667 'Connect on port 6667
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
    With sckIRC
        .SendData "QUIT : Time for me to go l8r" & vbCrLf
        .Close
        .RemoteHost = "irc.freenode.net" 'The IRC server
        .RemotePort = 6667 'Connect on port 6667
        .Connect
    End With
End Sub

Private Sub mnuOpt_Click()
    frmOptions.Show
End Sub

Private Sub mnuWbo_Click()
    With sckIRC
        .SendData "QUIT : Time for me to go l8r" & vbCrLf
        .Close
        .RemoteHost = "irc.winbeta.org" 'The IRC server
        .RemotePort = 6667 'Connect on port 6667
        .Connect
    End With
End Sub

'TODO: Write Options dialog to set Server address, server port, nickname and username,
'Also get it to either save the optios to a Var or into a File.
'TODO: now get to coding the dialog so it'll remember the settings.
Private Sub sckIRC_Connect()
    With sckIRC
        .SendData "NICK Guest_##" & vbCrLf
        .SendData "USER EvolvedIRC User " & sckIRC.LocalHostName & " " & _
            UCase(sckIRC.LocalHostName & ":" & sckIRC.LocalPort & "/0") & " :EvolvedIRC Pre-Alpha Client" & vbCrLf
        .SendData "JOIN #lobby" & vbCrLf
    End With
End Sub

Private Sub sckIRC_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    
    Dim sRecv As String
    
    sckIRC.GetData sRecv 'Put the data recieved into the string
    

    
    'Play ping pong with the server
    If Split(sRecv, " ")(0) = "PING" Then
        sckIRC.SendData "PONG " & Split(sRecv, " ")(1)
    End If
    
    'Update the buffer
    txtBuffer.Text = txtBuffer.Text & sRecv & vbCrLf
    txtBuffer.SelStart = Len(txtBuffer.Text)
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    'TODO: Add more client commands
    If KeyAscii = 13 Then
        'If the text is not a command (prefixed with '/'), then just speak the text
        'normally. Otherwise, see which command it is, and execute it accordingly.
        If Left$(txtChat.Text, 1) <> "/" Then
            sckIRC.SendData "PRIVMSG #lobby :" & txtChat.Text & vbCrLf
        Else
            If LCase$(Left$(txtChat.Text, 4)) = "/me " Then 'It's an action
                txtChat.Text = Right$(txtChat.Text, Len(txtChat.Text) - 4)
                sckIRC.SendData "PRIVMSG #lobby :" & Chr(1) & "ACTION " & txtChat.Text & Chr(1) & vbCrLf
            End If
            
            If LCase$(Left$(txtChat.Text, 6)) = "/join " Then 'JOIN
                txtChat.Text = Right$(txtChat.Text, Len(txtChat.Text) - 6)
                sckIRC.SendData "JOIN :" & txtChat.Text & vbCrLf
            End If
            
            If LCase$(Left$(txtChat.Text, 6)) = "/nick " Then 'NICK
                txtChat.Text = Right$(txtChat.Text, Len(txtChat.Text) - 6)
                sckIRC.SendData "NICK :" & txtChat.Text & vbCrLf
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
            
            'If LCase$(Left$(txtChat.Text, 8)) = "/version " Then 'Version
            '    txtChat.Text = Right$(txtChat.Text, Len(txtChat.Text) - 8)
            '    Dim Count, SendTo
            '    Count = 1
            '    SendTo = Left$(txtChat.Text, Count - 1)
            '    sckIRC.SendData "VERSION " & SendTo & vbCrLf
            'End If
        End If
        
        txtChat.Text = "" 'Clear the textbox
    End If
End Sub


Private Sub Form_Resize()
On Error Resume Next
    'width stuff
    lstUsers.Left = ScaleWidth - 0 - lstUsers.Width
    txtBuffer.Width = ScaleWidth - lstUsers.Width - 100
    txtChat.Width = ScaleWidth - lstUsers.Width - 100

    'now adjust the height
    txtChat.Top = ScaleHeight - 0 - txtChat.Height
    txtBuffer.Height = txtChat.Top - txtBuffer.Top - 120
    lstUsers.Height = ScaleHeight - 0 - lstUsers.Top
End Sub
