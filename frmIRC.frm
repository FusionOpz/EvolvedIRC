VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Caption         =   "EvolvedIRC - Connected to irc.FreeNode.Net Channel #Evolved-IRC"
   ClientHeight    =   4845
   ClientLeft      =   165
   ClientTop       =   750
   ClientWidth     =   10215
   Icon            =   "frmIRC.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox txtBuffer 
      Height          =   4455
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   7858
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmIRC.frx":058A
   End
   Begin VB.TextBox txtChat 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   4440
      Width           =   10215
   End
   Begin MSWinsockLib.Winsock sckIRC 
      Left            =   8640
      Top             =   4200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mFile 
      Caption         =   "File"
      NegotiatePosition=   3  'Right
      Begin VB.Menu mClose 
         Caption         =   "Close"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu mAbout 
      Caption         =   "About"
      NegotiatePosition=   1  'Left
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    With sckIRC
        .RemoteHost = "irc.freenode.net" 'The IRC server
        .RemotePort = 6667 'Connect on port 6667
        .Connect
    End With
    Call Form_Resize
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    txtBuffer.Width = ScaleWidth
    txtBuffer.Top = 0
    txtBuffer.Height = ScaleHeight - 285
    txtChat.Width = ScaleWidth
    txtChat.Top = txtBuffer.Height
End Sub


Private Sub Form_Unload(Cancel As Integer)
    'sckIRC.SendData "QUIT : Alpha Testing EvolvedIRC(http://myth.5ers.com" & vbCrLf
    sckIRC.Close
End Sub

Private Sub mAbout_Click()
Load frmAbout
frmAbout.Show
End Sub

Private Sub mClose_Click()
    Unload Me
    'Call Form_Unload
End Sub

Private Sub sckIRC_Connect()
    With sckIRC
        .SendData "NICK EvolvedIRC|Test" & vbCrLf
        .SendData "USER EvolvedIRC " & sckIRC.LocalHostName & " " & _
            UCase(sckIRC.LocalHostName & ":" & sckIRC.LocalPort & "/0") & " :EvolvedIRC Test Client" & vbCrLf
        .SendData "JOIN #Evolved-IRC" & vbCrLf
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
    
    If KeyAscii = 13 Then
        'If the text is not a command (prefixed with '/'), then just speak the text
        'normally. Otherwise, see which command it is, and execute it accordingly.
        If Left$(txtChat.Text, 1) <> "/" Then
            sckIRC.SendData "PRIVMSG #Evolved-IRC :" & txtChat.Text & vbCrLf
        Else
            'ME Command
            If LCase$(Left$(txtChat.Text, 4)) = "/me " Then 'It's an action
                txtChat.Text = Right$(txtChat.Text, Len(txtChat.Text) - 4)
                sckIRC.SendData "PRIVMSG #Evolved-IRC :" & Chr(1) & "ACTION " & txtChat.Text & Chr(1) & vbCrLf
            End If
        'Else
            'NICK Command
            If LCase$(Left$(txtChat.Text, 6)) = "/nick " Then 'It's to change your current nickname
                txtChat.Text = Right$(txtChat.Text, Len(txtChat.Text) - 6)
                sckIRC.SendData "NICK :" & Replace(txtChat.Text, "/nick ", "") & vbCrLf
            End If
        'Else
            'JOIN Command
            If LCase$(Left$(txtChat.Text, 6)) = "/join " Then 'It's to change the current channel your in
                txtChat.Text = Right$(txtChat.Text, Len(txtChat.Text) - 6)
                sckIRC.SendData "JOIN :" & Replace(txtChat.Text, "/join ", "") & vbCrLf
            End If
        'Else
            'MODE Command
            If LCase$(Left$(txtChat.Text, 6)) = "/mode " Then 'It's to change your current mode
                txtChat.Text = Right$(txtChat.Text, Len(txtChat.Text) - 6)
                sckIRC.SendData "MODE " & Replace(txtChat.Text, "/mode ", "MODE") & vbCrLf
            End If
        'Else
            'MSG Command
'            If LCase$(Left$(txtChat.Text, 6)) = "/msg " Then 'It's to change your current mode
'                txtChat.Text = Right$(txtChat.Text, Len(txtChat.Text) - 6)
'                sckIRC.SendData "PRIVMSG #Evolved-IRC :" & Chr(1) & "MSG " & txtChat.Text & Chr(1) & vbCrLf
'            End If
        End If
        
        txtChat.Text = "" 'Clear the textbox
    End If
End Sub
