VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "EvolvedIRD"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10230
   Icon            =   "frmIRC.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   10230
   StartUpPosition =   3  'Windows Default
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
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   4440
      Width           =   10215
   End
   Begin VB.TextBox txtBuffer 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   10215
   End
   Begin MSWinsockLib.Winsock sckIRC 
      Left            =   9120
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
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

Private Sub sckIRC_Connect()
    With sckIRC
        .SendData "NICK Wsvb_test" & vbCrLf
        .SendData "USER Wsvb_test " & sckIRC.LocalHostName & " " & _
            UCase(sckIRC.LocalHostName & ":" & sckIRC.LocalPort & "/0") & " :WinsockVB Test Client" & vbCrLf
        .SendData "JOIN #ignition-project" & vbCrLf
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
            sckIRC.SendData "PRIVMSG #ignition-project :" & txtChat.Text & vbCrLf
        Else
            'ME Command
            If LCase$(Left$(txtChat.Text, 4)) = "/me " Then 'It's an action
                txtChat.Text = Right$(txtChat.Text, Len(txtChat.Text) - 4)
                sckIRC.SendData "PRIVMSG #ignition-project :" & Chr(1) & "ACTION " & txtChat.Text & Chr(1) & vbCrLf
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
        End If
        
        txtChat.Text = "" 'Clear the textbox
    End If
End Sub
