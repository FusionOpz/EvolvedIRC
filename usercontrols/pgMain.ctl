VERSION 5.00
Begin VB.UserControl pgMain 
   ClientHeight    =   5280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6225
   ScaleHeight     =   5280
   ScaleWidth      =   6225
   Begin VB.TextBox txtDeChan 
      Height          =   285
      Left            =   1920
      TabIndex        =   11
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox txtRealnm 
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox txtServ 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Text            =   "6667"
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox txtNick 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Text            =   "Guest_##"
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Text            =   "EvolvedIRC User"
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label lblDeChan 
      Caption         =   "Default Channel:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label lblRealnm 
      Caption         =   "Real Name:"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblServ 
      Caption         =   "Server Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblPort 
      Caption         =   "Port(Default is 6667):"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblNick 
      Caption         =   "Nickname:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblUser 
      Caption         =   "Username:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "pgMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub UserControl_Initialize()
    'User Options
    txtNick.Text = ReadINI("userinfo", "nickname", App.Path + "\option.ini")
    txtUser.Text = ReadINI("userinfo", "username", App.Path + "\option.ini")
    txtRealnm.Text = ReadINI("userinfo", "realname", App.Path + "\option.ini")
    
    'Server Options
    txtServ.Text = ReadINI("server", "address", App.Path + "\option.ini")
    txtPort.Text = ReadINI("server", "port", App.Path + "\option.ini")
    txtDeChan.Text = ReadINI("server", "defaultchan", App.Path + "\option.ini")
End Sub
