VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8865
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fmMenu 
      Caption         =   "Menu"
      Height          =   5415
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   2415
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   5055
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   8916
         _Version        =   393217
         Style           =   7
         Appearance      =   1
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   255
      Left            =   5640
      TabIndex        =   3
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   6720
      TabIndex        =   2
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   255
      Left            =   7800
      TabIndex        =   1
      Top             =   5520
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   5415
      Left            =   2520
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      Begin VB.CheckBox chkIRCX 
         Caption         =   "Enable IRCX"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   2400
         Width           =   4335
      End
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   1920
         TabIndex        =   12
         Text            =   "EvolvedIRC User"
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtNick 
         Height          =   285
         Left            =   1920
         TabIndex        =   11
         Text            =   "Guest_##"
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   1920
         TabIndex        =   10
         Text            =   "6667"
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox txtServ 
         Height          =   285
         Left            =   1920
         TabIndex        =   9
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtRealnm 
         Height          =   285
         Left            =   1920
         TabIndex        =   8
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox txtDeChan 
         Height          =   285
         Left            =   1920
         TabIndex        =   7
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Label lblUser 
         Caption         =   "Username:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblNick 
         Caption         =   "Nickname:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblPort 
         Caption         =   "Port(Default is 6667):"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lblServ 
         Caption         =   "Server Address:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label lblRealnm 
         Caption         =   "Real Name:"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblDeChan 
         Caption         =   "Default Channel:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2040
         Width           =   1695
      End
   End
   Begin VB.Label lblBuildID 
      Caption         =   "BuildID:"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   5520
      Width           =   1935
   End
   Begin VB.Line Line1 
      X1              =   5520
      X2              =   1800
      Y1              =   5640
      Y2              =   5640
   End
End
Attribute VB_Name = "frmOptions"
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
' $Id: frmOptions.frm,v 1.7 2005/01/28 05:06:00 dj_dark Exp $
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

Option Explicit

Private inifile As String

Private Sub cmdApply_Click()
    'User Options
    WriteINI "userinfo", "nickname", txtNick.Text, App.Path + "\options.ini"
    WriteINI "userinfo", "username", txtUser.Text, App.Path + "\options.ini"
    WriteINI "userinfo", "realname", txtRealnm.Text, App.Path + "\options.ini"
    
    'Server Options
    WriteINI "server", "address", txtServ.Text, App.Path + "\options.ini"
    WriteINI "server", "port", txtPort.Text, App.Path + "\options.ini"
    WriteINI "server", "defaultchan", txtDeChan.Text, App.Path + "\options.ini"
    
    If chkIRCX.Value = "1" Then
        WriteINI "server", "IRCX", "1", App.Path + "\options.ini"
    ElseIf chkIRCX.Value = "2" Then
        WriteINI "server", "IRCX", "1", App.Path + "\options.ini"
    Else
        WriteINI "server", "IRCX", "0", App.Path + "\options.ini"
    End If
End Sub

Private Sub Form_Load()
    'User Options
    txtNick.Text = ReadINI("userinfo", "nickname", App.Path + "\options.ini")
    txtUser.Text = ReadINI("userinfo", "username", App.Path + "\options.ini")
    txtRealnm.Text = ReadINI("userinfo", "realname", App.Path + "\options.ini")

    'Server Options
    txtServ.Text = ReadINI("server", "address", App.Path + "\options.ini")
    txtPort.Text = ReadINI("server", "port", App.Path + "\options.ini")
    txtDeChan.Text = ReadINI("server", "defaultchan", App.Path + "\options.ini")
    chkIRCX.Value = ReadINI("server", "IRCX", App.Path + "\options.ini")
    
    'chkIRCX.Value = "2"
    chkIRCX.Enabled = False
    lblBuildID.Caption = "BuildID: " & BuildID
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
    'User Options
    WriteINI "userinfo", "nickname", txtNick.Text, App.Path + "\options.ini"
    WriteINI "userinfo", "username", txtUser.Text, App.Path + "\options.ini"
    WriteINI "userinfo", "realname", txtRealnm.Text, App.Path + "\options.ini"
    
    'Server Options
    WriteINI "server", "address", txtServ.Text, App.Path + "\options.ini"
    WriteINI "server", "port", txtPort.Text, App.Path + "\options.ini"
    WriteINI "server", "defaultchan", txtDeChan.Text, App.Path + "\options.ini"
    If chkIRCX.Value = "1" Then
        WriteINI "server", "IRCX", "1", App.Path + "\options.ini"
    Else
        WriteINI "server", "IRCX", "0", App.Path + "\options.ini"
    End If
    Unload Me
           
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

Private Sub Form_Unload(Cancel As Integer)
'Call frmMain.ReadINI
End Sub
