VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4515
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
   ScaleHeight     =   1890
   ScaleWidth      =   4515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   255
      Left            =   1680
      TabIndex        =   10
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Text            =   "EvolvedIRC User"
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox txtNick 
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Text            =   "Guest_##"
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Text            =   "6667"
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox txtServ 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblUser 
      Caption         =   "Username:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblNick 
      Caption         =   "Nickname:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblPort 
      Caption         =   "Port(Default is 6667):"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblServ 
      Caption         =   "Server Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
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
'                     Website <http://myth.ws4f.us/>
'
' $Id: frmOptions.frm,v 1.2 2004/09/07 20:31:10 dj_dark Exp $
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

Private iniFile As String
Private Sub Form_Load()

    iniFile = App.Path & "\options.ini"
    LoadFileToTextbox
    
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
Unload Me
End Sub
