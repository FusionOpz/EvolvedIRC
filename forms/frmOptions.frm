VERSION 5.00
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
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   7800
      TabIndex        =   2
      Top             =   5400
      Width           =   975
   End
   Begin VB.ListBox lstOptMenu 
      Appearance      =   0  'Flat
      Height          =   5490
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Height          =   5295
      Left            =   2520
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      Begin EvolvedIRC.pgMain pgMain 
         Height          =   4935
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   8705
      End
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
' $Id: frmOptions.frm,v 1.4 2004/10/22 03:56:44 dj_dark Exp $
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

    'iniFile = App.Path & "\options.ini"
    'LoadFileToTextbox
    
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
    
    Unload Me
    
    'Dim lpSectionName As String
    'Dim lpKeyName As String
    'Dim lpValue As String
    
    'lpSectionName = "Core"
    'lpKeyName = "Server" ' & "Port"
    'lpValue = txtServ.Text ' & txtPort.Text
    
    'Call ProfileSaveItem(lpSectionName, lpKeyName, lpValue, iniFile)
    
    'LoadFileToTextbox
        
End Sub
