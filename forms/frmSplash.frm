VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5790
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   8925
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   5790
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar pgb1 
      Height          =   135
      Left            =   0
      TabIndex        =   0
      Top             =   5640
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Timer timHideSplash 
      Interval        =   50
      Left            =   960
      Top             =   4920
   End
   Begin VB.Image imgEIRC 
      Appearance      =   0  'Flat
      Height          =   5625
      Left            =   0
      Picture         =   "frmSplash.frx":AD49
      Top             =   0
      Width           =   9000
   End
End
Attribute VB_Name = "frmSplash"
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
' $Id: frmSplash.frm,v 1.5 2005/01/03 06:13:38 dj_dark Exp $
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

Private Sub Form_Click()
Load frmMain
frmMain.Show
Unload Me
End Sub

'Private Sub Form_Load()
'    lblVersion.Caption = "Version: " & App.Major & "." & App.Minor & "." & App.Revision & " Build: 0001"
'    lblProductName.Caption = App.Title
'End Sub

Private Sub timHideSplash_Timer()
    If (pgb1.Value < 100) Then
        pgb1.Value = pgb1.Value + 1
    Else
        Load frmMain
        frmMain.Show
        Unload Me
    End If
End Sub

Private Sub imgEIRC_Click()
    Load frmMain
    frmMain.Show
    Unload Me
End Sub
