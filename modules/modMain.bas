Attribute VB_Name = "modMain"
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
' $Id: modMain.bas,v 1.2 2005/03/02 00:55:02 dj_dark Exp $
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

Public Server_AuthNotices As Boolean

'some constants
Public Const BuildID As String = "20050103-cvs"

'indent constant
'vbTab is too huge, but somehow you can set tab lengths in the Rich Text box
'Public Const ilIndent As String = "   "
Public Const ilIndent As String = ""

'sendmessage API
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'cursor stuff
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTL) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTL) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTL) As Long

'Windows Messages Constants
Public Const WM_SETFOCUS = &H7
Public Const WM_MOUSEACTIVATE = &H21
Public Const WM_SETFONT = &H30
Public Const WM_NOTIFY = &H4E
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_CHAR = &H102
Public Const WM_COMMAND = &H111
Public Const WM_HSCROLL = &H114
Public Const WM_VSCROLL = &H115
Public Const WM_MOUSEFIRST = &H200
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MBUTTONDBLCLK = &H208
Public Const WM_MOUSELAST = &H209
Public Const WM_USER = &H400

' Notification masks
Public Const ENM_NONE = &H0
Public Const ENM_CHANGE = &H1
Public Const ENM_UPDATE = &H2
Public Const ENM_SCROLL = &H4
Public Const ENM_SCROLLEVENTS = &H8
Public Const ENM_DRAGDROPDONE = &H10
Public Const ENM_PARAGRAPHEXPANDED = &H20
Public Const ENM_KEYEVENTS = &H10000
Public Const ENM_MOUSEEVENTS = &H20000
Public Const ENM_REQUESTRESIZE = &H40000
Public Const ENM_SELCHANGE = &H80000
Public Const ENM_DROPFILES = &H100000
Public Const ENM_PROTECTED = &H200000
Public Const ENM_CORRECTTEXT = &H400000
Public Const ENM_LANGCHANGE = &H1000000
Public Const ENM_OBJECTPOSITIONS = &H2000000
Public Const ENM_LINK = &H4000000

' RichEdit messages
Public Const EM_GETRECT = &HB2
Public Const EM_SETRECT = &HB3
Public Const EM_GETMODIFY = &HB8&
Public Const EM_SETMODIFY = &HB9&
Public Const EM_GETLINECOUNT = &HBA&
Public Const EM_LINEINDEX = &HBB&
Public Const EM_CANUNDO = &HC6&
Public Const EM_EMPTYUNDOBUFFER = &HCD&
Public Const EM_GETLIMITTEXT = WM_USER + 37
Public Const EM_CANPASTE = WM_USER + 50
Public Const EM_EXLIMITTEXT = WM_USER + 53
Public Const EM_EXLINEFROMCHAR = WM_USER + 54
Public Const EM_FORMATRANGE = WM_USER + 57
Public Const EM_GETCHARFORMAT = WM_USER + 58
Public Const EM_GETOLEINTERFACE = WM_USER + 60
Public Const EM_SETBKGNDCOLOR = WM_USER + 67
Public Const EM_SETCHARFORMAT = WM_USER + 68
Public Const EM_SETEVENTMASK = WM_USER + 69
Public Const EM_SETOLECALLBACK = WM_USER + 70
Public Const EM_SETTARGETDEVICE = WM_USER + 72
Public Const EM_SETUNDOLIMIT = WM_USER + 82
Public Const EM_CANREDO = WM_USER + 85
Public Const EM_GETUNDONAME = WM_USER + 86
Public Const EM_GETREDONAME = WM_USER + 87
Public Const EM_AUTOURLDETECT = WM_USER + 91
Public Const EM_GETAUTOURLDETECT = WM_USER + 92
Public Const EM_SETEDITSTYLE = WM_USER + 204
Public Const EM_GETEDITSTYLE = WM_USER + 205
Public Const EM_OUTLINE = WM_USER + 220
Public Const EM_GETZOOM = WM_USER + 224
Public Const EM_SETZOOM = WM_USER + 225

Public Const EN_LINK = &H70B&

'some types
Type NMHDR
   hwndFrom As Long
   idFrom As Long
   code As Long
End Type

Type CHARRANGE
   cpMin As Long
   cpMax As Long
End Type

Type ENLINK
   NMHDR As NMHDR
   MSG As Integer
   wParam As Long
   lParam As Long
   chrg As CHARRANGE
End Type

Type MSGFILTER
   NMHDR As NMHDR
   MSG As Long
   wParam As Long
   lParam As Long
End Type

Type POINTL
    x As Long
    y As Long
End Type

Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Public Enum ilBulletTypes
  Arrow = 1
  Dot = 2
  Info = 3
End Enum

Sub Main()
On Error Resume Next
Load frmMSplash
frmSplash.Show
End Sub

Public Sub RTF_SetColor(lngColor As Long)
On Error Resume Next
frmMain.rtfBuffer.SelStart = Len(frmMain.rtfBuffer.Text)
frmMain.rtfBuffer.SelColor = lngColor
End Sub
Public Sub RTF_SetBold(blnBold As Boolean)
On Error Resume Next
frmMain.rtfBuffer.SelStart = Len(frmMain.rtfBuffer.Text)
frmMain.rtfBuffer.SelBold = blnBold
End Sub
Public Sub RTF_AddText(strText As String)
On Error Resume Next
frmMain.rtfBuffer.SelStart = Len(frmMain.rtfBuffer.Text)
frmMain.rtfBuffer.SelText = strText
End Sub
Public Sub InternalDebug(strText As String)
Dim F As Long
F = FreeFile
Open App.Path & "\debug.log" For Append As F
Print #F, "[" & Now & "] " & strText
Close #F
Debug.Print strText
End Sub
Public Sub RTF_AutoURLDetect(blnAutoURLDetect As Boolean)
SendMessage frmMain.rtfBuffer.hWnd, EM_AUTOURLDETECT, Abs(blnAutoURLDetect), ByVal 0&
End Sub
Public Sub RTF_AddBullet(BulletType As ilBulletTypes, lngColor As Long)
If BulletType = Dot Then
  frmMain.rtfBuffer.SelFontName = "Wingdings"
  RTF_SetColor lngColor
  frmMain.rtfBuffer.SelStart = Len(frmMain.rtfBuffer.Text)
  RTF_AddText "w"
  frmMain.rtfBuffer.SelFontName = "Tahoma"
ElseIf BulletType = Arrow Then
  frmMain.rtfBuffer.SelFontName = "Webdings"
  RTF_SetColor lngColor
  RTF_AddText "4"
  frmMain.rtfBuffer.SelFontName = "Tahoma"
End If
End Sub

Public Sub RTF_SetViewRect(Optional Left, Optional Top, Optional Right, Optional Bottom)
Dim R As RECT
   ' Get the current rectangle
   RTF_GetViewRect R.Left, R.Top, R.Right, R.Bottom

   ' Set the new values
   If Not IsMissing(Left) Then R.Left = CLng(Left)
   If Not IsMissing(Top) Then R.Top = CLng(Top)
   If Not IsMissing(Right) Then R.Right = CLng(Right)
   If Not IsMissing(Bottom) Then R.Bottom = CLng(Bottom)

   ' Set the new rectangle
   SendMessage frmMain.rtfBuffer.hWnd, EM_SETRECT, 0, R
End Sub

Public Sub RTF_GetViewRect(Optional Left As Long, Optional Top As Long, Optional Right As Long, Optional Bottom As Long)
Dim tRECT As RECT
   ' Get the rectangle
   SendMessage frmMain.rtfBuffer.hWnd, EM_GETRECT, 0, tRECT

   With tRECT
      Left = .Left
      Top = .Top
      Right = .Right
      Bottom = .Bottom
      
      Debug.Print "Current ViewRect: L=" & .Left & " T=" & .Top & " R=" & .Right & " B=" & .Bottom
   End With
End Sub

Public Sub RTF_Indent()
frmMain.rtfBuffer.SelStart = Len(frmMain.rtfBuffer.Text)
frmMain.rtfBuffer.SelText = "   "
frmMain.rtfBuffer.SelStart = Len(frmMain.rtfBuffer.Text)
End Sub


