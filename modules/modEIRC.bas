Attribute VB_Name = "modEIRC"
'EvolvedIRC is (C)  Matthew Sporich.
'----------------------------------------------------------
'You must include this notice in any modifications you make. You must additionally
'follow the GPL's provisions for sourcecode distribution and binary distribution.
'If you are not familiar with the GPL, please read LICENSE.TXT.
'(you are welcome to add a "Based On" line above this notice, but this notice must
'remain intact!)
'Released under the GNU General Public License
'Contact information: Matthew Sporich (DJ_Dark) <djdark@gmail.com>
'                     Website <http://quantump.net/>
'
' $Id: modMain.bas,v 1.3 2005/03/02 23:47:25 dj_dark Exp $
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

'`mIRC format' Constants
'These characters are used by mIRC.
'CTCP: To indicate a CTCP request or reply the hole message is included in two of these characters
'Sample Request:
'VERSION
Public Const MIRC_CTCP As String = "" 'CTCP Request or Reply = ChrW$(1)
'For a message to appear in bold, it is included in the bold characters:
'This is some bold text
Public Const MIRC_BOLD As String = "" 'Bold = ChrW$(2)
'Some colored text is included in two color characters, the first followed by one or two numbers seperated by comma:
'1,2This is some colored text
'The first number is the forecolor whereas the second is the background color.
'The numbers are between 0 and 15.
Public Const MIRC_COLOR As String = "" 'Color = ChrW$(3)
'For a message to appear underlined, it is included in the underline characters:
'This is some underlined text
Public Const MIRC_UNDERLINE As String = "" 'Underline = ChrW$(31)
Public Const MIRC_ITALIC As String = ""
