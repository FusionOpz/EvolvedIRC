Attribute VB_Name = "modConfig"
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
' $Id: modConfig.bas,v 1.2 2005/03/02 00:55:02 dj_dark Exp $
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
'Public NName As String
'Public UName As String
'Public
'Public Server_AuthNotices As Boolean
'Public Server_AuthNotices As Boolean
'Public Server_AuthNotices As Boolean

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long
    
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
    (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpString As Any, _
    ByVal lpFileName As String) As Long
    
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
    
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

