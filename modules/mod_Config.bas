Attribute VB_Name = "mod_Config"
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
' $Id: mod_Config.bas,v 1.6 2005/03/02 23:47:25 dj_dark Exp $
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

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WitePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Sub ProfileSaveItem(lpSectionName As String, lpKeyName As String, lpValue As String, iniFile As String)

'This function saves the passed value to the file,
'under the section and key names specified.
'If the ini file does not exist, it is created.
'If the section does not exist, it is created.
'If the key name does not exist, it is created.
'If the key name exists, its value is replaced.

    Call WritePrivateProfileString(lpSectionName, lpKeyName, lpValue, iniFile)
    
End Sub

Public Function ProfileGetItem(lpSection As String, lpKeyName As String, lpValue As String, iniFile As String) As String
'Retrieves a value from an ini file corresponding
'to the section and key name passed.

    Dim success As Long
    Dim nSize As Long
    Dim ret As String
    
    'call the API with the parameters passed.
    'The return value is the length of the string
    'in ret, including the terminating null. If a
    'default calue was passed, and the section or
    'key name are not in the file, that value is
    'returned. If no default value was passed (""),
    'then success will = 0 if not found.
    
    'Pad a string large enough to hold the data.
    ret = Space$(2048)
    nSize = Len(ret)
    
    success = GetPrivateProfileString(lpSection, lpKeyName, lpValue, iniFile)
    
    If success Then
        ProfileGetItem = Left$(ret, success)
    End If
       
End Function

Public Sub ProfileDeleteItem(lpSection As String, lpKeyName As String, iniFile As String)
'this call will remove the keynames and it's
'corresponding value from the section specified
'in lpSectionName. This is accomplished by passing
'vbNullString as the lpValue parameter. For example,
'assuming that an ini file had:
' [Colours]
' Colour1=Red
' Colour2=Blue
' Colour3=Green
'
'and this sub was caller passing "Colour2"
'as lpKeyName, the resulting ini file
'would contain:
' [Colours]
' Colour1=Red
' Colour3=Green

    Call WritePrivateProfileString(lpSectionName, lpKeyName, vbNullString, iniFile)

End Sub

Public Sub ProfileDeleteSection(lpSection As String, iniFile As String)
'this call will remove the entire section
'corresponding to lpSectionName. This is
'accomplished by passing vbNullString
'as both the lpKeyName and lpValue parameters.
'For example, assuming that an ini file had:
' [Colours]
' Colour1=Red
' Colour2=Blue
' Colour3=Green
'
'and this sub was called passing "Colours
'as lpSectionName, the resulting Colours
'section in the ini file would be deleted.

    Call WritePrivateProfileString(lpSectionName, vbNullString, vbNullString, iniFile)
End Sub
