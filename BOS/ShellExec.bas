Attribute VB_Name = "ShellExec"
'Copyright (C) 2000 BSoft
'
'This program is free software; you can redistribute it and/or
'modify it under the terms of the GNU General Public License
'as published by the Free Software Foundation; either version 2
'of the License, or (at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Function ExtractFileName(ByVal strPath As String) As String
    ' StrReverse is only working in VB6
    strPath = StrReverse(strPath)
    strPath = Left(strPath, InStr(strPath, "\") - 1)
    ExtractFileName = StrReverse(strPath)
End Function

Public Function ExtractPath(ByVal strPath As String)
    strtmp = StrReverse(strPath)
    a = Len(strPath) - InStr(strtmp, "\")
    strPath = Left(strPath, a)
    ExtractPath = strPath
End Function


Function AddASlash(ByVal path As String)
If Right(path, 1) = "\" Then
    AddASlash = path
Else
    AddASlash = path & "\"
End If
End Function

Public Function FileExists(strPath As String) As Integer
    FileExists = Not (Dir(strPath) = "")
End Function
