Attribute VB_Name = "CheckOnline"
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

Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Const ERROR_SUCCESS = 0&
Public Const APINULL = 0&
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public ReturnCode As Long

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal _
HKEY As Long) As Long

Declare Function RegOpenKey Lib "advapi32.dll" Alias _
"RegOpenKeyA" (ByVal HKEY As Long, ByVal lpSubKey As _
String, phkResult As Long) As Long

Declare Function RegQueryValueEx Lib "advapi32.dll" Alias _
"RegQueryValueExA" (ByVal HKEY As Long, ByVal lpValueName _
As String, ByVal lpReserved As Long, lpType As Long, _
lpData As Any, lpcbData As Long) As Long

Public Function ActiveConnection() As Boolean
Dim HKEY As Long
Dim lpSubKey As String
Dim phkResult As Long
Dim lpValueName As String
Dim lpReserved As Long
Dim lpType As Long
Dim lpData As Long
Dim lpcbData As Long
ActiveConnection = False
lpSubKey = "System\CurrentControlSet\Services\RemoteAccess"
ReturnCode = RegOpenKey(HKEY_LOCAL_MACHINE, lpSubKey, _
phkResult)

If ReturnCode = ERROR_SUCCESS Then
HKEY = phkResult
lpValueName = "Remote Connection"
lpReserved = APINULL
lpType = APINULL
lpData = APINULL
lpcbData = APINULL
ReturnCode = RegQueryValueEx(HKEY, lpValueName, _
lpReserved, lpType, ByVal lpData, lpcbData)
lpcbData = Len(lpData)
ReturnCode = RegQueryValueEx(HKEY, lpValueName, _
lpReserved, lpType, lpData, lpcbData)

If ReturnCode = ERROR_SUCCESS Then
If lpData = 0 Then
ActiveConnection = False
Else
ActiveConnection = True
End If
End If

RegCloseKey (HKEY)
End If

End Function

Public Sub Connect(strConnectName As String, blnPressConnect As Boolean)
Shell "rundll32.exe rnaui.dll,RnaDial " & strConnectName, vbNormalFocus
If blnPressConnect Then
DoEvents
SendKeys "{ENTER}", True
End If
End Sub

Public Sub Disconnect(strConnectName As String, blnPressDisconnect As Boolean)
Shell "rundll32.exe rnaui.dll,RnaDial " & strConnectName, 1
If blnPressDisconnect Then
    DoEvents
    SendKeys "{TAB} ", True
End If
End Sub

