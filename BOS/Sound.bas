Attribute VB_Name = "Sound"
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

     Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
     Public Const SND_SYNC = &H0 ' Don't return until sound ends (default).
     Public Const SND_ASYNC = &H1 ' Return immediately after the sound starts.
     Public Const SND_NODEFAULT = &H2 ' If the sound file is not found, do NOT play default sound.
     Public Const SND_MEMORY = &H4 ' Play a sound from a buffer in memory.
     Public Const SND_LOOP = &H8 ' Loop sound continuously (used with SND_ASYNC)
     Public Const SND_NOSTOP = &H10 ' Don't stop current sound to play another.

Public Sub s_Playsound(strName As String)
    a = GetSetting("BoS", "BInterface", "Skin", "BoS Standard")
    If a <> "BoS Standard" Then
        strName = App.path & "\skins\" & a & "\" & strName & ".wav"
    Else
        strName = App.path & "\" & strName & ".wav"
    End If
    sndPlaySound strName, SND_ASYNC Or SND_NODEFAULT
End Sub

Sub PlaySound(strName As String)
    strName = App.path & "\BAim\" & strName
    sndPlaySound strName, SND_ASYNC Or SND_NODEFAULT
End Sub
