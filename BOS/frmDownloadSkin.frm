VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmDownloadSkin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Downloading Skin..."
   ClientHeight    =   1005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet inetDownloadSkin 
      Left            =   4920
      Top             =   1380
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar prgPercentDone 
      Height          =   255
      Left            =   900
      TabIndex        =   2
      Top             =   660
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   22
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmDownloadSkin.frx":0000
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblCurrentFile 
      Caption         =   "File [file number] of 22"
      Height          =   255
      Left            =   900
      TabIndex        =   1
      Top             =   360
      Width           =   4395
   End
   Begin VB.Label lblSkinName 
      Caption         =   "Downloading [skin name]"
      Height          =   255
      Left            =   900
      TabIndex        =   0
      Top             =   120
      Width           =   4395
   End
End
Attribute VB_Name = "frmDownloadSkin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Dim SkinPath As String
Dim SavePath As String

Public Sub InstallSkin(SkinName As String)
    lblSkinName.Caption = "Downloading " & SkinName
    Me.Show
    Me.Refresh
    
    MkDir App.path & "\skins\" & SkinName
    SkinPath = "http://thunder.prohosting.com/~pikared/skins/" & Replace(SkinName, " ", "%20") & "/"
    SavePath = App.path & "\skins\" & SkinName & "\"
    
    Downlaod SkinPath, "ClockFgColor.txt", SavePath
    Downlaod SkinPath, "DownArrowDown.bmp", SavePath
    Downlaod SkinPath, "DownArrowUp.bmp", SavePath
    Downlaod SkinPath, "hover.wav", SavePath
    Downlaod SkinPath, "IconBGColor.txt", SavePath
    Downlaod SkinPath, "IconFGColor.txt", SavePath
    Downlaod SkinPath, "IconShadowColor.txt", SavePath
    Downlaod SkinPath, "open.wav", SavePath
    Downlaod SkinPath, "ProgramDown.bmp", SavePath
    Downlaod SkinPath, "ProgramNone.bmp", SavePath
    Downlaod SkinPath, "ProgramUp.bmp", SavePath
    Downlaod SkinPath, "select.wav", SavePath
    Downlaod SkinPath, "Seperator.bmp", SavePath
    Downlaod SkinPath, "StartButtonDown.bmp", SavePath
    Downlaod SkinPath, "StartButtonUp.bmp", SavePath
    Downlaod SkinPath, "SystemTray.bmp", SavePath
    Downlaod SkinPath, "TaskbarBg.bmp", SavePath
    Downlaod SkinPath, "TaskbarFgColor.txt", SavePath
    Downlaod SkinPath, "UpArrowDown.bmp", SavePath
    Downlaod SkinPath, "UpArrowUp.bmp", SavePath
    Downlaod SkinPath, "PSCUp.bmp", SavePath
    Downlaod SkinPath, "PSCDown.bmp", SavePath
        
    Me.Hide
    Unload frmGetSkin
    Unload frmBSettings
    frmBSettings.Show
    Unload Me
    
    
End Sub


Public Sub Downlaod(Location As String, Filename As String, DirToSaveAt As String)
    lblCurrentFile.Caption = "Downloading " & Filename & " - file " & prgPercentDone.Value + 1 & " of 22"
    Dim mocha As String
    mocha = Location & Filename
    Dim bData() As Byte
    Dim intFile As Integer
    intFile = FreeFile()
    bData() = inetDownloadSkin.OpenURL(mocha, icByteArray)
    Open DirToSaveAt & "\" & Filename For Binary Access Write _
    As #intFile
    Put #intFile, , bData()
    Close #intFile
    prgPercentDone.Value = prgPercentDone.Value + 1
End Sub

Private Sub Form_Activate()
HideStartMenu
End Sub
