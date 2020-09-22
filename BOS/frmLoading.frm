VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLoading 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "frmLoading.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicTemp 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   120
      Picture         =   "frmLoading.frx":1042
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   5
      Top             =   2520
      Width           =   240
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2820
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label4 
      Caption         =   "Starting up..."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   420
      TabIndex        =   4
      Top             =   2520
      Width           =   3675
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   99.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   2235
      Left            =   1710
      TabIndex        =   3
      Top             =   150
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   99.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   2235
      Left            =   1740
      TabIndex        =   2
      Top             =   180
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   99.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   2235
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmLoading"
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

Dim loadedPrograms() As String
Dim loadedRegString As String

Private Sub Form_Load()
    TransLevel = Val(GetSetting("BoS", "BSystem", "Translucency", "55"))
    loadedRegString = GetSetting("BoS", "BSystem", "StartupPrograms", "")
    loadedPrograms = Split(loadedRegString, "<NewProg>")
    Me.Show
    Me.Refresh
    If UBound(loadedPrograms) = -1 Then
        ProgressBar1.Max = 90
    Else
        ProgressBar1.Max = 90 + (UBound(loadedPrograms) * 5) + 1
    End If
    Load frmTaskbar
    ProgressBar1.Value = 1
    Load frmBDesktopIcons
    ProgressBar1.Value = 90
    For i = 0 To UBound(loadedPrograms)
        Shell (loadedPrograms(i))
        ProgressBar1.Value = 91 + (i * 5)
        DrawIcon loadedPrograms(i)
        Label4.Caption = "Starting " & ExtractFileName(loadedPrograms(i))
    Next
    Me.Refresh
    If GetSetting("BoS", "BToday", "Enabled", "True") = False Then
        frmTaskbar.Show
        frmBDesktopIcons.Show
    Else
        frmStartup.Show
    End If
    Me.Hide
    If GetSetting("Bos", "BIM", "AutoLogin", "False") = "True" Then frmSignOn.cmdSignOn_Click
End Sub

Sub DrawIcon(path As String)
    hImgLarge& = SHGetFileInfo(path, 0&, shinfo, Len(shinfo), _
    BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
    PicTemp.Cls
    ImageList_Draw hImgLarge&, shinfo.iIcon, PicTemp.hdc, 0, 0, ILD_TRANSPARENT
    PicTemp.Refresh
End Sub

