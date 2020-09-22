VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBStartupItems 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Startup Programs"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmBStartupItems.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Remove"
      Height          =   315
      Left            =   1620
      TabIndex        =   5
      Top             =   2820
      Width           =   1395
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add"
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   2820
      Width           =   1335
   End
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   2880
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   3
      Top             =   900
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2040
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   315
      Left            =   3180
      TabIndex        =   2
      Top             =   2820
      Width           =   1395
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2055
      Left            =   60
      TabIndex        =   0
      Top             =   660
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   3625
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   7805
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Add or remove programs to be run at startup"
      Height          =   315
      Left            =   660
      TabIndex        =   1
      Top             =   180
      Width           =   3315
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   7
      Left            =   60
      Picture         =   "frmBStartupItems.frx":0442
      Top             =   60
      Width           =   480
   End
End
Attribute VB_Name = "frmBStartupItems"
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

Private Sub Command1_Click()
        For i = 1 To ListView1.ListItems.Count
            a = a & ListView1.ListItems(i) & "<NewProg>"
        Next
        If Len(a) > 0 Then a = Left(a, Len(a) - 9)
        SaveSetting "BoS", "BSystem", "StartupPrograms", a
        Unload Me
End Sub

Private Sub Command2_Click()
        ListView1.ListItems.Add
        Set ListView1.SelectedItem = ListView1.ListItems(ListView1.ListItems.Count)
        ListView1.SetFocus
        ListView1.StartLabelEdit
End Sub

Private Sub Command3_Click()
    ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
    If ListView1.ListItems.Count < 1 Then Command3.Enabled = False
End Sub

Private Sub Form_Load()
    loadedRegString = GetSetting("BoS", "BSystem", "StartupPrograms", "")
    loadedPrograms = Split(loadedRegString, "<NewProg>")
    For i = 0 To UBound(loadedPrograms)
        DrawIcon (loadedPrograms(i))
        ListView1.ListItems.Add , , loadedPrograms(i), i + 1, i + 1
    Next
    If ListView1.ListItems.Count < 1 Then Command3.Enabled = False
End Sub

Private Sub Form_Activate()
HideStartMenu
End Sub

Sub DrawIcon(path)
    hImgLarge& = SHGetFileInfo(path, 0&, shinfo, Len(shinfo), _
    BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
    PicTemp.Cls
    ImageList_Draw hImgLarge&, shinfo.iIcon, PicTemp.hdc, 0, 0, ILD_TRANSPARENT
    ImageList1.ListImages.Add , , PicTemp.Image
End Sub

Private Sub ListView1_AfterLabelEdit(Cancel As Integer, NewString As String)
If FileExists(NewString) Then
    DrawIcon (NewString)
    ListView1.SelectedItem.SmallIcon = ImageList1.ListImages.Count
Else
    Beep
    ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
End If
End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)
If ListView1.SelectedItem.Index <> ListView1.ListItems.Count Then
    Cancel = True
End If
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Command3.Enabled = True
End Sub

