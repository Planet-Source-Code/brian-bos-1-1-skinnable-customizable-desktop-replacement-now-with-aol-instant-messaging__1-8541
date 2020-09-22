VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBstart 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmBstart.frx":0000
   ScaleHeight     =   249
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBottomBorder 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   375
      Picture         =   "frmBstart.frx":4A2E
      ScaleHeight     =   150
      ScaleWidth      =   2550
      TabIndex        =   8
      Top             =   3585
      Visible         =   0   'False
      Width           =   2550
   End
   Begin VB.PictureBox picRightBorder 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   2925
      Picture         =   "frmBstart.frx":5E70
      ScaleHeight     =   3735
      ScaleWidth      =   75
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.PictureBox picButton 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   4
      Left            =   375
      Picture         =   "frmBstart.frx":6E52
      ScaleHeight     =   600
      ScaleWidth      =   2550
      TabIndex        =   6
      Top             =   600
      Width           =   2550
   End
   Begin VB.PictureBox picButton 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   5
      Left            =   375
      Picture         =   "frmBstart.frx":BE94
      ScaleHeight     =   600
      ScaleWidth      =   2550
      TabIndex        =   5
      Top             =   0
      Width           =   2550
   End
   Begin VB.PictureBox picButton 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   3
      Left            =   375
      Picture         =   "frmBstart.frx":10ED6
      ScaleHeight     =   600
      ScaleWidth      =   2550
      TabIndex        =   4
      Top             =   1200
      Width           =   2550
   End
   Begin VB.PictureBox picButton 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   2
      Left            =   375
      Picture         =   "frmBstart.frx":15F18
      ScaleHeight     =   600
      ScaleWidth      =   2550
      TabIndex        =   3
      Top             =   1800
      Width           =   2550
   End
   Begin MSComctlLib.ImageList imagesdown 
      Left            =   1200
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   170
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBstart.frx":1AF5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBstart.frx":1FFAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBstart.frx":25002
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBstart.frx":2A056
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBstart.frx":2F0AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBstart.frx":340FE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picButton 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   1
      Left            =   375
      Picture         =   "frmBstart.frx":39152
      ScaleHeight     =   600
      ScaleWidth      =   2550
      TabIndex        =   2
      Top             =   2400
      Width           =   2550
   End
   Begin MSComctlLib.ImageList imageshover 
      Left            =   540
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   170
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBstart.frx":3E194
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBstart.frx":431E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBstart.frx":4823C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBstart.frx":4D290
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBstart.frx":522E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBstart.frx":57338
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList images 
      Left            =   1860
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   170
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBstart.frx":5C38C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBstart.frx":613E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBstart.frx":66434
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBstart.frx":6B488
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBstart.frx":704DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBstart.frx":75530
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picButton 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   0
      Left            =   375
      Picture         =   "frmBstart.frx":7A584
      ScaleHeight     =   600
      ScaleWidth      =   2550
      TabIndex        =   1
      Top             =   3000
      Width           =   2550
   End
   Begin VB.Timer tmrHide 
      Interval        =   1
      Left            =   960
      Top             =   1200
   End
   Begin VB.PictureBox picDesktopCapture 
      AutoRedraw      =   -1  'True
      Height          =   495
      Left            =   3060
      ScaleHeight     =   435
      ScaleWidth      =   360
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   420
   End
End
Attribute VB_Name = "frmBstart"
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

Dim i As Integer
Dim Over(0 To 5) As Boolean
Dim OldIndex As Integer

Public Sub showme()
HideSubs
SetWindowPos Me.hWnd, -1, Me.ScaleLeft, Screen.Height / Screen.TwipsPerPixelY - Me.ScaleHeight - 30, Me.ScaleWidth, Me.ScaleHeight, SWP_NOREPOSITION
Me.Left = 0
BitBlt Me.hdc, picRightBorder.Left, picRightBorder.Top, picRightBorder.Width, picRightBorder.Height, picRightBorder.hdc, 0, 0, vbSrcCopy
BitBlt Me.hdc, picBottomBorder.Left, picBottomBorder.Top, picBottomBorder.Width, picBottomBorder.Height, picBottomBorder.hdc, 0, 0, vbSrcCopy

picDesktopCapture.Width = Me.ScaleWidth + 10
picDesktopCapture.Height = Me.ScaleHeight + 10
picDesktopCapture.Left = 0
picDesktopCapture.Top = 0

DeskHdc = GetDC(0)
ret = BitBlt(picDesktopCapture.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, DeskHdc, Me.Left / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY, vbSrcCopy)
ret = ReleaseDC(0&, DeskHdc)
Blend Me, picDesktopCapture, 20, 0, 0, Me.ScaleWidth, Me.ScaleHeight
For i = 0 To picButton.Count - 1
    AlphaBlending picButton(i).hdc, 0, 0, 170, 40, picDesktopCapture.hdc, picButton(i).Left, picButton(i).Top, 170, 40, 20
Next

Me.Show
Me.Refresh
End Sub


Private Sub cmdPrograms_Click()
M0.GetMenu ("C:\windows\start menu")
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Over(OldIndex) = True Then
    picButton(OldIndex).Picture = images.ListImages(OldIndex + 1).Picture
    AlphaBlending picButton(OldIndex).hdc, 0, 0, 170, 40, picDesktopCapture.hdc, picButton(OldIndex).Left, picButton(OldIndex).Top, 170, 40, 20
    Over(OldIndex) = False
End If
HideSubs
End Sub

Private Sub picButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    picButton(Index).Picture = imagesdown.ListImages(Index + 1).Picture
    AlphaBlending picButton(Index).hdc, 0, 0, 170, 40, picDesktopCapture.hdc, picButton(Index).Left, picButton(Index).Top, 170, 40, 20
    Select Case Index
        Case 0
                frmShutdownSubMenu.SetFocus
        Case 2
                frmHelpSubMenu.SetFocus
        Case 3
                frmSettingsSubMenu.SetFocus
        Case 4
                frmBAppsSubMenu.SetFocus
        End Select
End Sub

Private Sub picButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Over(Index) = False Then
    picButton(OldIndex).Picture = images.ListImages(OldIndex + 1).Picture
    AlphaBlending picButton(OldIndex).hdc, 0, 0, 170, 40, picDesktopCapture.hdc, picButton(OldIndex).Left, picButton(OldIndex).Top, 170, 40, 20
    Over(OldIndex) = False
    OldIndex = Index
    picButton(Index).Picture = imageshover.ListImages(Index + 1).Picture
    AlphaBlending picButton(Index).hdc, 0, 0, 170, 40, picDesktopCapture.hdc, picButton(Index).Left, picButton(Index).Top, 170, 40, 20
    Over(Index) = True
    Select Case Index
        Case 0
            HideSubs
            frmShutdownSubMenu.showme
            SubShown(0) = True
        Case 2
            HideSubs
            frmHelpSubMenu.showme
            SubShown(1) = True
        Case 3
            HideSubs
            frmSettingsSubMenu.showme
            SubShown(2) = True
        Case 4
            HideSubs
            frmBAppsSubMenu.showme
            SubShown(4) = True
        Case 5
            HideSubs
            Load M0
            M0.Top = Me.Top - M0.Height
            M0.Left = Me.Left + Me.Width - 200
            M0.GetMenu StartMenuPath
            SubShown(3) = True
        Case Else
            HideSubs
    End Select
    s_Playsound "hover"
End If
End Sub

Private Sub picButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
picButton(Index).Picture = imageshover.ListImages(Index + 1).Picture
AlphaBlending picButton(Index).hdc, 0, 0, 170, 40, picDesktopCapture.hdc, picButton(Index).Left, picButton(Index).Top, 170, 40, 20
Select Case Index
Case 1
    Load frmRun
    HideStartMenu
End Select
    s_Playsound "select"
End Sub

Private Sub tmrHide_Timer()
If TaskbarOpen Then
    a = GetForegroundWindow
    b = 0
    If SubShown(0) Then b = frmShutdownSubMenu.hWnd
    If SubShown(1) Then b = frmHelpSubMenu.hWnd
    If SubShown(2) Then b = frmSettingsSubMenu.hWnd
    If SubShown(3) Then b = M0.hWnd
    
    If a <> Me.hWnd And a <> b Then
        For i = 2 To Forms.Count - 1
            If Forms(i).hWnd = a Then Exit Sub
        Next
        HideStartMenu
    End If
End If
DoEvents
End Sub
