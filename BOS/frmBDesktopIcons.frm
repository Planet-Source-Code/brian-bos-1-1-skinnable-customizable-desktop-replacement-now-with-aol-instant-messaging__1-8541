VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBDesktopIcons 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2940
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmBDesktopIcons.frx":0000
   Moveable        =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   515
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   196
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDesktopCapture 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2475
      Left            =   -360
      ScaleHeight     =   165
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   0
      Top             =   5100
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picBG 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   10
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picTemp2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   9
      Top             =   540
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picBlack 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   3600
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   8
      Top             =   -1500
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picOver 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   3300
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   -1500
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.FileListBox File1 
      Height          =   4185
      Left            =   4920
      TabIndex        =   6
      Top             =   180
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.DirListBox Dir1 
      Height          =   4140
      Left            =   3900
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox lstDesktop 
      Height          =   4320
      IntegralHeight  =   0   'False
      Left            =   2520
      TabIndex        =   3
      Top             =   60
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer tmrCheckRefresh 
      Interval        =   500
      Left            =   3600
      Top             =   3600
   End
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.FileListBox fleDesktop 
      Height          =   4380
      Hidden          =   -1  'True
      Left            =   1620
      TabIndex        =   1
      Top             =   -60
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2280
      Top             =   6720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBDesktopIcons.frx":0152
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBDesktopIcons.frx":0D26
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBDesktopIcons.frx":18FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBDesktopIcons.frx":24CE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1080
      Top             =   7080
   End
   Begin VB.Shape shpFill 
      BorderColor     =   &H00404040&
      BorderStyle     =   3  'Dot
      Height          =   315
      Left            =   540
      Top             =   2580
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgMyComputer 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   0
      Picture         =   "frmBDesktopIcons.frx":31AE
      Top             =   960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblIcon 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "My Computer"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   480
      TabIndex        =   4
      Top             =   720
      Width           =   1035
   End
   Begin VB.Image imgIcon 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   0
      Left            =   750
      OLEDragMode     =   1  'Automatic
      Picture         =   "frmBDesktopIcons.frx":5950
      Top             =   180
      Width           =   480
   End
End
Attribute VB_Name = "frmBDesktopIcons"
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

Dim OldIndex As Integer
Dim Dragging As Boolean
Dim IconColor As ColorConstants
Dim IconsPerCol As Integer



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And OldIndex <> -1 Then
    If OldIndex > fleDesktop.ListCount Then
        a = Dir1.List(OldIndex - fleDesktop.ListCount - 1)
        MsgBox "Do you really want to delete the folder """ & a & """?", vbYesNo Or vbQuestion Or vbSystemModal Or vbMsgBoxSetForeground, "Delete Folder?"
    Else
        a = "C:\windows\desktop\" & fleDesktop.List(OldIndex - 1)
        b = MsgBox("Do you really want to delete the file """ & a & """?", vbYesNo Or vbQuestion Or vbSystemModal Or vbMsgBoxSetForeground, "Delete File?")
        If b = vbYes Then
            Kill a
            shpFill.Visible = False

        End If
    End If
End If
End Sub

Private Sub Form_Load()
Me.Height = Screen.Height - frmTaskbar.Height
OldIndex = -1
picDesktopCapture.Width = Me.ScaleWidth
picDesktopCapture.Height = Me.ScaleHeight


Me.Top = 0
Me.Left = 0


Me.BackColor = Val(GetSetting("BoS", "BDesktopIcons", "BgColor", "&H00000000&"))
Me.ForeColor = Val(GetSetting("BoS", "BDesktopIcons", "ShadowColor", "&H00000000&"))
IconColor = Val(GetSetting("BoS", "BDesktopIcons", "FgColor", "&H00FFFFFF&"))
lblIcon(0).ForeColor = IconColor

picBlack.BackColor = Me.BackColor


    DeskHdc = GetDC(0)
    BitBlt picDesktopCapture.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, DeskHdc, 0, 0, vbSrcCopy
    ReleaseDC 0, DeskHdc

If Translucent Then Blend Me, picDesktopCapture, 95, 0, 0, Me.ScaleWidth - 5, Me.ScaleHeight
frmLoading.ProgressBar1.Value = frmLoading.ProgressBar1.Value + 1

fleDesktop.path = DesktopPath
Dir1.path = DesktopPath
frmLoading.ProgressBar1.Value = frmLoading.ProgressBar1.Value + 1
IconsPerCol = Int(Me.ScaleHeight / 80)
frmLoading.ProgressBar1.Value = frmLoading.ProgressBar1.Value + 1


SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub



Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If OldIndex <> -1 And Index <> OldIndex Then DrawIcon (OldIndex): lblIcon(OldIndex).BackStyle = 0: shpFill.Visible = False: lblIcon(OldIndex).ForeColor = IconColor: OldIndex = -1
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim start As String, whereto As String
For i = 1 To Data.Files.Count
    start = Data.Files(i)
    If GetAttr(start) And vbDirectory Then
        If FileExists("C:\windows\desktop\" & ExtractFileName(start) & "\") Then
            MsgBox "The directory " & start & " could not be moved because a directory with the same name already exists.", vbInformation Or vbSystemModal Or vbMsgBoxSetForeground, "Copy Fail"
        Else
             MkDir "C:\windows\desktop\" & ExtractFileName(start)
             File1.path = AddASlash(start)
             For j = 0 To File1.ListCount - 1
                 FileCopy AddASlash(start) & File1.List(j), "C:\windows\desktop\" & AddASlash(ExtractFileName(start)) & File1.List(j)
            Next
        End If
    Else
        If FileExists("C:\windows\desktop\" & ExtractFileName(start)) Then
            If MsgBox("Overwrite the file ""C:\windows\desktop\" & ExtractFileName(start) & """?", vbYesNo Or vbMsgBoxSetForeground Or vbQuestion Or vbSystemModal, "Overwrite file?") = vbNo Then GoTo no
            Kill "C:\windows\desktop\" & ExtractFileName(start)
        End If
            FileCopy start, "C:\windows\desktop\" & ExtractFileName(start)
no:
    End If
Next

UpdateDesktop
End Sub

Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    If Data.GetFormat(vbCFFiles) And Dragging = False Then
        Effect = vbDropEffectCopy
    Else
        Effect = vbDropEffectNone
    End If
End Sub




Private Sub imgIcon_DblClick(Index As Integer)
If Index = 0 Then
    If FileExists("C:\My Computer\") Then
        ShellExecute Me.hWnd, "open", "C:\my computer\", "", "", 1
    Else
        MsgBox "To add My Computer functionality to BoS, make a folder called ""My Computer"" in your C drive and add shortcuts to your drives.", vbInformation, "My Computer"
    End If
Else
    ShellExecute Me.hWnd, "open", "C:\windows\desktop\" & lblIcon(Index).Caption, "", "", 1
End If
End Sub

Private Sub imgIcon_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index <> OldIndex Then
    picTemp2.Cls
    PicTemp.Cls
    If Translucent Then
    BitBlt picTemp2.hdc, 0, 0, 32, 32, picDesktopCapture.hdc, imgIcon(Index).Left, imgIcon(Index).Top, vbSrcCopy
    Blend picTemp2, picBG, 95, 0, 0, 32, 32
    PicTemp.Picture = imgIcon(Index).Picture
    Blend PicTemp, picTemp2, 95, 0, 0, 32, 32
    
    Else
        PicTemp.Picture = imgIcon(Index).Picture
        Blend PicTemp, picOver, 95, 0, 0, 32, 32
    End If
    imgIcon(Index).Picture = PicTemp.Image
    lblIcon(Index).BackStyle = 1
    shpFill.Move lblIcon(Index).Left, lblIcon(Index).Top, lblIcon(Index).Width, lblIcon(Index).Height
    shpFill.Visible = True
    lblIcon(Index).ForeColor = vbHighlightText
End If
If OldIndex > -1 And Index <> OldIndex Then DrawIcon (OldIndex): lblIcon(OldIndex).BackStyle = 0: lblIcon(OldIndex).ForeColor = IconColor
OldIndex = Index
End Sub

Private Sub imgIcon_OLECompleteDrag(Index As Integer, Effect As Long)
imgIcon_MouseDown Index, 0, 0, 1, 1
Dragging = False
End Sub

Private Sub imgIcon_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "I'm sorry Dave but I can't let you do that until I reach version 1.1", vbInformation, "Dave"
End Sub

Private Sub imgIcon_OLEDragOver(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    If Data.GetFormat(vbCFFiles) And Dragging = False Then
            Effect = vbDropEffectCopy
    Else
        Effect = vbDropEffectNone
    End If
    If State = vbEnter Then lblIcon_MouseDown Index, 1, 1, 1, 1
    If State = vbLeave Then Form_MouseDown 1, 1, 1, 1
End Sub

Private Sub imgIcon_OLEStartDrag(Index As Integer, Data As DataObject, AllowedEffects As Long)
Data.SetData , vbCFFiles
Data.Files.Add "C:\windows\desktop\" & lblIcon(Index).Caption
Dragging = True
End Sub

Private Sub lblIcon_DblClick(Index As Integer)
If Index = 0 Then
    If FileExists("C:\My Computer\") Then
        ShellExecute Me.hWnd, "open", "C:\my computer\", "", "", 1
    Else
        MsgBox "To add My Computer functionality to BoS, make a folder called ""My Computer"" in your C drive and add shortcuts to your drives.", vbInformation, "My Computer"
    End If
Else
    ShellExecute Me.hWnd, "open", "C:\windows\desktop\" & lblIcon(Index).Caption, "", "", 1
End If
End Sub

Private Sub lblIcon_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
imgIcon_MouseDown Index, 0, 0, 1, 1
End Sub


Private Sub Timer1_Timer()
     
     DoEvents
    Dim a As Long
Dim b As RECT

a = GetForegroundWindow
GetWindowRect a, b
If a <> frmBstart.hWnd And a <> frmHelpSubMenu.hWnd And a <> frmTaskbar.hWnd And a <> Me.hWnd And a <> frmPSCTicker.hWnd And a <> frmSettingsSubMenu.hWnd And a <> frmBAppsSubMenu.hWnd And a <> M0.hWnd And a <> frmShutdownSubMenu.hWnd And ((b.Right - b.Left) < Screen.Width / Screen.TwipsPerPixelX Or IsZoomed(a)) Then
    If b.Right - b.Left > Screen.Width / Screen.TwipsPerPixelX Then
        SetWindowPos a, 0, frmBDesktopIcons.ScaleWidth, 0, Screen.Width / Screen.TwipsPerPixelX - frmBDesktopIcons.ScaleWidth, Screen.Height / Screen.TwipsPerPixelY - 30, SWP_NOACTIVATE
    Else
        If b.Left < frmBDesktopIcons.ScaleWidth Then
            SetWindowPos a, 0, frmBDesktopIcons.ScaleWidth, b.Top, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE
        End If
        If b.Bottom > Screen.Height / Screen.TwipsPerPixelY - 30 Then
            SetWindowPos a, 0, b.Left, Screen.Height / Screen.TwipsPerPixelY - 30 - b.Bottom + b.Top, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE
        End If
    End If
End If

End Sub

Private Sub tmrCheckRefresh_Timer()
fleDesktop.Refresh
Dir1.Refresh
For i = 0 To fleDesktop.ListCount + Dir1.ListCount - 1
    If i > fleDesktop.ListCount - 1 Then
        a = Dir1.List(i - fleDesktop.ListCount)
    Else
        a = fleDesktop.List(i)
    End If
    
    If lstDesktop.List(i) <> a Then
        UpdateDesktop
        Exit For
    End If
Next
lstDesktop.Clear
For i = 0 To fleDesktop.ListCount - 1
    lstDesktop.AddItem fleDesktop.List(i)
Next
For i = 0 To Dir1.ListCount - 1
    lstDesktop.AddItem Dir1.List(i)
Next
End Sub

Sub UpdateDesktop()
Dim i As Integer
Dim a As String
col = 0
For i = imgIcon.Count To fleDesktop.ListCount + Dir1.ListCount Step -1
    imgIcon(i - 1).Picture = LoadPicture()
    lblIcon(i - 1).BackStyle = 0
    lblIcon(i - 1).Caption = ""
Next

For i = 1 To fleDesktop.ListCount + Dir1.ListCount
    If i >= imgIcon.Count Then
        If i = IconsPerCol Then col = col + 1
        Load imgIcon(i)
        Load lblIcon(i)
        imgIcon(i).Top = (10 + 80 * (i - (IconsPerCol * col)))
        lblIcon(i).Top = (42 + 80 * (i - (IconsPerCol * col)))
        lblIcon(i).Left = col * 140 + 32
        imgIcon(i).Left = col * 140 + 50
        imgIcon(i).Visible = True
        lblIcon(i).Visible = True
    End If
    
    If i > fleDesktop.ListCount Then
        a = Dir1.List(i - fleDesktop.ListCount - 1)
        a = ExtractFileName(a)
    Else
        a = fleDesktop.List(i - 1)
    End If
    
    DrawIcon i
    
    If Right(a, 4) = ".lnk" Then
        lblIcon(i).Caption = Left(a, Len(a) - 4)
    Else
        lblIcon(i).Caption = a
    End If
Next
nw = (col + 1) * 140 * Screen.TwipsPerPixelX
ow = Me.Width
If Me.Width <> nw Then
    Me.Width = nw
    Me.Cls
    picDesktopCapture.Width = Me.ScaleWidth
    If nw > ow Then
        DeskHdc = GetDC(0)
        BitBlt picDesktopCapture.hdc, ow / Screen.TwipsPerPixelX, 0, Me.ScaleWidth - (ow / Screen.TwipsPerPixelX), Me.ScaleHeight, DeskHdc, ow / Screen.TwipsPerPixelX, 0, vbSrcCopy
        ReleaseDC 0, DeskHdc
        picDesktopCapture.Refresh
    End If
    If Translucent Then
        Blend Me, picDesktopCapture, TransLevel, 0, 0, Me.ScaleWidth - 5, Me.ScaleHeight
    End If
    For i = 1 To 5
        Me.Line (Me.ScaleWidth - 5 + i, 0)-(Me.ScaleWidth - 5 + i, Me.ScaleHeight)
    Next
    For i = 0 To 5
        AlphaBlending Me.hdc, Me.ScaleWidth - 5, 5 + i, 5, 1, picDesktopCapture.hdc, Me.ScaleWidth - 5, 5 + i, 5, 1, (5 - i) * 50
    Next
    BitBlt Me.hdc, Me.ScaleWidth - 5, 0, 5, 5, picDesktopCapture.hdc, Me.ScaleWidth - 5, 0, vbSrcCopy
    For i = 1 To 5
        AlphaBlending Me.hdc, Me.ScaleWidth - 5 + i, 5, 1, Me.ScaleHeight - 5, picDesktopCapture.hdc, Me.ScaleWidth - 5 + i, 5, 1, Me.ScaleHeight - 5, 70 * i - 50
    Next
    For i = 0 To 4
        frmTaskbar.Line (0, i)-(Me.ScaleWidth, i)
    Next
    For i = 0 To 4
        AlphaBlending frmTaskbar.hdc, i + 5, 0, 1, 5, frmBDesktopIcons.hdc, i + 5, frmBDesktopIcons.ScaleHeight - 5, 1, 5, (5 - i) * 50
    Next
    For i = 0 To 4
        AlphaBlending frmTaskbar.hdc, 5, i, frmBDesktopIcons.ScaleWidth - 5, 1, frmBDesktopIcons.hdc, 5, frmBDesktopIcons.ScaleHeight - 5 + i, frmBDesktopIcons.ScaleWidth - 5, 1, (5 - i) * 50
    Next
    BitBlt frmTaskbar.hdc, 0, 0, 5, 5, Me.hdc, 0, Me.ScaleHeight - 5, vbSrcCopy
    If nw < ow Then
            For i = 0 To 4
                frmTaskbar.Line (nw - 5, i)-(nw, i)
            Next
            For i = 0 To 4
                AlphaBlending Me.hdc, frmBDesktopIcons.ScaleWidth - 5, i, (ow - nw) / Screen.TwipsPerPixelX + 5, 1, picDesktopCapture.hdc, frmBDesktopIcons.ScaleWidth, i, Me.ScaleWidth - frmBDesktopIcons.ScaleWidth, 1, (5 - i) * 50
            Next
    End If
End If
End Sub

Sub DrawIcon(OldIndex As Integer)
If OldIndex = 0 Then
    imgIcon(i).Picture = imgMyComputer.Picture
    Exit Sub
End If

If OldIndex > fleDesktop.ListCount Then
    a = Dir1.List(OldIndex - fleDesktop.ListCount - 1)
    a = ExtractFileName(a)
Else
    a = fleDesktop.List(OldIndex - 1)
End If

hImgLarge& = SHGetFileInfo("C:\windows\desktop\" & a, 0&, shinfo, Len(shinfo), _
BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
PicTemp.Picture = LoadPicture
If Translucent Then
    BitBlt PicTemp.hdc, 0, 0, 32, 32, picDesktopCapture.hdc, 50, 10 + 80 * OldIndex, vbSrcCopy
    AlphaBlending PicTemp.hdc, 0, 0, 32, 32, picBlack.hdc, 0, 0, 32, 32, 160
Else
    PicTemp.BackColor = Me.BackColor
End If
ImageList_Draw hImgLarge&, shinfo.iIcon, PicTemp.hdc, 0, 0, ILD_TRANSPARENT
imgIcon(OldIndex).Picture = PicTemp.Image
    
End Sub



