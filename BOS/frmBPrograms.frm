VERSION 5.00
Begin VB.Form M0 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFDCB7&
   BorderStyle     =   0  'None
   ClientHeight    =   6210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3780
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   414
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   252
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picWhiteArrow 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF8B0F&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1020
      Picture         =   "frmBPrograms.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   4440
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   3540
      TabIndex        =   1
      Top             =   4980
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Timer tmrHold 
      Left            =   1320
      Top             =   2280
   End
   Begin VB.PictureBox picBlackArrow 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFDCB7&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   420
      Picture         =   "frmBPrograms.frx":058A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   4440
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFDCB7&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   15
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   3
      Top             =   15
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFDCB7&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   15
      ScaleHeight     =   300
      ScaleWidth      =   4035
      TabIndex        =   2
      Top             =   0
      Width           =   4035
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   2280
      TabIndex        =   0
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "M0"
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
Dim Over() As Boolean
Public MenuIndex As Integer
Dim MaxLen As Integer
Dim curIndex As Integer
Dim mWidth As Integer



Private Sub Form_Load()
MenuIndex = -1
End Sub



Private Sub picItem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    If Index > Dir1.ListCount - 1 Then
        ShellExecute Me.hWnd, "open", Dir1.path & "\" & Right(picItem(Index).Tag, Len(picItem(Index).Tag) - 8), "", "", 1
        M0.HideMe
        HideStartMenu
        s_Playsound "select"
    End If
ElseIf Button = vbRightButton Then
    
End If
End Sub

Private Sub picItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Dir1.ListCount + File1.ListCount = 0 Then Exit Sub
If Over(Index) = False Then
    picItem(Index).BackColor = &HFF8B0F
    picItem(Index).Cls
    picItem(Index).Line (1, 1)-(picItem(Index).ScaleWidth, 1), vbWhite
    picItem(Index).Line (1, picItem(Index).ScaleHeight - 20)-(picItem(Index).ScaleWidth, picItem(Index).ScaleHeight - 20), &H804000
    picItem(Index).CurrentX = 0
    picItem(Index).CurrentY = ((20 * Screen.TwipsPerPixelY) - picItem(0).TextHeight(picItem(Index).Tag)) / 2
    picItem(Index).Print picItem(Index).Tag
    If Index < Dir1.ListCount Then BitBlt picItem(Index).hdc, Me.ScaleWidth - 24, 1, 16, 16, picWhiteArrow.hdc, 0, 0, vbSrcCopy
    If Index > Dir1.ListCount - 1 Then
        PicTemp.BackColor = &HFF8B0F
        DrawIcon Dir1.path & "\" & File1.List(Index - Dir1.ListCount), Index, False
        PicTemp.Line (0, 0)-(22, 0), vbWhite
        PicTemp.Line (0, 19)-(22, 19), &H804000
        BitBlt picItem(Index).hdc, 0, 0, 21, 20, PicTemp.hdc, 0, 0, vbSrcCopy
    Else
        PicTemp.BackColor = &HFF8B0F
        DrawIcon Dir1.List(Index), Index, False
        PicTemp.Line (0, 0)-(22, 0), vbWhite
        PicTemp.Line (0, 19)-(22, 19), &H804000
        BitBlt picItem(Index).hdc, 0, 0, 21, 20, PicTemp.hdc, 0, 0, vbSrcCopy
    End If
    Over(Index) = True
    If Index <> OldIndex Then
        picItem(OldIndex).BackColor = &HFFDCB7
        picItem(OldIndex).Cls
        picItem(OldIndex).CurrentY = ((20 * Screen.TwipsPerPixelY) - picItem(0).TextHeight(picItem(OldIndex).Tag)) / 2
        picItem(OldIndex).Print picItem(OldIndex).Tag
        PicTemp.Cls
            PicTemp.BackColor = &HFFDCB7
        If OldIndex < Dir1.ListCount Then BitBlt picItem(OldIndex).hdc, Me.ScaleWidth - 24, 1, 18, 18, picBlackArrow.hdc, 0, 0, vbSrcCopy
        Over(OldIndex) = False
        If OldIndex > Dir1.ListCount - 1 Then
            DrawIcon Dir1.path & "\" & File1.List(OldIndex - Dir1.ListCount), OldIndex
        Else
            DrawIcon Dir1.List(OldIndex), OldIndex
        End If
        tmrHold.Interval = 0
    End If
    
    If Index < Dir1.ListCount Then
        curIndex = Index
        tmrHold.Interval = 300
    End If
    If Index <> OldIndex Then
        If MenuIndex <> -1 Then
            Forms(MenuIndex).HideMe
            MenuIndex = -1
        End If
        OldIndex = Index
    End If
    If Index > Dir1.ListCount - 1 Then
        s_Playsound "hover"
    End If
End If
End Sub

Public Sub GetMenu(path As String)
DoEvents
Dir1.path = path
File1.path = path
If File1.ListCount + Dir1.ListCount = 0 Then
    picItem(0).CurrentY = ((picItem(0).Height * Screen.TwipsPerPixelY) - picItem(0).TextHeight("|")) / 2
    picItem(0).Print "[ Empty ]"
    MaxLen = picItem(0).TextWidth("[ Empty ]")
Else
    If Dir1.ListCount > 0 Then
        For i = 1 To Dir1.ListCount + File1.ListCount - 1
            Load picItem(i)
            picItem(i).Visible = True
            picItem(i).Top = 20 * i + 1
        Next
        For i = 0 To Dir1.ListCount - 1
            DrawIcon Dir1.List(i), i
            picItem(i).CurrentY = ((20 * Screen.TwipsPerPixelY) - picItem(0).TextHeight("        " & Left(Dir1.List(i), Len(Dir1.List(i)) - 4))) / 2
            picItem(i).Print "        " & ExtractFileName(Dir1.List(i))
            picItem(i).Tag = "        " & ExtractFileName(Dir1.List(i))
            
            If picItem(0).TextWidth(picItem(i).Tag) > MaxLen Then MaxLen = picItem(0).TextWidth(picItem(i).Tag)
        Next
        For i = 0 To File1.ListCount - 1
            DrawIcon Dir1.path & "\" & File1.List(i), i + Dir1.ListCount
            picItem(i + Dir1.ListCount).CurrentY = ((20 * Screen.TwipsPerPixelY) - picItem(0).TextHeight("        " & Left(File1.List(i), Len(File1.List(i)) - 4))) / 2
            picItem(i + Dir1.ListCount).Print "        " & Left(File1.List(i), Len(File1.List(i)) - 4)
            picItem(i + Dir1.ListCount).Tag = "        " & Left(File1.List(i), Len(File1.List(i)) - 4)
            If picItem(0).TextWidth(picItem(i + Dir1.ListCount).Tag) > MaxLen Then MaxLen = picItem(0).TextWidth(picItem(i + Dir1.ListCount).Tag)
        Next
    Else
        For i = 1 To File1.ListCount - 1
            Load picItem(i)
            picItem(i).Visible = True
            picItem(i).Top = 20 * i + 1
        Next
        For i = 0 To File1.ListCount - 1
            DrawIcon Dir1.path & "\" & File1.List(i), i
            picItem(0).CurrentY = ((20 * Screen.TwipsPerPixelY) - picItem(0).TextHeight("        " & Left(File1.List(i), Len(File1.List(i)) - 4))) / 2
            picItem(i).Print "        " & Left(File1.List(i), Len(File1.List(i)) - 4)
            picItem(i).Tag = "        " & Left(File1.List(i), Len(File1.List(i)) - 4)
            If picItem(0).TextWidth(picItem(i).Tag) > MaxLen Then MaxLen = picItem(0).TextWidth(picItem(i).Tag)
        Next
    End If
    ReDim Over(Dir1.ListCount + File1.ListCount - 1)
End If
mWidth = MaxLen + 500
MHeight = (picItem.Count * 20 + 2) * Screen.TwipsPerPixelY
Me.Width = 0
Me.Height = 0
SetWindowPos Me.hWnd, -1, Me.Left / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY, Me.ScaleWidth, Me.ScaleHeight + 10, SWP_NOREPOSITION
mleft = Me.Left
If Dir1.path = StartMenuPath Then
    Me.Top = frmBstart.Top - MHeight + (frmBstart.picButton(5).Height * Screen.TwipsPerPixelY)
    mtop = Me.Top
    foldup = True
End If

For i = 0 To picItem.Count - 1
    picItem(i).Width = (mWidth / Screen.TwipsPerPixelX) - 2
Next

Me.Show
Me.Refresh
For i = 0 To Dir1.ListCount - 1
    BitBlt picItem(i).hdc, (mWidth / Screen.TwipsPerPixelX) - 24, 1, 16, 16, picBlackArrow.hdc, 0, 0, vbSrcCopy
Next
If foldup = False Then
    For i = 1 To mWidth Step (mWidth / 30)
        Me.Width = i
        Me.Height = i * (MHeight / mWidth)
        Me.Cls
        Me.Refresh
    Next
End If
Me.Width = mWidth
Me.Height = MHeight
DoEvents
Me.Line (0, 0)-(0, Me.ScaleHeight), vbWhite
Me.Line (0, 0)-(Me.ScaleWidth, 0), vbWhite
Me.Line (0, Me.ScaleHeight - 1)-(Me.ScaleWidth, Me.ScaleHeight - 1), &H804000
Me.Line (Me.ScaleWidth - 1, Me.ScaleHeight - 1)-(Me.ScaleWidth - 1, -1), &H804000

End Sub

Public Sub HideMe()
    If MenuIndex > Forms.Count Then Exit Sub
    If MenuIndex > -1 Then Forms(MenuIndex).HideMe
    Unload Me
End Sub

Sub DrawIcon(path, Index, Optional blt = True)
    hImgLarge& = SHGetFileInfo(path, 0&, shinfo, Len(shinfo), _
    BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
    PicTemp.Cls
    If blt Then
        ImageList_Draw hImgLarge&, shinfo.iIcon, PicTemp.hdc, 2, 2, ILD_TRANSPARENT
        BitBlt picItem(Index).hdc, 0, 0, 20, 20, PicTemp.hdc, 0, 0, vbSrcCopy
    Else
        ImageList_Draw hImgLarge&, shinfo.iIcon, PicTemp.hdc, 2, 2, ILD_TRANSPARENT
    End If
End Sub


Private Sub tmrHold_Timer()
If tmrHold.Interval = 0 Then Exit Sub
If curIndex > Dir1.ListCount - 1 Then
    tmrHold.Interval = 0
    Exit Sub
End If
Dim f As New M0
f.Top = Me.Top + picItem(curIndex).Top * Screen.TwipsPerPixelX
f.Left = Me.Left + Me.Width - 50
f.GetMenu Dir1.path & "\" & Right(picItem(curIndex).Tag, Len(picItem(curIndex).Tag) - 8)
MenuIndex = Forms.Count - 1
s_Playsound "open"
tmrHold.Interval = 0
End Sub

