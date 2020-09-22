VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStartup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BToday"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10605
   Icon            =   "frmStartup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   10605
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   7095
      Left            =   180
      ScaleHeight     =   7095
      ScaleWidth      =   10275
      TabIndex        =   2
      Top             =   600
      Width           =   10275
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Loading..."
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
         Left            =   2340
         TabIndex        =   4
         Top             =   3600
         Width           =   4995
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start BoS"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   0
      Top             =   60
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3480
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartup.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartup.frx":089E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartup.frx":0CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartup.frx":1156
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartup.frx":16F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartup.frx":1FCE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   7035
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   10155
      ExtentX         =   17912
      ExtentY         =   12409
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   7575
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   13361
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "TV Listings"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Weather"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "News"
            ImageVarType    =   2
            ImageIndex      =   3
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Sports"
            ImageVarType    =   2
            ImageIndex      =   4
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "PSC Newest Code"
            ImageVarType    =   2
            ImageIndex      =   5
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmStartup"
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

Dim a As String, Zip As String, lc As String, TVC As String

Private Sub Command1_Click()
frmTaskbar.Show
frmBDesktopIcons.Show
Unload Me
End Sub

Private Sub Form_Load()
Zip = GetSetting("BoS", "BToday", "Zipcode", "None")
TVC = GetSetting("BoS", "BToday", "TVC", "")
If Zip = "None" Then
    a = InputBox("Please enter your zipcode:", "Zipcode")
    Zip = a
    SaveSetting "BoS", "BToday", "Zipcode", Zip
End If
TabStrip1_MouseUp 1, 1, 1, 1
End Sub

Private Sub TabStrip1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Select Case TabStrip1.SelectedItem.Index
        Case 1
            If TVC = "" Then
                lc = "http://tv.yahoo.com/yahoo/register/cableproviders.dpg?zipcode=" & Zip
            Else
                lc = TVC
            End If
        Case 2
            lc = "http://search.weather.yahoo.com/weather/query.cgi?q=" & Zip
        Case 3
            lc = "http://dailynews.yahoo.com/"
        Case 4
            lc = "http://sports.yahoo.com/"
        Case 5
            lc = "http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?grpCategories=-1&optSort=DateDescending&txtMaxNumberOfEntriesPerPage=10&blnNewestCode=TRUE&blnResetAllVariables=TRUE&lngWId=1"
        Case 6
            lc = "http://thunder.prohosting.com/~pikared/home.shtml"
    End Select
    On Error Resume Next
    WebBrowser1.Stop
    WebBrowser1.Navigate2 lc
    Picture1.Visible = True
ElseIf Button = vbRightButton Then

End If
End Sub

Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
If Left(WebBrowser1.LocationURL, 52) = "http://tv.yahoo.com/yahoo/listings/tv1.dpg?chanArea=" Then
    TVC = WebBrowser1.LocationURL
    SaveSetting "BoS", "BToday", "TVC", TVC
End If
Picture1.Visible = False
End Sub
