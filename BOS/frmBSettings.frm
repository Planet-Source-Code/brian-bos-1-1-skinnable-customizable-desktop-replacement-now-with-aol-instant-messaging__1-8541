VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bos Settings"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   Icon            =   "frmBSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   5280
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1020
      Top             =   6900
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBSettings.frx":27A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBSettings.frx":4F56
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBSettings.frx":770A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBSettings.frx":9EBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBSettings.frx":A312
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBSettings.frx":B596
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBSettings.frx":B93E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBSettings.frx":E0F2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.DirListBox dirSkins 
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   -480
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3000
      Width           =   1455
   End
   Begin VB.PictureBox picSettings 
      BorderStyle     =   0  'None
      Height          =   2055
      Index           =   0
      Left            =   180
      ScaleHeight     =   2055
      ScaleWidth      =   4935
      TabIndex        =   3
      Top             =   720
      Width           =   4935
      Begin VB.CheckBox chkShowDesktopIcons 
         Caption         =   "Show Desktop Icons"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   600
         TabIndex        =   6
         Top             =   1380
         Value           =   1  'Checked
         Width           =   3855
      End
      Begin VB.CheckBox chkShowWebAddress 
         Caption         =   "Show web address box on taskbar"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   600
         TabIndex        =   5
         Top             =   120
         Width           =   3855
      End
      Begin VB.CheckBox chkShowPSCButton 
         Caption         =   "Show Planet Source Code Button"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   600
         TabIndex        =   4
         Top             =   780
         Value           =   1  'Checked
         Width           =   3855
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   6
         Left            =   0
         Picture         =   "frmBSettings.frx":E546
         Stretch         =   -1  'True
         Top             =   1260
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   0
         Picture         =   "frmBSettings.frx":10CE8
         Top             =   0
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   0
         Picture         =   "frmBSettings.frx":1348A
         Stretch         =   -1  'True
         Top             =   660
         Width           =   480
      End
   End
   Begin VB.PictureBox picSettings 
      BorderStyle     =   0  'None
      Height          =   1995
      Index           =   3
      Left            =   240
      ScaleHeight     =   1995
      ScaleWidth      =   4695
      TabIndex        =   13
      Top             =   780
      Width           =   4695
      Begin VB.CommandButton cmdConfigureStartupPrograms 
         Caption         =   "Configure Startup Programs..."
         Height          =   375
         Left            =   2160
         TabIndex        =   23
         Top             =   1380
         Width           =   2475
      End
      Begin VB.TextBox txtDunConnection 
         Height          =   315
         Left            =   2580
         TabIndex        =   16
         Top             =   240
         Width           =   1875
      End
      Begin VB.OptionButton optTime 
         Caption         =   "12 Hour"
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
         Index           =   0
         Left            =   1800
         TabIndex        =   15
         Top             =   900
         Width           =   915
      End
      Begin VB.OptionButton optTime 
         Caption         =   "24 Hour"
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
         Index           =   1
         Left            =   2760
         TabIndex        =   14
         Top             =   900
         Width           =   1155
      End
      Begin VB.Label Label7 
         Caption         =   "Startup Programs:"
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
         Left            =   600
         TabIndex        =   24
         Top             =   1440
         Width           =   2355
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   7
         Left            =   0
         Picture         =   "frmBSettings.frx":13D54
         Top             =   1260
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   0
         Picture         =   "frmBSettings.frx":14196
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Name of Dial Up Networking Connection:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   18
         Top             =   180
         Width           =   1995
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   4
         Left            =   0
         Picture         =   "frmBSettings.frx":16938
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label3 
         Caption         =   "Clock Format:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   17
         Top             =   900
         Width           =   1035
      End
   End
   Begin VB.PictureBox picSettings 
      BorderStyle     =   0  'None
      Height          =   1995
      Index           =   1
      Left            =   240
      ScaleHeight     =   1995
      ScaleWidth      =   4635
      TabIndex        =   8
      Top             =   720
      Width           =   4635
      Begin VB.ComboBox cmbSkin 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmBSettings.frx":16D7A
         Left            =   600
         List            =   "frmBSettings.frx":16D81
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   540
         Width           =   3795
      End
      Begin VB.CommandButton cmdInstallSkin 
         Caption         =   "Install a skin from the BoS internet skin archive"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   60
         Picture         =   "frmBSettings.frx":16D93
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1020
         Width           =   4515
      End
      Begin VB.Label Label5 
         Caption         =   "Choose your skin:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   660
         TabIndex        =   12
         Top             =   120
         Width           =   3555
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   5
         Left            =   0
         Picture         =   "frmBSettings.frx":16EDD
         Top             =   0
         Width           =   480
      End
      Begin VB.Label Label4 
         Caption         =   "Skin:"
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
         Left            =   60
         TabIndex        =   11
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.PictureBox picSettings 
      BorderStyle     =   0  'None
      Height          =   2055
      Index           =   4
      Left            =   180
      ScaleHeight     =   2055
      ScaleWidth      =   4995
      TabIndex        =   19
      Top             =   720
      Width           =   4995
      Begin VB.CheckBox chkEnableBToday 
         Caption         =   "Enable BToday"
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
         Left            =   720
         TabIndex        =   22
         Top             =   840
         Width           =   2835
      End
      Begin VB.TextBox txtZipcode 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1380
         TabIndex        =   20
         Text            =   "12345"
         Top             =   180
         Width           =   975
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   60
         Picture         =   "frmBSettings.frx":1967F
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label6 
         Caption         =   "Zipcode:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   660
         TabIndex        =   21
         Top             =   240
         Width           =   1155
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   60
         Picture         =   "frmBSettings.frx":19AC1
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.PictureBox picSettings 
      BorderStyle     =   0  'None
      Height          =   1995
      Index           =   5
      Left            =   240
      ScaleHeight     =   1995
      ScaleWidth      =   4815
      TabIndex        =   25
      Top             =   780
      Width           =   4815
      Begin VB.CheckBox chkAutoLogin 
         Caption         =   "Automaticlly sign on when BoS starts"
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
         Left            =   60
         TabIndex        =   34
         Top             =   840
         Width           =   4515
      End
      Begin VB.CheckBox chkImSounds 
         Caption         =   "Play IM Sounds"
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
         Left            =   60
         TabIndex        =   33
         Top             =   540
         Value           =   1  'Checked
         Width           =   4515
      End
      Begin VB.Frame Frame3 
         Caption         =   "Security Options"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   0
         TabIndex        =   27
         Top             =   1140
         Width           =   4755
         Begin VB.OptionButton optControl 
            Caption         =   "Block all users"
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
            Index           =   3
            Left            =   3240
            TabIndex        =   32
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optControl 
            Caption         =   "Allow buddies only"
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
            Index           =   1
            Left            =   1500
            TabIndex        =   31
            Top             =   240
            Width           =   1935
         End
         Begin VB.OptionButton optControl 
            Caption         =   "Allow all users"
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
            Index           =   0
            Left            =   60
            TabIndex        =   30
            Top             =   240
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.OptionButton optControl 
            Caption         =   "Block listed users"
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
            Index           =   4
            Left            =   2100
            TabIndex        =   29
            Top             =   480
            Width           =   1935
         End
         Begin VB.OptionButton optControl 
            Caption         =   "Allow listed users only"
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
            Index           =   2
            Left            =   60
            TabIndex        =   28
            Top             =   480
            Width           =   2235
         End
      End
      Begin VB.Label Label8 
         Caption         =   "Configure the BoS Instant Messanger."
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   660
         TabIndex        =   26
         Top             =   120
         Width           =   4035
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   9
         Left            =   60
         Picture         =   "frmBSettings.frx":19F03
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.PictureBox picSettings 
      BorderStyle     =   0  'None
      Height          =   2055
      Index           =   2
      Left            =   180
      ScaleHeight     =   137
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   329
      TabIndex        =   35
      Top             =   780
      Width           =   4935
      Begin VB.CheckBox DisableTrans 
         Caption         =   "Disable Translucency"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   3195
      End
      Begin VB.PictureBox picBIcon 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   4920
         Picture         =   "frmBSettings.frx":1B175
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   38
         Top             =   2040
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.PictureBox picTransTest 
         AutoRedraw      =   -1  'True
         Height          =   540
         Left            =   4140
         Picture         =   "frmBSettings.frx":1C1B7
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   37
         Top             =   420
         Width           =   540
      End
      Begin MSComctlLib.Slider sliTranslucency 
         Height          =   375
         Left            =   240
         TabIndex        =   36
         Top             =   1380
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   661
         _Version        =   393216
         LargeChange     =   32
         SmallChange     =   8
         Max             =   255
         TickFrequency   =   32
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000014&
         X1              =   250
         X2              =   250
         Y1              =   12
         Y2              =   124
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000010&
         X1              =   249
         X2              =   249
         Y1              =   12
         Y2              =   124
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   4
         X2              =   248
         Y1              =   49
         Y2              =   49
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   4
         X2              =   248
         Y1              =   48
         Y2              =   48
      End
      Begin VB.Label Label10 
         Caption         =   "Opaque"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2820
         TabIndex        =   41
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label Label9 
         Caption         =   "Transpearent"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         TabIndex        =   40
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Preview:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4020
         TabIndex        =   39
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Translucency Level:"
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
         Left            =   0
         TabIndex        =   43
         Top             =   840
         Width           =   3015
      End
   End
   Begin MSComctlLib.TabStrip tsSettings 
      Height          =   2835
      Left            =   60
      TabIndex        =   7
      Top             =   60
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   5001
      MultiRow        =   -1  'True
      TabFixedWidth   =   2999
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Visible Items"
            Object.ToolTipText     =   "Choose what items to show or hide"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Skin "
            Object.ToolTipText     =   "Choose a skin to customize the appearence of BoS"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Translucency"
            Object.ToolTipText     =   "Configure translucency settings for the BTaskbar and BDesktopIcons"
            ImageVarType    =   2
            ImageIndex      =   6
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Behavior"
            Object.ToolTipText     =   "Customize options like startup programs, clock format and more"
            ImageVarType    =   2
            ImageIndex      =   3
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "BToday"
            Object.ToolTipText     =   "Enable or disable BToday and change settings for it"
            ImageVarType    =   2
            ImageIndex      =   4
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "BIM"
            Object.ToolTipText     =   "Configure BIM, BoS's integrated AOL Instant Message compatible Instant Messanger"
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
Attribute VB_Name = "frmBSettings"
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

Dim TaskbarRefresh As Boolean, IconRefresh As Boolean, ShowAddress As Boolean, ShowPSC As Boolean

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdConfigureStartupPrograms_Click()
    frmBStartupItems.Show
End Sub

Private Sub cmdInstallSkin_Click()
Dim tFilePath As String
Dim NewFileName As String
Dim NewFolder As String
Dim ExecString As String
frmGetSkin.Show
'If ExtractPath(cdBrowse.FileName) <> App.path Then
'    FileCopy cdBrowse.FileName, App.path & "\" & ExtractFileName(cdBrowse.FileName)
'    tFilePath = App.path & "\" & ExtractFileName(cdBrowse.FileName)
'Else
'    tFilePath = cdBrowse.FileName
'End If
'NewFileName = Left(tFilePath, Len(tFilePath) - 4) & ".zip"
'If FileExists(NewFileName) Then Exit Sub
'Name tFilePath As NewFileName
'NewFolder = App.path & "\skins\" & Left(ExtractFileName(NewFileName), Len(ExtractFileName(NewFileName)) - 4)
'If FileExists(NewFolder) Then Exit Sub
'MkDir NewFolder
'FileCopy NewFileName, NewFolder & "\" & ExtractFileName(NewFileName)
'ExecString = App.path & "\" & "unzip.exe " & NewFolder & "\" & ExtractFileName(NewFileName) & " -d " & NewFolder
'Debug.Print ExecString
'Shell ExecString
'Kill NewFileName
'Kill NewFolder & "\" & ExtractFileName(NewFileName)
End Sub

Private Sub cmdOK_Click()
Me.Hide
TaskbarRefresh = False
IconRefresh = False

SaveSetting "Bos", "BInternet", "ConnectionName", txtDunConnection.Text
If chkShowWebAddress.Value = 1 Then
    SaveSetting "Bos", "BTaskbar", "ShowAddressBar", "True"
    If ShowAddress = False Then TaskbarRefresh = True
Else
    SaveSetting "Bos", "BTaskbar", "ShowAddressBar", "False"
    If ShowAddress = True Then TaskbarRefresh = True
End If

If chkShowPSCButton.Value = 1 Then
    SaveSetting "Bos", "BTaskbar", "ShowPscButton", "True"
    If ShowPSC = False Then TaskbarRefresh = True
Else
    SaveSetting "Bos", "BTaskbar", "ShowPscButton", "False"
    If ShowPSC = True Then TaskbarRefresh = True
End If

a = GetSetting("BoS", "BInterface", "Skin", "BoS Standard")
If a <> cmbSkin.List(cmbSkin.ListIndex) Then
    SaveSetting "BoS", "BInterface", "Skin", cmbSkin.List(cmbSkin.ListIndex)
    ChangeSkin cmbSkin.List(cmbSkin.ListIndex)
    TaskbarRefresh = True
    IconRefresh = True
End If

If optTime(0).Value = True Then
    SaveSetting "Bos", "BSystemTray", "TimeFormat", "12"
Else
    SaveSetting "Bos", "BSystemTray", "TimeFormat", "24"
End If

If chkEnableBToday.Value = 1 Then
    SaveSetting "BoS", "BToday", "Enabled", "True"
Else
    SaveSetting "BoS", "BToday", "Enabled", "False"
End If
SaveSetting "BoS", "BToday", "Zipcode", txtZipcode.Text
 If optControl(0).Value = True Then
    m_strMode$ = 1
  ElseIf optControl(1).Value = True Then
    m_strMode$ = 5
  ElseIf optControl(2).Value = True Then
    m_strMode$ = 3
  ElseIf optControl(3).Value = True Then
    m_strMode$ = 2
  Else
    m_strMode$ = 4
  End If
  Call WriteINIString(m_strScreenName$, "mode", m_strMode$, App.path & "\BAim\aim.ini")
If chkAutoLogin.Value = 1 Then
    SaveSetting "Bos", "BIM", "AutoLogin", "True"
Else
    SaveSetting "Bos", "BIM", "AutoLogin", "False"
End If
If chkImSounds.Value = 1 Then
    SaveSetting "Bos", "BIM", "PlayImSounds", "True"
Else
    SaveSetting "Bos", "BIM", "PlayImSounds", "False"
End If
If TransLevel <> 255 - sliTranslucency.Value Then
    TransLevel = 255 - sliTranslucency.Value
    TaskbarRefresh = True
    IconRefresh = True
End If
SaveSetting "BoS", "BSystem", "Translucency", Str(255 - sliTranslucency.Value)
If IconRefresh Then
    frmLoading.ProgressBar1.Value = 0
    Unload frmBDesktopIcons
End If

If TaskbarRefresh Then
    frmLoading.ProgressBar1.Value = 0
    Unload frmTaskbar
    Unload frmBDesktopIcons
    DoEvents
    frmTaskbar.Show
    frmBDesktopIcons.Show
End If


frmTaskbar.UpdateTime
Unload Me

End Sub

Private Sub Form_Activate()
HideStartMenu
End Sub

Private Sub Form_Load()
dirSkins.path = App.path & "\skins"
For i = 0 To dirSkins.ListCount - 1
    cmbSkin.AddItem ExtractFileName(dirSkins.List(i))
Next
a = GetSetting("BoS", "BInterface", "Skin", "BoS Standard")

For i = 0 To cmbSkin.ListCount - 1
    If a = cmbSkin.List(i) Then
        cmbSkin.ListIndex = i
        Exit For
    End If
Next

txtDunConnection.Text = GetSetting("Bos", "BInternet", "ConnectionName", "")
If GetSetting("Bos", "BSystemTray", "TimeFormat", "12") = "12" Then
    optTime(0).Value = True
Else
    optTime(1).Value = True
End If

If GetSetting("Bos", "BTaskbar", "ShowAddressBar", "True") = "True" Then
    chkShowWebAddress.Value = 1
    ShowAddress = True
Else
    chkShowWebAddress.Value = 0
    ShowAddress = False
End If

If GetSetting("Bos", "BTaskbar", "ShowPSCButton", "True") = "True" Then
    chkShowPSCButton.Value = 1
    ShowPSC = True
Else
    chkShowPSCButton.Value = 0
    ShowPSC = False
End If

If GetSetting("BoS", "BToday", "Enabled", "True") = "True" Then
    chkEnableBToday.Value = 1
Else
    chkEnableBToday.Value = 0
End If
txtZipcode.Text = GetSetting("BoS", "BToday", "Zipcode", "")

If GetSetting("Bos", "BIM", "AutoLogin", "False") = "True" Then
    chkAutoLogin.Value = 1
Else
    chkAutoLogin.Value = 0
End If

If GetSetting("Bos", "BIM", "PlayImSounds", "True") = "True" Then
    chkImSounds.Value = 1
Else
    chkImSounds.Value = 0
End If
sliTranslucency.Value = 255 - Val(GetSetting("BoS", "BSystem", "Translucency", "55"))
sliTranslucency_Change
m_strMode$ = GetINIString(m_strScreenName$, "mode", App.path & "\BAim\aim.ini", "1")
Select Case m_strMode$
  Case "1"
    optControl(0).Value = True
  Case "2"
    optControl(3).Value = True
  Case "3"
    optControl(2).Value = True
  Case "4"
    optControl(4).Value = True
  Case "5"
    optControl(1).Value = True
  Case Else
    optControl(0).Value = True
End Select
    
End Sub

Sub ChangeSkin(SkinName As String)
Dim TmpString As String
Dim tmpArray() As String
If SkinName = "BoS Standard" Then
    SaveSetting "BoS", "BDesktopIcons", "BgColor", "&H00000000&"
    SaveSetting "BoS", "BDesktopIcons", "FgColor", "&H00FFFFFF&"
    SaveSetting "BoS", "BDesktopIcons", "ShadowColor", "&H00000000&"
    SaveSetting "BoS", "BTaskbar", "FgColor", "&H00000000&"
    SaveSetting "BoS", "BTaskbar", "ClockFgColor", "&H00000000&"
    SaveSetting "Bos", "BSystem", "Translucent", "True"
Else
    
    ' Load the icon background color
        Open App.path & "\skins\" & SkinName & "\IconBGColor.txt" For Input As #1
        Line Input #1, TmpString
        tmpArray = Split(TmpString, ",")
        SaveSetting "BoS", "BDesktopIcons", "BgColor", RGB(Val(tmpArray(0)), Val(tmpArray(1)), Val(tmpArray(2)))
        Close #1
    ' Load the icon foreground color
        Open App.path & "\skins\" & SkinName & "\IconFGColor.txt" For Input As #1
        Line Input #1, TmpString
        tmpArray = Split(TmpString, ",")
        SaveSetting "BoS", "BDesktopIcons", "FgColor", RGB(Val(tmpArray(0)), Val(tmpArray(1)), Val(tmpArray(2)))
        Close #1
    ' Load the icon shadow color
        Open App.path & "\skins\" & SkinName & "\IconShadowColor.txt" For Input As #1
        Line Input #1, TmpString
        tmpArray = Split(TmpString, ",")
        SaveSetting "BoS", "BDesktopIcons", "ShadowColor", RGB(Val(tmpArray(0)), Val(tmpArray(1)), Val(tmpArray(2)))
        Close #1
    ' Load the taskbar button foreground color
        Open App.path & "\skins\" & SkinName & "\TaskbarFgColor.txt" For Input As #1
        Line Input #1, TmpString
        tmpArray = Split(TmpString, ",")
        SaveSetting "BoS", "BTaskbar", "FgColor", RGB(Val(tmpArray(0)), Val(tmpArray(1)), Val(tmpArray(2)))
        Close #1
    ' Load the clock foreground color
        Open App.path & "\skins\" & SkinName & "\ClockFgColor.txt" For Input As #1
        Line Input #1, TmpString
        tmpArray = Split(TmpString, ",")
        SaveSetting "BoS", "BTaskbar", "ClockFgColor", RGB(Val(tmpArray(0)), Val(tmpArray(1)), Val(tmpArray(2)))
        Close #1
    ' Load the translucency setting
        Open App.path & "\skins\" & SkinName & "\Translucent.txt" For Input As #1
        Line Input #1, TmpString
        SaveSetting "Bos", "BSystem", "Translucent", TmpString
        Close #1
End If

End Sub

Private Sub TabStrip1_Click()

End Sub

Private Sub sliTranslucency_Change()
picTransTest.Cls
AlphaBlending picTransTest.hdc, 0, 0, 32, 32, picBIcon.hdc, 0, 0, 32, 32, sliTranslucency.Value
End Sub

Private Sub tsSettings_Click()
For i = 0 To picSettings.Count - 1
    picSettings(i).Visible = False
Next
picSettings(tsSettings.SelectedItem.Index - 1).Visible = True
End Sub

