VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "by DosFX"
      Height          =   195
      Left            =   540
      TabIndex        =   13
      Top             =   1380
      Width           =   1515
   End
   Begin VB.Label lblInfo 
      Caption         =   "BoS is Copyrighted by BSoft and has been released under the GPL (please see copying.txt). BIM is released as public domain."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Index           =   11
      Left            =   3000
      TabIndex        =   12
      Top             =   2460
      Width           =   3975
   End
   Begin VB.Label lblInfo 
      Caption         =   $"frmAbout.frx":1272
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Index           =   0
      Left            =   3060
      TabIndex        =   11
      Top             =   60
      Width           =   3975
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Brian"
      Height          =   195
      Index           =   10
      Left            =   1140
      TabIndex        =   10
      Top             =   2220
      Width           =   360
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "BoS Author:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   120
      TabIndex        =   9
      Top             =   2220
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "BIM"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   8
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "dosfx or dosrox"
      Height          =   195
      Index           =   8
      Left            =   1140
      TabIndex        =   7
      Top             =   3180
      Width           =   1065
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "www.dosfx.com"
      Height          =   195
      Index           =   7
      Left            =   1140
      TabIndex        =   6
      Top             =   2940
      Width           =   1125
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "chad@dosfx.com"
      Height          =   195
      Index           =   6
      Left            =   1140
      TabIndex        =   5
      Top             =   2700
      Width           =   1245
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Chad J. Cox"
      Height          =   195
      Index           =   5
      Left            =   1140
      TabIndex        =   4
      Top             =   2460
      Width           =   855
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "AIM:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   120
      TabIndex        =   3
      Top             =   3180
      Width           =   360
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "WebSite:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   120
      TabIndex        =   2
      Top             =   2940
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Email:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   120
      TabIndex        =   1
      Top             =   2700
      Width           =   480
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Author:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   2460
      Width           =   615
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
HideStartMenu
End Sub
