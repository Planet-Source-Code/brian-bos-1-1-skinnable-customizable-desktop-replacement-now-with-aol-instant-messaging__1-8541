VERSION 5.00
Begin VB.Form frmError 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Error"
   ClientHeight    =   930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   Icon            =   "frmError.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   930
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3300
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmError.frx":1272
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lblErrorType 
      Caption         =   "Error Type:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   60
      Width           =   2775
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "You have been connecting and disconnecting too frequently."
      Height          =   390
      Left            =   720
      TabIndex        =   1
      Top             =   360
      Width           =   2775
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
   frmSignOn.Show
   Unload Me
End Sub
Private Sub Form_Activate()
HideStartMenu
End Sub

Private Sub Form_Load()
Me.Height = lblInfo.Top + lblInfo.Height + 500
End Sub
