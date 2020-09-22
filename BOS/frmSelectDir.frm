VERSION 5.00
Begin VB.Form frmSelectDir 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BInstaller"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5130
   Icon            =   "frmSelectDir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "<< Previous"
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
      Left            =   2280
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txtDir 
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
      Left            =   180
      TabIndex        =   2
      Text            =   "C:\program files\bos"
      Top             =   1200
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next >>"
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
      Left            =   3720
      TabIndex        =   0
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmSelectDir.frx":0442
      Top             =   180
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Please select the directory to install BoS in:"
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
      Left            =   180
      TabIndex        =   1
      Top             =   840
      Width           =   4455
   End
End
Attribute VB_Name = "frmSelectDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Right(txtDir.Text, 1) = "\" Then txtDir.Text = Left(txtDir.Text, Len(txtDir.Text) - 1)
On Error GoTo DirExists
    File1.Path = txtDir.Text
    If MsgBox("The specified directoty already exists. Do you want to overwrite it?", vbYesNo Or vbQuestion, "Delete Direcory?") = vbNo Then Exit Sub
    
    
DirExists:
err.Clear
frmSetAsShell.Show
Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
