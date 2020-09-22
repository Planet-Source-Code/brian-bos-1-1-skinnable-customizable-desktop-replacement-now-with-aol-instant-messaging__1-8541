VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCopyFiles 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BInstalller"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5130
   Icon            =   "frmCopyFiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   2280
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1620
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   109
   End
   Begin VB.Label Label2 
      Caption         =   "Copying << File Name >> - File <#> of 109"
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
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "Please wait while BInstaller copies files to your hard drive. This may take a few minutes."
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
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4455
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   120
      Picture         =   "frmCopyFiles.frx":0442
      Top             =   60
      Width           =   480
   End
End
Attribute VB_Name = "frmCopyFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Files(0 To 108) As String

Private Sub Form_Load()
On Error GoTo err
File1.Path = App.Path & "\Install Files"
MkDir frmSelectDir.txtDir.Text
Me.Show
Me.Refresh
For i = 0 To File1.ListCount - 1
    FileCopy App.Path & "\" & File1.List(i), frmSelectDir.txtDir.Text & "\" & File1.List(i)
    Label2.Caption = "Copying " & File1.List(i) & " - File " & i & " of 109"
    ProgressBar1.Value = i + 1
    Label2.Refresh
Next
MkDir frmSelectDir.txtDir.Text & "\skins"
MkDir frmSelectDir.txtDir.Text & "\skins\Red Shades"
MkDir frmSelectDir.txtDir.Text & "\skins\Green Shades"

File1.Path = App.Path & "\Install Files\Skins"
Exit Sub
err:
MsgBox "Error " & err.Number & " - " & err.Description, vbOKOnly Or vbCritical, "Error"
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
If CopyDone = False Then Cancel = True
End Sub

