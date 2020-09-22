VERSION 5.00
Begin VB.Form frmLiscenceAgreement 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Liscence Agreement"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
   Icon            =   "frmLiscenceAgreement.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   410
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   315
      Left            =   7980
      TabIndex        =   2
      Top             =   5760
      Width           =   975
   End
   Begin VB.TextBox txtAgreement 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5235
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   420
      Width           =   9015
   End
   Begin VB.Label Label1 
      Caption         =   "Please read the below liscence agreement:"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmLiscenceAgreement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AllText As String, LineOfText As String, GPLPath As String

Private Sub Command1_Click()
 Unload Me
End Sub

Private Sub Form_Load()
    If GetSetting("BoS", "GPL", "Read", "False") = True Then
        Unload Me
    Else
        GPLPath = App.path & "\Copying.txt"
        Open GPLPath For Input As #1
        Do Until EOF(1)          'then read lines from file
            Line Input #1, LineOfText
            AllText = AllText & LineOfText & vbCrLf
        Loop
        txtAgreement.Text = AllText  'display file
        Close #1                 'close file
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSetting "BoS", "GPL", "Read", "True"
Load frmLoading
End Sub
