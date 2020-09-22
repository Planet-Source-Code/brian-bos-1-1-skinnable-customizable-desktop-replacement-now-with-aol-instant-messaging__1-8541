VERSION 5.00
Begin VB.Form frmLogonStatus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Connecting..."
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4110
   Icon            =   "frmLogonStatus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "frmLogonStatus.frx":1272
      Top             =   360
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Please wait while you are connected to the Instant Message server."
      Height          =   495
      Left            =   900
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Closed"
      Height          =   255
      Left            =   1020
      TabIndex        =   0
      Top             =   660
      Width           =   2955
   End
End
Attribute VB_Name = "frmLogonStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
