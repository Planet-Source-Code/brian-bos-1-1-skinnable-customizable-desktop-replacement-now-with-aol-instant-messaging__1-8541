VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "User List"
      Height          =   2295
      Left            =   2700
      TabIndex        =   3
      Top             =   240
      Width           =   2055
      Begin VB.ListBox lstUsers 
         Height          =   1425
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   855
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         Top             =   1800
         Width           =   855
      End
   End
   Begin MSComDlg.CommonDialog dlgSound 
      Left            =   120
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "wav (*.wav)|*.wav"
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   2700
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   2700
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   2700
      Width           =   975
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
  Dim strRes As String, lngDo As Long, blnMatch As Boolean
  strRes$ = InputBox("Enter a screen name", "AIM")
  strRes$ = LCase(Replace(strRes$, " ", ""))
  If strRes$ <> "" Then
    For lngDo& = 0 To lstUsers.ListCount
      If lstUsers.List(lngDo&) = strRes$ Then
        MsgBox Chr(34) & strRes$ & Chr(34) & " already exists in the list."
        blnMatch = True
        Exit For
      End If
    Next
    If blnMatch = False Then
      lstUsers.AddItem strRes$
      If optControl(2).Value = True Then
        Call SendProc(2, "toc_add_permit " & strRes$ & " " & Chr(0))
      ElseIf optControl(3).Value = True Then
        Call SendProc(2, "toc_add_deny " & strRes$ & Chr(0))
      End If
    End If
  End If
End Sub

Private Sub cmdApply_Click()
  Dim lngDo As Long
  m_strPDList$ = ""
  If lstUsers.ListCount > 0 Then
    For lngDo& = 0 To lstUsers.ListCount
      m_strPDList$ = m_strPDList$ & " " & lstUsers.List(lngDo&)
    Next
  End If
  m_strPDList$ = Trim(m_strPDList$)
  strSoundSignOn$ = txtSoundSignon.Text
  strSoundSignOff$ = txtSoundSignoff.Text
  strSoundFirstIM$ = txtSoundFirstIM.Text
  strSoundIMIn$ = txtSoundIMin.Text
  strSoundIMOut$ = txtSoundIMout.Text
 
  Call WriteINIString(m_strScreenName$, "signon sound", strSoundSignOn$, App.path & "\aim.ini")
  Call WriteINIString(m_strScreenName$, "signoff sound", strSoundSignOff$, App.path & "\aim.ini")
  Call WriteINIString(m_strScreenName$, "firstim sound", strSoundFirstIM$, App.path & "\aim.ini")
  Call WriteINIString(m_strScreenName$, "imin sound", strSoundIMIn$, App.path & "\aim.ini")
  Call WriteINIString(m_strScreenName$, "imout sound", strSoundIMOut$, App.path & "\aim.ini")
  Call WriteINIString(m_strScreenName$, "pdlist", m_strPDList$, App.path & "\aim.ini")
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdFirstIM_Click()
  dlgSound.ShowOpen
  txtSoundFirstIM.Text = dlgSound.Filename
End Sub

Private Sub cmdIMIn_Click()
  dlgSound.ShowOpen
  txtSoundIMin.Text = dlgSound.Filename
End Sub

Private Sub cmdIMOut_Click()
  dlgSound.ShowOpen
  txtSoundIMout.Text = dlgSound.Filename
End Sub

Private Sub cmdOK_Click()
  cmdApply_Click
  Unload Me
End Sub

Private Sub cmdRemove_Click()
  If lstUsers.ListIndex > -1 Then
    If optControl(2).Value = True Then
      Call SendProc(2, "toc_add_deny " & lstUsers.Text & " " & Chr(0))
    ElseIf optControl(3).Value = True Then
      Call SendProc(2, "toc_add_permit " & lstUsers.Text & Chr(0))
    End If
    lstUsers.RemoveItem lstUsers.ListIndex
  End If
End Sub

Private Sub cmdSignOff_Click()
  dlgSound.ShowOpen
  txtSoundSignoff.Text = dlgSound.Filename
End Sub

Private Sub cmdSignOn_Click()
  dlgSound.ShowOpen
  txtSoundSignon.Text = dlgSound.Filename
End Sub

Private Sub Form_Load()
  Dim lngDo As Long, arrUsers() As String
  If m_strScreenName$ <> "" Then
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
    txtSoundSignon.Text = strSoundSignOn$
    txtSoundSignoff.Text = strSoundSignOff$
    txtSoundFirstIM.Text = strSoundFirstIM$
    txtSoundIMin.Text = strSoundIMIn$
    txtSoundIMout.Text = strSoundIMOut$
    If m_strPDList$ <> "" And m_strMode$ <> "5" Then
      lstUsers.Clear
      arrUsers$() = Split(m_strPDList$, " ")
      For lngDo& = LBound(arrUsers$) To UBound(arrUsers$)
        lstUsers.AddItem arrUsers$(lngDo&)
      Next
    End If
  End If
End Sub
