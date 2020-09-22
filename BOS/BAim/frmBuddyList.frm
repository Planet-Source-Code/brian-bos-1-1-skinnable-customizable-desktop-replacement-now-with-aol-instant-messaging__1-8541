VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmBuddyList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BIM Buddy List"
   ClientHeight    =   5295
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   2655
   Icon            =   "frmBuddyList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   2655
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgList 
      Left            =   1680
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.ImageList imgBuddy 
      Left            =   1560
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuddyList.frx":1272
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuddyList.frx":13CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuddyList.frx":1526
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab tabBuddy 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   8916
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Online"
      TabPicture(0)   =   "frmBuddyList.frx":27A8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tvwBuddies"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdIM"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdInvite"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdChat"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "List Setup"
      TabPicture(1)   =   "frmBuddyList.frx":27C4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdDelete"
      Tab(1).Control(1)=   "cmdAddGroup"
      Tab(1).Control(2)=   "cmdAddBuddy"
      Tab(1).Control(3)=   "tvwSetup"
      Tab(1).ControlCount=   4
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   375
         Left            =   -74880
         TabIndex        =   4
         Top             =   4560
         Width           =   2175
      End
      Begin VB.CommandButton cmdAddGroup 
         Caption         =   "Add Group"
         Height          =   375
         Left            =   -74880
         TabIndex        =   5
         Top             =   4140
         Width           =   2175
      End
      Begin VB.CommandButton cmdChat 
         Caption         =   "Chat"
         Height          =   375
         Left            =   1620
         TabIndex        =   1
         Top             =   4560
         Width           =   675
      End
      Begin VB.CommandButton cmdInvite 
         Caption         =   "Invite"
         Height          =   375
         Left            =   840
         TabIndex        =   2
         Top             =   4560
         Width           =   675
      End
      Begin VB.CommandButton cmdAddBuddy 
         Caption         =   "Add Buddy"
         Height          =   375
         Left            =   -74880
         TabIndex        =   6
         Top             =   3720
         Width           =   2175
      End
      Begin VB.CommandButton cmdIM 
         Caption         =   "IM"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   4560
         Width           =   615
      End
      Begin MSComctlLib.TreeView tvwSetup 
         Height          =   3135
         Left            =   -74880
         TabIndex        =   7
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   5530
         _Version        =   393217
         Indentation     =   88
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "imgBuddy"
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView tvwBuddies 
         Height          =   3975
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   7011
         _Version        =   393217
         Indentation     =   88
         Style           =   7
         ImageList       =   "imgBuddy"
         Appearance      =   1
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuGetAccount 
         Caption         =   "Get an &AIM account"
      End
      Begin VB.Menu mnuFileSignOff 
         Caption         =   "&Sign Off"
      End
   End
   Begin VB.Menu mnuPeople 
      Caption         =   "&People"
      Begin VB.Menu cmdSetupList 
         Caption         =   "Setup &Buddy List"
      End
      Begin VB.Menu mnuPeopleJoinChat 
         Caption         =   "&Join Chat"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About BIM"
      End
   End
End
Attribute VB_Name = "frmBuddyList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddBuddy_Click()
  Dim nodBuddy As Node
  If ExistsInTree(tvwSetup, "New Buddy", True) = False Then
    If tvwSetup.Nodes.Count < 1 Then
      MsgBox "You need a group to add buddies to.", vbOKOnly + vbCritical, "Error"
      Exit Sub
    End If
    If tvwSetup.SelectedItem Is Nothing Then
      Set nodBuddy = tvwSetup.Nodes.Add(1, tvwChild, , "New Buddy", 3, 3)
    Else
      If tvwSetup.SelectedItem.Parent Is Nothing Then
        Set nodBuddy = tvwSetup.Nodes.Add(tvwSetup.SelectedItem.Index, tvwChild, , "New Buddy", 3, 3)
      Else
        Set nodBuddy = tvwSetup.Nodes.Add(tvwSetup.SelectedItem.Index, tvwPrevious, , "New Buddy", 3, 3)
      End If
    End If
    nodBuddy.Selected = True
    tvwSetup.SetFocus
    tvwSetup.StartLabelEdit
  End If
End Sub

Private Sub cmdAddGroup_Click()
  Dim lngCounter As Long, strKey As String, nodGroup As Node
  If ExistsInTree(tvwSetup, "New Group", True) = False Then
    If tvwSetup.SelectedItem Is Nothing Then
      Set nodGroup = tvwSetup.Nodes.Add(, , , "New Group", 1, 1)
    Else
      If tvwSetup.SelectedItem.Parent Is Nothing Then
        Set nodGroup = tvwSetup.Nodes.Add(tvwSetup.SelectedItem.Index, tvwNext, , "New Group", 1, 1)
      Else
        Set nodGroup = tvwSetup.Nodes.Add(tvwSetup.SelectedItem.Parent.Index, tvwNext, , "New Group", 1, 1)
      End If
    End If
    nodGroup.Selected = True
    tvwSetup.SetFocus
    tvwSetup.StartLabelEdit
  End If
End Sub

Private Sub cmdDelete_Click()
  Dim lngDo As Long
  If tvwSetup.SelectedItem Is Nothing Then Exit Sub
  If Not tvwSetup.SelectedItem.Parent Is Nothing Then
    Call SendProc(2, "toc_remove_buddy " & Chr(34) & Replace(tvwSetup.SelectedItem.Text, " ", "") & Chr(34) & Chr(0))
  Else
    If tvwSetup.SelectedItem.Children > 0 Then
      For lngDo& = 1 To tvwSetup.SelectedItem.Children
        Call SendProc(2, "toc_remove_buddy " & Chr(34) & Replace(tvwSetup.SelectedItem.Child.FirstSibling, " ", "") & Chr(34) & Chr(0))
        DoEvents
        tvwSetup.Nodes.Remove (tvwSetup.SelectedItem.Child.FirstSibling.Index)
      Next
    End If
  End If
  tvwSetup.Nodes.Remove (tvwSetup.SelectedItem.Index)
End Sub

Private Sub cmdInvite_Click()
  frmInvite.Show
End Sub

Private Sub cmdIM_Click()
  Dim lngFormIndex As Long
  If Not tvwBuddies.SelectedItem Is Nothing Then
    If Not tvwBuddies.SelectedItem.Parent Is Nothing Then
      lngFormIndex& = FormByTag(LCase(Replace(tvwBuddies.SelectedItem.Text, " ", "")))
      If lngFormIndex& > -1 Then
        Forms(lngFormIndex&).SetFocus
      Else
        Dim frmNewIM As New frmIM
        With frmNewIM
          .Caption = tvwBuddies.SelectedItem.Text
          .Tag = LCase(Replace(tvwBuddies.SelectedItem.Text, " ", ""))
          .Show
        End With
      End If
    End If
  End If
End Sub

Private Sub cmdChat_Click()
  frmChatJoin.Show
End Sub

Private Sub Form_Activate()
HideStartMenu
End Sub

Private Sub Form_Load()
  Me.Left = Screen.Width - Me.Width
  Me.Top = 200
  Dim lngDo As Long, nod() As Node
  m_strMode$ = GetINIString(m_strScreenName$, "mode", App.path & "\BAim\aim.ini", "1")
  Call LoadBuddies(m_strScreenName$)
  For lngDo& = 1 To tvwSetup.Nodes.Count
    If tvwSetup.Nodes.Item(lngDo&).Parent Is Nothing Then
      ReDim Preserve nod(tvwBuddies.Nodes.Count + 1)
      Set nod(tvwBuddies.Nodes.Count) = tvwBuddies.Nodes.Add(, , , tvwSetup.Nodes.Item(lngDo&).Text, 1)
    End If
  Next
  m_strPDList$ = GetINIString(m_strScreenName$, "pdlist", App.path & "\BAim\aim.ini", "")
  Me.Show
  Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call SaveBuddies(m_strScreenName$)
  Call WriteINIString(m_strScreenName$, "mode", m_strMode$, App.path & "\aim.ini")
  frmSignOn.wskAIM.Close
End Sub


Private Sub mnuFileSignOff_Click()
  Unload Me
End Sub

Private Sub mnuGetAccount_Click()
ShellExecute Me.hWnd, "open", "http://aim.aol.com/aimnew/Aim/register.adp?promo=106723&pageset=Aim&client=no", "", "", 0
End Sub

Private Sub mnuHelpAbout_Click()
  frmAbout.Show
End Sub

Private Sub mnuPeopleJoinChat_Click()
  frmChatJoin.Show
End Sub

Private Sub tvwBuddies_DblClick()
  Dim lngFormIndex As Long
  If Not tvwBuddies.SelectedItem Is Nothing Then
    If Not tvwBuddies.SelectedItem.Parent Is Nothing Then
      lngFormIndex& = FormByTag(LCase(Replace(tvwBuddies.SelectedItem.Text, " ", "")))
      If lngFormIndex& > -1 Then
        Forms(lngFormIndex&).SetFocus
      Else
        Dim frmNewIM As New frmIM
        With frmNewIM
          .Caption = tvwBuddies.SelectedItem.Text
          .Tag = LCase(Replace(tvwBuddies.SelectedItem.Text, " ", ""))
          .Show
        End With
      End If
    End If
  End If
End Sub

Private Sub tvwSetup_AfterLabelEdit(Cancel As Integer, NewString As String)
  Dim nodGroup As Node
  If Trim(NewString) = "" Then
    MsgBox "Item can not be nothing.", vbCritical + vbOKOnly, "Error"
    tvwSetup.Nodes.Remove (tvwSetup.SelectedItem.Index)
  ElseIf IsValidItem(NewString$) = False Then
    MsgBox "Item can contain only letters, numbers, and spaces.", vbCritical + vbOKOnly, "Error"
    tvwSetup.Nodes.Remove (tvwSetup.SelectedItem.Index)
  ElseIf ExistsInTree(tvwSetup, NewString$) = True Then
    MsgBox Chr(34) & NewString$ & Chr(34) & "Already exists.", vbCritical + vbOKOnly, "Error"
    tvwSetup.Nodes.Remove (tvwSetup.SelectedItem.Index)
  Else
    If Not tvwSetup.SelectedItem.Parent Is Nothing Then
      If ExistsInTree(tvwBuddies, tvwSetup.SelectedItem.Text, False, True) = True Then
        Call SendProc(2, "toc_remove_buddy " & Chr(34) & Replace(tvwSetup.SelectedItem.Text, " ", "") & Chr(34) & Chr(0))
      End If
      Call SendProc(2, "toc_add_buddy " & Replace(NewString, " ", "") & Chr(0))
    Else
      If ExistsInTree(tvwBuddies, tvwSetup.SelectedItem.Text, False, False, NewString$) = False Then
        Set nodGroup = tvwBuddies.Nodes.Add(, , , NewString$, 1, 1)
      End If
    End If
  End If
End Sub

Private Function IsValidItem(strItem As String) As Boolean
  Dim lngDo As Long, blnIsValid As Boolean, strChar As String
  blnIsValid = True
  For lngDo& = 1 To Len(strItem$)
    strChar$ = Mid(strItem$, lngDo&, 1)
    If Asc(strChar$) < 65 Or Asc(strChar$) > 90 Then
      If Asc(strChar$) < 97 Or Asc(strChar$) > 122 Then
        If IsNumeric(strChar$) = False Then
          If strChar$ <> " " Then
            blnIsValid = False
            Exit For
          End If
        End If
      End If
    End If
  Next
  IsValidItem = blnIsValid
End Function

Private Sub tvwSetup_DblClick()
  If tvwSetup.SelectedItem Is Nothing Then Exit Sub
  tvwSetup.StartLabelEdit
End Sub
