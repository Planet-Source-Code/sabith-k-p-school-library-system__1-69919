VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIFrm 
   BackColor       =   &H8000000C&
   ClientHeight    =   6915
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10005
   Icon            =   "MDIFrm.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIFrm.frx":628A
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImgList32 
      Left            =   480
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrm.frx":26875
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrm.frx":2754F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrm.frx":283A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrm.frx":2907B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrm.frx":29D55
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrm.frx":2AA2F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrm.frx":2B709
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrm.frx":2C3E3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImgList32"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Issue Book"
            Object.ToolTipText     =   "Issue Book"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Return Book"
            Object.ToolTipText     =   "Return Book"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Separator"
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Book Records"
            Object.ToolTipText     =   "Book Records"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Member Records"
            Object.ToolTipText     =   "Member Records"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Reports"
            Object.ToolTipText     =   "Reminder"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Separator"
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "About"
            Object.ToolTipText     =   "About"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuUser 
      Caption         =   "User"
      Begin VB.Menu mnumngUser 
         Caption         =   "Manage User"
      End
      Begin VB.Menu mnuLogoff 
         Caption         =   "Log off"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuAction 
      Caption         =   "Action"
   End
   Begin VB.Menu mnuRecords 
      Caption         =   "Records"
      Begin VB.Menu mnuMember 
         Caption         =   "Manage Member"
      End
      Begin VB.Menu mnuBooks 
         Caption         =   "Manage Books"
      End
      Begin VB.Menu mnuManagebookIssue 
         Caption         =   "Manage Book Issue"
         Begin VB.Menu mnuAllissue 
            Caption         =   "All Books Issue"
         End
         Begin VB.Menu mnuIssue 
            Caption         =   "New Books Issue"
         End
      End
      Begin VB.Menu mnuBooksReturn 
         Caption         =   "Books Return"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Reports"
      Begin VB.Menu mnuBookReport 
         Caption         =   "All Book Report"
      End
      Begin VB.Menu mnuMemberReport 
         Caption         =   "All Member Report"
      End
      Begin VB.Menu mnuPendingReport 
         Caption         =   "Unreturned Book Report"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuDatabase 
         Caption         =   "DataBase Utilities"
         Begin VB.Menu mnuBackup 
            Caption         =   "DataBase BackUp"
         End
         Begin VB.Menu mnuRestore 
            Caption         =   "DateBase Restore"
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "MDIFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub fd_Click()
frmAllBooks.ShowForm
End Sub

Private Sub Command1_Click()
frmReminder.Show
End Sub

Private Sub MDIForm_Activate()
            mnuLogoff.Caption = "Logoff  " & CurrentUser.UserID
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
OpenMySite
End Sub

Private Sub mnuAbout_Click()
frmSplash.ShowForm
End Sub

Private Sub mnuAllissue_Click()
frmAllIssues.ShowForm
End Sub

Private Sub mnuBackup_Click()
frmDBBackup.ShowForm
End Sub

Private Sub mnuBookReport_Click()
ShowReport "tblBookS", DataReport1
DataReport1.Show
End Sub

Private Sub mnuBooks_Click()
frmAllBooks.ShowForm
End Sub

Private Sub mnuBooksReturn_Click()
frm_aed_BookReturn.ShowForm
End Sub

Private Sub mnuExit_Click()
OpenMySite
Unload Me
End Sub

Private Sub mnuIssue_Click()
frm_aed_BookIssue.ShowAdd
End Sub

Private Sub mnuLogoff_Click()
frmLogin.ShowForm
End Sub

Private Sub mnuMember_Click()
frmAllMembers.ShowForm
End Sub

Private Sub mnuMemberReport_Click()
ShowReport "tblMember", DataReport2
DataReport2.Show
End Sub

Private Sub mnumngUser_Click()
frmAllUser.ShowForm
End Sub

Private Sub mnuPendingReport_Click()
ShowReport "tblTrans", DataReport3, " WHERE BReturn=False And rDate <" & CLng(Date)
DataReport3.Show
End Sub

Private Sub mnuRestore_Click()
frmRestore.ShowForm
End Sub

Public Function ShowForm()
    
    'default
    bUserLoggedOn = False
        
    
    'show form
    Me.WindowState = vbMaximized
    Me.Show
    DoEvents
   
    'show weclome
    'frmWelcome.ShowForm
    
    'unload splash
    frmSplash.UnloadSplash
    
BeginLogin:
    If AnyUserExist = False Then
        If frmUserEntry.ShowAddAdmin = False Then
            Unload Me
        Else
            GoTo BeginLogin
        End If
    Else
        If frmLogin.ShowForm = False Then
            Unload Me
            Exit Function
        End If
    End If
    'If LCase(Trim(CurrentUser.UserID)) <> "administrator" Then
        
      '  End If


End Function

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1: frm_aed_BookIssue.ShowAdd
    Case 2: frm_aed_BookReturn.ShowForm
    Case 4: frm_aed_Books.ShowAdd
    Case 5: frm_aed_Member.ShowAdd
    Case 6: frmReminder.ShowForm
    'Case 8: mnuSettings_Click
    Case 9: mnuAbout_Click
    End Select

End Sub
