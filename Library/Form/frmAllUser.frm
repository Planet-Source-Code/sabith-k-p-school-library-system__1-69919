VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAllUser 
   BackColor       =   &H00F5F5F5&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage User Entries"
   ClientHeight    =   5265
   ClientLeft      =   225
   ClientTop       =   480
   ClientWidth     =   7680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAllUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   351
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   512
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   6
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      Top             =   4800
      Width           =   855
   End
   Begin MSComctlLib.ListView listEntries 
      Height          =   4425
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   7805
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ilUser"
      SmallIcons      =   "ilUser"
      ForeColor       =   -2147483640
      BackColor       =   14737632
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "User ID"
         Object.Width           =   38100
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Password"
         Object.Width           =   38100
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Creation Date"
         Object.Width           =   38100
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Created By"
         Object.Width           =   38100
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Modified Date"
         Object.Width           =   38100
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Modified By"
         Object.Width           =   38100
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   0
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   513
      TabIndex        =   1
      Top             =   0
      Width           =   7695
   End
   Begin MSComctlLib.ImageList ilUser 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAllUser.frx":000C
            Key             =   "admin"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAllUser.frx":0EE6
            Key             =   "user"
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000004&
      X1              =   16
      X2              =   504
      Y1              =   312
      Y2              =   312
   End
End
Attribute VB_Name = "frmAllUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Public Sub ShowForm()
        
    'check current user
    If LCase(CurrentUser.UserID) <> "administrator" Then
        MsgBox "You are not permitted to access user entries.", vbExclamation
        Unload Me
        Exit Sub
    End If
        
    Me.Show vbModal
End Sub

Private Sub cmdAdd_Click()
    If frmUserEntry.ShowAdd = True Then
        RefreshUsers
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    If listEntries.ListItems.Count > 0 Then
        If LCase(Trim(listEntries.SelectedItem.Text)) <> "administrator" Then
            If MsgBox("Are you sure you want to delete User '" & listEntries.SelectedItem.Text & "'?", vbQuestion + vbYesNo) = vbYes Then
                If DeleteUser(listEntries.SelectedItem.Text) = True Then
                    RefreshUsers
                Else
                    'WriteErrorLog Me.Name, "cmdDelete_Click", "Falied: DeleteUser(listEntries.SelectedItem.Text) = True"
                End If
            End If
        Else
            MsgBox "Administrator account cannot be deleted", vbExclamation
        End If
    End If
End Sub

Private Sub cmdEdit_Click()
    If listEntries.ListItems.Count > 0 Then
        If frmUserEntry.ShowEdit(listEntries.SelectedItem.Text) = True Then
            RefreshUsers
        End If
    End If
End Sub

Private Sub cmdRefresh_Click()
    RefreshUsers
End Sub

Private Sub Form_Activate()
    RefreshUsers
End Sub




Private Sub RefreshUsers()
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim ci As ComboItem
    
    listEntries.ListItems.Clear
    
    sSQL = "SELECT tblUser.UserID, tblUser.Password, tblUser.CreationDate, tblUser.CreatedBy, tblUser.ModifiedDate, tblUser.ModifiedBy" & _
            " FROM tblUser"


    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog Me.Name, "RefreshUsers", "Unable to connect Recordset. SQL Expression: '" & sSQL & "'"
        GoTo ReleaseAndExit
    End If
    
    If AnyRecordExisted(vRS) = False Then
        MsgBox "The are no User yet. Please add some User first.", vbExclamation
        GoTo ReleaseAndExit
    End If

    vRS.MoveFirst
    
    While vRS.EOF = False
        
        listEntries.ListItems.Add , , ReadField(vRS.Fields("UserID")), IIf(LCase(Trim(ReadField(vRS.Fields("UserID")))) = "administrator", "admin", "user")
        With listEntries.ListItems(listEntries.ListItems.Count)
            .SubItems(1) = ReadField(vRS.Fields("Password"))
            .SubItems(2) = ReadField(vRS.Fields("CreationDate"))
            .SubItems(3) = ReadField(vRS.Fields("CreatedBy"))
            .SubItems(4) = ReadField(vRS.Fields("ModifiedDate"))
            .SubItems(5) = ReadField(vRS.Fields("ModifiedBy"))
        End With
        
        vRS.MoveNext
    Wend
    
    
    
ReleaseAndExit:
    Set vRS = Nothing
End Sub

