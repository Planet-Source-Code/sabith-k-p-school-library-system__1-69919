VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_aed_BookIssue 
   BackColor       =   &H00D2DAD3&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   Icon            =   "frm_aed_BookIssue.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   7440
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAll 
      BackColor       =   &H00D8E9EC&
      Height          =   285
      Index           =   1
      Left            =   4080
      Picture         =   "frm_aed_BookIssue.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   990
      Width           =   345
   End
   Begin VB.TextBox txtID 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdAll 
      BackColor       =   &H00D8E9EC&
      Height          =   285
      Index           =   0
      Left            =   4080
      Picture         =   "frm_aed_BookIssue.frx":0596
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   495
      Width           =   345
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Cancel"
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
      Left            =   3480
      TabIndex        =   9
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
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
      Left            =   2400
      TabIndex        =   8
      Top             =   2640
      Width           =   975
   End
   Begin MSComCtl2.DTPicker DTReturn 
      Height          =   330
      Left            =   1080
      TabIndex        =   7
      Top             =   1920
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   20185091
      CurrentDate     =   39220
   End
   Begin MSComCtl2.DTPicker DTIssue 
      Height          =   330
      Left            =   1080
      TabIndex        =   6
      Top             =   1440
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   20185091
      CurrentDate     =   39220
   End
   Begin VB.TextBox txtBooKId 
      Height          =   325
      Left            =   1080
      TabIndex        =   5
      Top             =   960
      Width           =   3375
   End
   Begin VB.TextBox txtMemberID 
      Height          =   325
      Left            =   1080
      TabIndex        =   4
      Top             =   480
      Width           =   3375
   End
   Begin MSComctlLib.ListView list 
      Height          =   3015
      Left            =   4560
      TabIndex        =   12
      Top             =   0
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483647
      BackColor       =   14737632
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   4440
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Issue"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Book ID"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Member ID"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   945
   End
End
Attribute VB_Name = "frm_aed_BookIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim SIssue As aIssue
Dim mShowAdd As Boolean
Dim mShowEdit As Boolean
Dim mFormState As String
Dim AllIssue As Boolean

Public Function ShowForm() As Boolean

    'show form
    Me.Show vbModal
    
    ShowForm = mShowForm
End Function
    
Public Function ShowEdit(ID As Long) As Boolean
    
    'set form state
    mFormState = "edit"
    
    'set parameter
    SIssue.ID = ID
    
    'show form
    Me.Show vbModal
    
    'return
    ShowEdit = mShowEdit
    
End Function
Public Function ShowAdd() As Boolean
    
      'set form state
    mFormState = "add"
    
    'show form
    Me.Show vbModal
    
    'return
    ShowAdd = mShowAdd
    
End Function


Private Sub cmdAll_Click(Index As Integer)
Select Case Index
    Case 0
        AllIssue = True
        LoadEntries
    Case 1
    AllIssue = False
        LoadEntries
End Select
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
    Select Case mFormState
        Case "add"
            SaveAdd
        Case "edit"
            SaveEdit
    End Select
    End Sub
Private Sub Form_Activate()
Select Case mFormState
        Case "add"
        
            'set caption
            Me.Caption = "New Issue"
            Me.cmdSave.Caption = "&Save"
            
        Case "edit"
            'get info
            If GetIssueID(SIssue.ID, SIssue) = False Then
                'show failed
                MsgBox "User entry with User ID : '" & SIssue.ID & "' does not exist.", vbExclamation
                'close this form
                Unload Me
                Exit Sub
            End If
            
            'set form ui info
        txtID.Text = SIssue.ID
        txtBookID.Text = SIssue.BookID
        txtMemberID.Text = SIssue.MemberID
        DTIssue.Value = SIssue.IssueDate
        DTReturn.Value = SIssue.Returndate
            'set caption
            Me.Caption = "Edit Issue"
            Me.cmdSave.Caption = "&Save"
            Me.cmdSave.Enabled = False
            
            txtID.Locked = True
    End Select
    End Sub


Private Function SaveAdd()
    Dim NewIssue As aIssue
    Dim oldIssue As aIssue
    
   If IsEmpty(txtBookID.Text) Then
        MsgBox "Please Enter the Book ID", vbExclamation
        HLTxt txtBookID
        Exit Function
    End If
    
     If IsEmpty(txtMemberID.Text) Then
        MsgBox "Please Select the Member ID", vbExclamation
        HLTxt txtBookID
        Exit Function
    End If
    
    If UnreturnedBook(txtMemberID.Text) = False Then
    'Exit Function
    End If
    
     If CheckBookStatus(txtBookID.Text) = True Then
        Exit Function
    End If
    If mFormState = "add" Then
       
    'check duplication
'    If GetIssueID(txtID.Text, oldIssue) = True Then
     '   MsgBox "The IssueID that you have entered is already exist." & vbNewLine & vbNewLine & _
            "Please enter different value.", vbExclamation
     ''   HLTxt txtID
      '  Exit Function
    'End If
    'NewIssue.ID = txtID.Text
    NewIssue.BookID = txtBookID.Text
    NewIssue.MemberID = txtMemberID.Text
    NewIssue.IssueDate = DTIssue.Value
    NewIssue.Returndate = DTReturn.Value
      
    'try
    'if modRsIssue.AddmemberS (NewMEmb
    
    If modRsIssue.AddIssue(NewIssue) = True Then
        MsgBox "New Issue entry was successfully created.", vbInformation
        'set flag
        mShowAdd = True
        'close form and return
        Unload Me
    Else
        MsgBox "Unable to add new Issue entry.", vbExclamation
        'set flag
        mShowAdd = False
    End If
    
    End If
End Function

Private Function SaveEdit()

    Dim NewIssue As aIssue
    Dim oldIssue As aIssue
     'check form field
    
   If IsEmpty(txtBookID.Text) Then
        MsgBox "Please Enter the Book ID", vbExclamation
        HLTxt txtBookID
        Exit Function
    End If
    
     If IsEmpty(txtMemberID.Text) Then
        MsgBox "Please Select the Member ID", vbExclamation
        HLTxt txtBookID
        Exit Function
    End If
     
     If UnreturnedBook(txtMemberID.Text) = False Then
    Exit Function
    End If
    
     If CheckBookStatus(txtBookID.Text) = True Then
        Exit Function
    End If
    
    'set new
    NewIssue.ID = txtID.Text
    NewIssue.BookID = txtBookID.Text
    NewIssue.MemberID = txtMemberID.Text
    NewIssue.IssueDate = DTIssue.Value
    NewIssue.Returndate = DTReturn.Value
      
    'try
    'add new Issue
    If modRsIssue.EditIssue(NewIssue) = True Then
        MsgBox "Issue entry was updated.", vbInformation
        'set flag
        mShowEdit = True
        'close form and return
        Unload Me
        
    Else
    
        MsgBox "Unable to update Issue entry.", vbExclamation
        'set flag
        mShowAdd = False
        
    End If
    
End Function
Private Sub LoadEntries()
    On Error GoTo RAE
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim rec_count, X
    Dim l_results As ListItem
        list.ListItems.Clear
    'set SQL Expression
    If AllIssue = True And txtMemberID.Text = "" Then
        sSQL = "SELECT * From tblMember" ' where ID like " & txtBookID.Text '& _
            " ORDER BY tblMember.id"
    ElseIf AllIssue = True And txtMemberID.Text <> "" Then
         sSQL = "SELECT * From tblMember where ID like '" & txtMemberID.Text & "%'"
        
    ElseIf AllIssue = False And txtBookID.Text = "" Then
              sSQL = "SELECT * From tblBooks"
    ElseIf AllIssue = False And txtBookID.Text <> "" Then
        sSQL = "SELECT * From tblbooks where ID like '" & txtBookID.Text & "%'"
    End If
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        MsgBox Err.Description, vbExclamation
    End If
    
   With vRS
     
        rec_count = .RecordCount
    '.MoveLast
    '.MoveFirst
        For X = 1 To rec_count
    
    Set l_results = list.ListItems.Add(X, , !ID)
                    l_results.SubItems(1) = !Name
                    
                   
    .MoveNext
        Next
End With
RAE:
    Set vRS = Nothing
    list.Refresh
End Sub







Private Sub list_DblClick()
If KeyAscii = 13 And AllIssue = True Then
    txtMemberID.Text = list.SelectedItem.Text
    txtMemberID.SetFocus
ElseIf KeyAscii = 13 And AllIssue = False Then
    txtBookID.Text = list.SelectedItem.Text
    txtBookID.SetFocus
End If
End Sub

Private Sub list_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And AllIssue = True Then
    txtMemberID.Text = list.SelectedItem.Text
    txtMemberID.SetFocus
ElseIf KeyAscii = 13 And AllIssue = False Then
    txtBookID.Text = list.SelectedItem.Text
    txtBookID.SetFocus
ElseIf KeyAscii = vbKeyBack And AllIssue = True Then
    txtMemberID.SetFocus
ElseIf KeyAscii = vbKeyBack And AllIssue = False Then
    txtBookID.SetFocus
End If
    
End Sub

Private Sub txtBooKId_Change()
AllIssue = False
LoadEntries
End Sub

Private Sub txtBooKId_GotFocus()
AllIssue = False
LoadEntries
End Sub

Private Sub txtBooKId_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    list.SetFocus
End If
End Sub

Private Sub txtMemberID_Change()
AllIssue = True
LoadEntries
End Sub

Private Sub txtMemberID_GotFocus()
AllIssue = True
LoadEntries
End Sub

Private Sub txtMemberID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    list.SetFocus
End If
End Sub
'''#################################### check Book Status######################################
Public Function CheckBookStatus(BookID As String) As Boolean
        Dim vRS As New ADODB.Recordset
        Dim sSQL As String
        'default
On Error GoTo Err:
        CheckBookStatus = False
        
    sSQL = "SELECT * FROM tblTrans WHERE BookID='" & BookID & "' AND " & " bReturn=False "
        
        If ConnectRS(PrimeDB, vRS, sSQL) = True Then
            'MsgBox "The book is alredy issued to another Member" & vbCrLf & " Select another Book and Try Again", vbExclamation
            'CheckBookStatus = True
            'Exit Function
    End If
    BookID = vRS.Fields("BookID")
    BReturn = vRS.Fields("bReturn")
    MsgBox "The book is alredy issued to another Member" & vbCrLf & " Select another Book and Try Again", vbExclamation
CheckBookStatus = True
Exit Function
Err:
    Set vRS = Nothing
    CheckBookStatus = False
End Function
''''##############################################################################################

'''''###############################Check Previous un returndBook information of a Member ###################
Public Function UnreturnedBook(MemberID As String) As Boolean
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim msg As Long
    
    UnreturnedBook = False
    
    sSQL = "SELECT * FROM tblTrans WHERE MemberID='" & MemberID & "' AND " & " IDate <" & CLng(Date) & " AND " & " bReturn=False"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        MsgBox Err.Description, vbExclamation
    End If
    
    If vRS.RecordCount = 0 Then
        Exit Function
        'UnreturnedBook = True
    End If
    
        With vRS
            For i = 1 To .RecordCount
                msg = msg + 1
            Next
        End With
       If MsgBox("This Member Have   " & msg & "   Books return Pending" & vbCrLf & " Do you want to Continnue ?", vbYesNo + vbExclamation) = vbNo Then
        Exit Function
       End If
    UnreturnedBook = True
            

End Function

