VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_aed_BookReturn 
   BackColor       =   &H00D2DAD3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Book Return"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   Icon            =   "frm_aed_BookReturn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   6690
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAll 
      BackColor       =   &H00D8E9EC&
      Height          =   285
      Index           =   0
      Left            =   3840
      Picture         =   "frm_aed_BookReturn.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   240
      Width           =   345
   End
   Begin VB.CommandButton cmdAll 
      BackColor       =   &H00D8E9EC&
      Height          =   285
      Index           =   1
      Left            =   3840
      Picture         =   "frm_aed_BookReturn.frx":0596
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   735
      Width           =   345
   End
   Begin VB.CommandButton cmdCancel 
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
      Left            =   3240
      TabIndex        =   7
      Top             =   1920
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
      Left            =   2160
      TabIndex        =   6
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox txtBookID 
      Height          =   325
      Left            =   1200
      TabIndex        =   4
      Top             =   720
      Width           =   3015
   End
   Begin VB.TextBox txtMemberID 
      Height          =   325
      Left            =   1200
      TabIndex        =   3
      Top             =   240
      Width           =   3015
   End
   Begin MSComCtl2.DTPicker DTReturn 
      Height          =   330
      Left            =   1200
      TabIndex        =   5
      Top             =   1200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   20316163
      CurrentDate     =   39220
   End
   Begin MSComctlLib.ListView list 
      Height          =   2295
      Left            =   4320
      TabIndex        =   8
      Top             =   0
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   4048
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
      Icons           =   "ilRecordIco"
      SmallIcons      =   "ilRecordIco"
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
      X1              =   120
      X2              =   4200
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Return Date"
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
      Top             =   1320
      Width           =   1035
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
      Top             =   840
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
      Top             =   360
      Width           =   945
   End
End
Attribute VB_Name = "frm_aed_BookReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AllIssue As Boolean
Public Function ShowForm() As Boolean

    'show form
    Me.Show vbModal
    
    ShowForm = mShowForm
End Function
    
Public Function UpdateReturn() As Boolean
Dim oldReturn As aReturn
        Dim NewReturn As aReturn
             
UpdateReturn = False

    If GetReturnInform(txtMemberID, txtMemberID, True, oldReturn) = False Then
            MsgBox "Error During the time of Book Return,Check MemberID and BookId that You Enterd?:", vbExclamation
    End If
        
        NewReturn.MemberID = txtMemberID.Text
        NewReturn.BookID = txtBookID.Text
        NewReturn.BReturn = True
        NewReturn.Returndate = DTReturn.Value
    
    If modRSReturn.InsertReturn(NewReturn) = True Then
            MsgBox "MemberID : " & NewReturn.MemberID & vbCrLf & "BookID :" & NewReturn.BookID & vbCrLf & "Return  Record Save Succesfuly", vbInformation
            UpdateReturn = False
    Else
    
    MsgBox "Unable to Update Return Entry", vbExclamation
        UpdateReturn = False
    End If
        
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

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
If LoadTransactionRecord = False Then
    Exit Sub
End If
UpdateReturn
End Sub

Public Function LoadTransactionRecord() As Boolean

Dim vReturn As aReturn
Dim oldReturn As aReturn
Dim NewReturn As aReturn

LoadTransactionRecord = False

        If GetReturnInform(txtMemberID, txtBookID, False, vReturn) = False Then
            MsgBox "MemberID : " & txtMemberID & " and BookID : " & txtBookID & " that you have Enter is incorrect please Check MemberID and BookID", vbExclamation
            Exit Function
        End If
    'NewReturn.ID = oldReturn.ID
    NewReturn.MemberID = oldReturn.MemberID
    NewReturn.BookID = oldReturn.BookID
    NewReturn.BReturn = oldReturn.BReturn
    NewReturn.Returndate = oldReturn.Returndate
    
LoadTransactionRecord = True
End Function

Private Sub Command2_Click()

End Sub

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

