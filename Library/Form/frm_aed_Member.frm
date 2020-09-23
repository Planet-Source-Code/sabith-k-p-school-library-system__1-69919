VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_aed_Member 
   BackColor       =   &H00D2DAD3&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6450
   Icon            =   "frm_aed_Member.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   6450
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   330
      Left            =   4800
      TabIndex        =   15
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      Format          =   20316161
      CurrentDate     =   39220
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
      Left            =   5400
      TabIndex        =   13
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
      Left            =   4200
      TabIndex        =   12
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox txtAge 
      Height          =   325
      Left            =   1080
      TabIndex        =   11
      Top             =   1080
      Width           =   3135
   End
   Begin VB.TextBox txtDivision 
      Height          =   325
      Left            =   3120
      TabIndex        =   10
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtClass 
      Height          =   325
      Left            =   1080
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtAddress 
      Height          =   325
      Left            =   1080
      TabIndex        =   8
      Top             =   2040
      Width           =   5295
   End
   Begin VB.TextBox txtName 
      Height          =   325
      Left            =   1080
      TabIndex        =   7
      Top             =   600
      Width           =   3135
   End
   Begin VB.TextBox txtID 
      Height          =   325
      Left            =   1080
      TabIndex        =   5
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
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
      Left            =   4200
      TabIndex        =   14
      Top             =   240
      Width           =   405
   End
   Begin VB.Line Line1 
      X1              =   600
      X2              =   6600
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   690
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
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
      Left            =   2400
      TabIndex        =   4
      Top             =   1680
      Width           =   660
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Class"
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
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   435
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Age"
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
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   330
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   195
   End
End
Attribute VB_Name = "frm_aed_Member"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SMember As aMember
Dim mShowAdd As Boolean
Dim mShowEdit As Boolean
Dim mFormState As String

Public Function ShowForm() As Boolean

    'show form
    Me.Show vbModal
    
    ShowForm = mShowForm
End Function
    
Public Function ShowEdit(ID As String) As Boolean
    
    'set form state
    mFormState = "edit"
    
    'set parameter
    SMember.ID = ID
    
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
            Me.Caption = "Add Member"
            Me.cmdSave.Caption = "&Save"
            
        Case "edit"
            'get info
            If GetMemberID(SMember.ID, SMember) = False Then
                'show failed
                MsgBox "User entry with User ID : '" & SMember.ID & "' does not exist.", vbExclamation
                'close this form
                Unload Me
                Exit Sub
            End If
            
            'set form ui info
        txtID.Text = SMember.ID
        txtName.Text = SMember.Name
        txtAge.Text = SMember.Age
        txtClass.Text = SMember.Class
        txtDivision.Text = SMember.Division
        txtAddress.Text = SMember.Address
        DTPicker1.Value = SMember.mDate
            
                       
            'set caption
            Me.Caption = "Edit Member"
            Me.cmdSave.Caption = "&Update"
            
            txtID.Locked = True
    End Select
    End Sub


Private Function SaveAdd()
    Dim NewMember As aMember
    Dim oldMember As aMember
    
    
    'check form field
    If IsEmpty(txtID.Text) Then
        MsgBox "Please enter Member ID", vbExclamation
       HLTxt txtID
        Exit Function
    End If
   If IsEmpty(txtName.Text) Then
        MsgBox "Please Enter the Member Name", vbExclamation
        HLTxt txtName
        Exit Function
    End If
    
    If mFormState = "add" Then
       
    'check duplication
    If GetMemberID(txtID.Text, oldMember) = True Then
        MsgBox "The MemberID that you have entered is already exist." & vbNewLine & vbNewLine & _
            "Please enter different value.", vbExclamation
        HLTxt txtID
        Exit Function
    End If
    NewMember.ID = txtID.Text
    NewMember.Name = txtName.Text
    NewMember.Age = Val(txtAge.Text)
    NewMember.Class = txtClass.Text
    NewMember.Division = txtDivision.Text
    NewMember.Address = txtAddress.Text
    NewMember.mDate = DTPicker1.Value
      
    'try
    
    If ModRsMember.AddMember(NewMember) = True Then
        MsgBox "New Member entry was successfully created.", vbInformation
        'set flag
        mShowAdd = True
        'close form and return
        Unload Me
        Me.ShowAdd
        
        
    Else
    
        MsgBox "Unable to add new Member entry.", vbExclamation
        'set flag
        mShowAdd = False
        
    End If
    End If
End Function

Private Function SaveEdit()

    Dim NewMember As aMember
    Dim oldMember As aMember
    
    'check form field
    
    
      If IsEmpty(txtID.Text) Then
        MsgBox "Please enter Member ID", vbExclamation
       HLTxt txtID
        Exit Function
    End If
   If IsEmpty(txtName.Text) Then
        MsgBox "Please Enter the Member Name", vbExclamation
        HLTxt txtName
        Exit Function
    End If
    
    
    'set new
    NewMember.ID = txtID.Text
    NewMember.Name = txtName.Text
    NewMember.Age = txtAge.Text
    NewMember.Class = txtClass.Text
    NewMember.Division = txtDivision.Text
    NewMember.Address = txtAddress.Text
    NewMember.mDate = DTPicker1.Value
      
    'try
    'add new Member
    If ModRsMember.EditMember(NewMember) = True Then
        MsgBox "Member entry was updated.", vbInformation
        'set flag
        mShowEdit = True
        'close form and return
        Unload Me
        
    Else
    
        MsgBox "Unable to update Member entry.", vbExclamation
        'set flag
        mShowAdd = False
        
    End If
    
End Function







