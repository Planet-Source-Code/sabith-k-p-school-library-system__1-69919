VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_aed_Books 
   BackColor       =   &H00D2DAD3&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   Icon            =   "frmNewBooks.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
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
      Left            =   4440
      TabIndex        =   15
      Top             =   3120
      Width           =   1095
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
      Left            =   3120
      TabIndex        =   14
      Top             =   3120
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DtDate 
      Height          =   330
      Left            =   3960
      TabIndex        =   13
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   20316163
      CurrentDate     =   39220
   End
   Begin VB.TextBox txtAuthor 
      Height          =   325
      Left            =   1200
      TabIndex        =   11
      Top             =   2040
      Width           =   4335
   End
   Begin VB.TextBox txtPrice 
      Height          =   325
      Left            =   1200
      TabIndex        =   10
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox txtName 
      Height          =   325
      Left            =   1200
      TabIndex        =   9
      Top             =   600
      Width           =   4335
   End
   Begin VB.TextBox txtPublication 
      Height          =   325
      Left            =   1200
      TabIndex        =   8
      Top             =   1080
      Width           =   4335
   End
   Begin VB.TextBox txtSubject 
      Height          =   325
      Left            =   1200
      TabIndex        =   7
      Top             =   1560
      Width           =   4335
   End
   Begin VB.TextBox txtID 
      Height          =   325
      Left            =   1200
      TabIndex        =   6
      Top             =   120
      Width           =   2175
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5640
      Y1              =   3000
      Y2              =   3000
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
      Left            =   3480
      TabIndex        =   12
      Top             =   120
      Width           =   405
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
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
      TabIndex        =   5
      Top             =   2640
      Width           =   420
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Publication"
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
      TabIndex        =   4
      Top             =   1080
      Width           =   930
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject"
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
      Top             =   1560
      Width           =   645
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Author"
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
      Top             =   2040
      Width           =   585
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
      Left            =   120
      TabIndex        =   1
      Top             =   600
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   195
   End
End
Attribute VB_Name = "frm_aed_Books"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SBooks As aBooks
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
    SBooks.ID = ID
    
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
            Me.Caption = "Add Books"
            Me.cmdSave.Caption = "&Save"
            
        Case "edit"
            'get info
            If GetBooksID(SBooks.ID, SBooks) = False Then
                'show failed
                MsgBox "User entry with User ID : '" & SBooks.ID & "' does not exist.", vbExclamation
                'close this form
                Unload Me
                Exit Sub
            End If
            
            'set form ui info
        txtID.Text = SBooks.ID
        txtName.Text = SBooks.Name
        txtPublication.Text = SBooks.Publisher
        txtSubject.Text = SBooks.Subject
        txtAuthor.Text = SBooks.Author
        txtPrice.Text = SBooks.Price
        'txtNoofBooks.Text = SBooks.NoofBooks
        DtDate.Value = SBooks.bDate
            
                       
            'set caption
            Me.Caption = "Edit Books"
            Me.cmdSave.Caption = "&Update"
            
            txtID.Locked = True
    End Select
    End Sub


Private Function SaveAdd()
    Dim NewBookS As aBooks
    Dim oldBooks As aBooks
    
    
    'check form field
    If IsEmpty(txtID.Text) Then
        MsgBox "Please enter Books ID", vbExclamation
       HLTxt txtID
        Exit Function
    End If
   If IsEmpty(txtName.Text) Then
        MsgBox "Please Enter the Books Name", vbExclamation
        HLTxt txtName
        Exit Function
    End If
    
    If mFormState = "add" Then
       
    'check duplication
    If GetBooksID(txtID.Text, oldBooks) = True Then
        MsgBox "The BooksID that you have entered is already exist." & vbNewLine & vbNewLine & _
            "Please enter different value.", vbExclamation
        HLTxt txtID
        Exit Function
    End If
    NewBookS.ID = txtID.Text
    NewBookS.Name = txtName.Text
    NewBookS.Publisher = txtPublication.Text
    NewBookS.Subject = txtSubject.Text
    NewBookS.Author = txtAuthor.Text
    NewBookS.Price = txtPrice.Text
   ' NewBookS.NoofBooks = txtNoofBooks.Text
    NewBookS.bDate = DtDate.Value
    'try
    
    If ModRsBooks.AddBookS(NewBookS) = True Then
        MsgBox "New Books entry was successfully created.", vbInformation
        'set flag
        mShowAdd = True
        'close form and return
        Unload Me
        Me.ShowAdd
        
        
    Else
    
        MsgBox "Unable to add new Books entry.", vbExclamation
        'set flag
        mShowAdd = False
        
    End If
    End If
End Function

Private Function SaveEdit()

    Dim NewBookS As aBooks
    Dim oldBooks As aBooks
    
    'check form field
    
    
      If IsEmpty(txtID.Text) Then
        MsgBox "Please enter Books ID", vbExclamation
       HLTxt txtID
        Exit Function
    End If
   If IsEmpty(txtName.Text) Then
        MsgBox "Please Enter the Books Name", vbExclamation
        HLTxt txtName
        Exit Function
    End If
    
    
    'set new
     NewBookS.ID = txtID.Text
    NewBookS.Name = txtName.Text
    NewBookS.Publisher = txtPublication.Text
    NewBookS.Subject = txtSubject.Text
    NewBookS.Author = txtAuthor.Text
    NewBookS.Price = txtPrice.Text
    'NewBookS.NoofBooks = txtNoofBooks.Text
    NewBookS.bDate = DtDate.Value
    'try
    'add new Books
    If ModRsBooks.EditBookS(NewBookS) = True Then
        MsgBox "Books entry was updated.", vbInformation
        'set flag
        mShowEdit = True
        'close form and return
        Unload Me
        
    Else
    
        MsgBox "Unable to update Books entry.", vbExclamation
        'set flag
        mShowAdd = False
        
    End If
    
End Function








