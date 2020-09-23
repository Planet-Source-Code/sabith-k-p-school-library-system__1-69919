VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAllBooks 
   BackColor       =   &H00D2DAD3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "All Books"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   10920
   Icon            =   "frmAllBooks.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   10920
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
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
      Left            =   5640
      TabIndex        =   4
      Top             =   4800
      Width           =   1215
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
      Left            =   6960
      TabIndex        =   3
      Top             =   4800
      Width           =   1215
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
      Left            =   8280
      TabIndex        =   2
      Top             =   4800
      Width           =   1215
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
      Left            =   9600
      TabIndex        =   1
      Top             =   4800
      Width           =   1215
   End
   Begin MSComctlLib.ListView listBooks 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   7646
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
      ColHdrIcons     =   "ilRecordIco"
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Book Title"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Publisher"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Subject"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Author"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Price"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ilRecordIco 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAllBooks.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   10800
      Y1              =   4560
      Y2              =   4560
   End
End
Attribute VB_Name = "frmAllBooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mShowForm As Boolean
Public Function ShowForm() As Boolean
    'show form
    Me.Show vbModal
    
    ShowForm = mShowForm
    
End Function
Private Sub LoadEntries()
    On Error GoTo RAE
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim rec_count, X
    Dim l_results As ListItem
        listBooks.ListItems.Clear
    'set SQL Expression
    sSQL = "SELECT * From tblBooks" & _
            " ORDER BY tblBooks.id"
             
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        MsgBox Err.Description, vbExclamation
    End If
    
   With vRS
     .MoveLast
        rec_count = .RecordCount
        
    .MoveFirst
        For X = 1 To rec_count
    
    Set l_results = listBooks.ListItems.Add(X, , !ID, 1, 1)
                    l_results.SubItems(1) = !Name
                    l_results.SubItems(2) = !Publisher
                    l_results.SubItems(3) = !Subject
                    l_results.SubItems(4) = !Author
                    l_results.SubItems(5) = !Price
                    'l_results.SubItems(6) = !NoofBooks
                    'l_results.SubItems(6) = !Barowed
                   
    .MoveNext
        Next
End With
RAE:
    Set vRS = Nothing
    listBooks.Refresh
End Sub



Private Sub cmdDelete_Click()
If listBooks.ListItems.Count > 0 Then
            If MsgBox("Are you sure you want to delete Book '" & listBooks.SelectedItem.Text & "'?", vbQuestion + vbYesNo) = vbYes Then
                If DeleteBookS(listBooks.SelectedItem.Text) = True Then
                    LoadEntries
                Else
                    MsgBox Me.Name & "cmdDelete_Click" & "Faild:" & DeleteBookS(listBooks.SelectedItem.Text) = True
                End If
            End If
    End If
End Sub

Private Sub cmdEdit_Click()
 If listBooks.ListItems.Count > 0 Then
        If frm_aed_Books.ShowEdit(listBooks.SelectedItem.Text) = True Then
            LoadEntries
        End If
    End If
    End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdNew_Click()
If frm_aed_Books.ShowAdd = True Then
    LoadEntries
End If
End Sub




Private Sub Form_Load()
    LoadEntries
End Sub

Private Sub listBooks_DblClick()
    cmdEdit_Click
End Sub






