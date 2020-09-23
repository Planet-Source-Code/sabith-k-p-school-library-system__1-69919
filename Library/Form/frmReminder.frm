VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReminder 
   BackColor       =   &H00D2DAD3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UnReturned Book Reminder"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10740
   Icon            =   "frmReminder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   10740
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   120
   End
   Begin MSComctlLib.ListView list 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "MemberID"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "BookID"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "IssuDate"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Return Date"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   10680
      Y1              =   4440
      Y2              =   4440
   End
End
Attribute VB_Name = "frmReminder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LoadEntries()
    On Error GoTo RAE
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim rec_count, X
    Dim l_results As ListItem
        list.ListItems.Clear
    'set SQL Expression
    sSQL = "SELECT * From tblTrans WHERE BReturn=False And rDate <" & CLng(Date) _

    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        MsgBox Err.Description, vbExclamation
    End If
    
   With vRS
    If .RecordCount = 0 Then
        Exit Sub
    End If
     .MoveLast
        rec_count = .RecordCount
        
        
    .MoveFirst
        For X = 1 To rec_count
    
    Set l_results = list.ListItems.Add(X, , !ID)
                    l_results.SubItems(1) = !MemberID
                    l_results.SubItems(2) = !BookID
                    l_results.SubItems(3) = !IDate
                    l_results.SubItems(4) = !RDate
                    'l_results.SubItems(5) = !Address
                   
                   
    .MoveNext
        Next
End With
RAE:
    Set vRS = Nothing
    list.Refresh
End Sub

Public Sub ShowForm()
    LoadEntries
    Me.Show vbModal
End Sub

Private Sub Timer1_Timer()
If list.ListItems.Count = 0 Then
    Unload Me
End If
End Sub
