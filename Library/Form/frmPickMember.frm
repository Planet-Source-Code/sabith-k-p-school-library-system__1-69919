VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmPickMember 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFind 
      Height          =   315
      Left            =   480
      TabIndex        =   1
      Top             =   60
      Width           =   3180
   End
   Begin MSComctlLib.ImageList ilRecordIco 
      Left            =   4140
      Top             =   2175
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
            Picture         =   "frmPickMember.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView listRecord 
      Height          =   3240
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   5715
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      Icons           =   "ilRecordIco"
      SmallIcons      =   "ilRecordIco"
      ForeColor       =   8399906
      BackColor       =   16777215
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmPickMember.frx":059A
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   5821
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3795
      TabIndex        =   2
      Top             =   15
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   661
      Caption         =   "&Cancel"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   0
      cBhover         =   16185592
      cGradient       =   16185592
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      cBack           =   14215660
   End
   Begin lvButton.lvButtons_H cmdSelect 
      Height          =   375
      Left            =   4665
      TabIndex        =   3
      Top             =   15
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   661
      Caption         =   "&Select"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   0
      cBhover         =   16185592
      cGradient       =   16185592
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      cBack           =   14215660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Find"
      ForeColor       =   &H0030A0B8&
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   300
   End
   Begin VB.Image Image1 
      Height          =   435
      Left            =   0
      Picture         =   "frmPickMember.frx":0E74
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5505
   End
End
Attribute VB_Name = "frmPickMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Dim R As RECT
Dim Alignable As Boolean


Dim tmpDepartment As String
Dim vRS As New ADODB.Recordset

Dim MaxEntryCount As Long
Dim CurRecPos As Long
Private Sub Form_Activate()
Dim NewLeft As Long
    Dim NewTop As Long
    
    If Alignable = True Then
        If (R.Left * Screen.TwipsPerPixelX + Me.Width) > Screen.Width Then
            NewLeft = (R.Right * Screen.TwipsPerPixelX) - Me.Width
        Else
            NewLeft = R.Left * Screen.TwipsPerPixelX
        End If
        
        If (R.Bottom * Screen.TwipsPerPixelY + Me.Height) > Screen.Height Then
            NewTop = (R.Top * Screen.TwipsPerPixelY) - Me.Height
            If NewTop < 0 Then NewTop = 0
        Else
            NewTop = R.Bottom * Screen.TwipsPerPixelY
        End If
        
        Me.Left = NewLeft
        Me.Top = NewTop
        
    Else
    
        CenterForm Me
        
    End If
End Sub

