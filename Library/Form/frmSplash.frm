VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   4860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   324
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   576
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      TabIndex        =   6
      Top             =   4320
      Width           =   45
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   4845
      Left            =   0
      Top             =   0
      Width           =   8625
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   4560
      MouseIcon       =   "frmSplash.frx":11549
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   4440
      Width           =   3720
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   0
      MouseIcon       =   "frmSplash.frx":1169B
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   0
      Width           =   3720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COPYRIGHT © Seanet Technologies 2007"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   4920
      TabIndex        =   3
      Top             =   4620
      Width           =   3000
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This product is licensed in to:"
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
      Top             =   4080
      Width           =   2430
   End
   Begin VB.Label lblComputerName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   2640
      TabIndex        =   1
      Top             =   4080
      Width           =   45
   End
   Begin VB.Label lblStat 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading.."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6840
      TabIndex        =   0
      Top             =   3840
      Width           =   780
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'API for accept Computer Name
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'API for Top Most form
Private Declare Function SetWindowPos Lib "user32" _
    (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Private Const HWND_NOTOPMOST = -2

Dim isOn As Boolean



Public Function ShowSplash()
    
    'show form
    SetWindowPos Me.Hwnd, HWND_TOPMOST, _
    0, 0, 0, 0, FLAGS
    Me.Show
    
    DoEvents
    DoEvents
    DoEvents
    
    'continue loading...
    Call ModMain.Main_AfterSD

End Function


Public Function ShowForm()
    
    lblStat.Caption = ""
    'show form
    Me.Show
End Function

Public Function UnloadSplash()
    Unload Me
End Function


Private Sub Form_Activate()
    
    If isOn = True Then
        Exit Sub
    End If
    isOn = True
    
    'backup database
  
   
End Sub

Private Sub Form_Deactivate()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub


Private Sub Form_Load()
Dim computer_name As String
Dim Length As Long

    computer_name = Space$(256)
    Length = Len(computer_name)
    GetComputerName computer_name, Length
    computer_name = Left$(computer_name, Length)
    lblComputerName.Caption = computer_name
Label4.Caption = "Serial number : " & GetSetting(App.Title, "RegKey", "Key")


End Sub

Private Sub Label1_Click()
Shell "explorer.exe http://www.Seanettechnologies.com"

End Sub


