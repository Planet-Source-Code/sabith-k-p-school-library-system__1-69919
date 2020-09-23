Attribute VB_Name = "ModMain"
Public Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Public CurrentUser As tUser
Public ACCStart As Date
Public ACCEnd As Date



Public Sub Main()

'CrackKey = GetSetting("SKG", "SKGKey", "SCheck")
   
'CrackSKey = GetSetting(App.Title, "RegKey", "Key")
   
   'If CrackKey = "" Then
   '''''''''''''''If CrackKey = "" Then
    '    frmRegistration.Show 'vbModalc
        
  ' ElseIf CrackSKey = "" Then
    
   '     frmRegistration.Show
    
  '  Else
    'use system appearance style
    InitCommonControls
    
  'set Database Path
    If InitDB = False Then
        Exit Sub
    End If
    
    '  View Splash
   frmSplash.ShowSplash
End Sub
Public Sub Main_AfterSD()
    

    'Open Database File
    If OpenDB = False Then
        Exit Sub
    End If
     
    
    'TestUnit
   MDIFrm.ShowForm
End Sub
