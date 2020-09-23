Attribute VB_Name = "ModRsMember"
Public Type aMember
   ID As String
   Name As String
   Age As Integer
   Class As String
   Division As String
   Address As String
   BookStatus As Integer
   mDate As Date
End Type
Public Function AddMember(vMember As aMember) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String

    
    
    'default
    AddMember = False
    
    sSQL = "SELECT * FROM tblMember WHERE ID='" & vMember.ID & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
       MsgBox Err.Description, vbExclamation
    End If
    
    If AnyRecordExisted(vRS) = True Then
        AddMember = True
        GoTo RAE
    End If
    
    'add new record
    vRS.AddNew
    
    If WriteMember(vRS, vMember) = False Then
        'GoTo RAE
    End If
    
    vRS.Update
   
    
    AddMember = True
    
RAE:
    Set vRS = Nothing
End Function
Public Function EditMember(vMember As aMember) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    EditMember = False
    
    sSQL = "SELECT * FROM tblMember WHERE ID= '" & vMember.ID & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        MsgBox Err.Description, vbExclamation
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        MsgBox Err.Description, vbExclamation
        GoTo RAE
    End If
    
    'edit
    If WriteMember(vRS, vMember) = False Then
        'GoTo RAE
    End If
    
    vRS.Update

    EditMember = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function DeleteMember(ByVal iMemberNo As String) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
        
    
    On Error GoTo RAE
    'default
    DeleteAgent = False
    
    sSQL = "DELETE * FROM tblMember WHERE ID= '" & iMemberNo & "'"

    Dim sErrD As String
    Dim iErrN As Long
    If ConnectRS(PrimeDB, vRS, sSQL, False, iErrN, sErrD) = False Then
        If iErrN = -2147467259 Then
            'it includes releted data
            MsgBox "Unable to delete entry. It includes other related record." & vbNewLine & vbNewLine & _
                    "Details: " & sErrD, vbExclamation
        Else
            'WriteErrorLog "modRAgent", "DeleteAgent", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
            'GoTo RAE
        End If
    End If
     
    DeleteMember = True
    
RAE:
    Set vRS = Nothing
End Function
Public Function WriteMember(ByRef vRS As ADODB.Recordset, ByRef vMember As aMember) As Boolean
    
    'default
    WriteMember = False
    
    'On Error GoTo RAE

    With vMember
       vRS.Fields("ID") = .ID
       vRS.Fields("Name") = .Name
       vRS.Fields("Age") = .Age
       vRS.Fields("Class") = .Class
       vRS.Fields("Division") = .Division
       vRS.Fields("Address") = .Address
       vRS.Fields("BookStatus") = .BookStatus
       vRS.Fields("mDate") = .mDate
       
        
    End With

    WriteMember = True
    Exit Function
    
'RAE:
 '  MsgBox Err.Address
End Function


Public Function ReadMember(ByRef vRS As ADODB.Recordset, ByRef vMember As aMember) As Boolean
    
    'default
    ReadMember = False
    
    On Error GoTo RAE
    
    With vMember
        .ID = ReadField(vRS.Fields("ID"))
        .Name = ReadField(vRS.Fields("name"))
        .Age = ReadField(vRS.Fields("Age"))
        .Class = ReadField(vRS.Fields("Class"))
        .Division = ReadField(vRS.Fields("Division"))
        .Address = ReadField(vRS.Fields("Address"))
        .BookStatus = ReadField(vRS.Fields("BookStatus"))
        .mDate = ReadField(vRS.Fields("mDate"))
        
        
    End With
    
    ReadMember = True
    Exit Function
    
RAE:
    
End Function
Public Function GetMemberID(sMemberID As String, vMember As aMember) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetMemberID = False
    
    sSQL = "SELECT * FROM tblMember WHERE ID='" & sMemberID & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
       MsgBox Err.Description, vbExclamation
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadMember(vRS, vMember) = False Then
        GoTo RAE
    End If
    
    GetMemberID = True
    
RAE:
    Set vRS = Nothing
End Function


