Attribute VB_Name = "modRsIssue"
Public Type aIssue
    ID As Long
   BookID As String
   MemberID As String
   IssueDate As Date
   Returndate As Date
End Type

Public Function AddIssue(vIssue As aIssue) As Boolean
    
    Dim vRs As New ADODB.Recordset
    Dim sSQL As String

    
    
    'default
    AddIssue = False
    
    sSQL = "SELECT * FROM tbltrans WHERE ID=" & vIssue.ID
    
    If ConnectRS(PrimeDB, vRs, sSQL) = False Then
       MsgBox Err.Description, vbExclamation
    End If
    
    If AnyRecordExisted(vRs) = True Then
        AddIssue = True
        GoTo RAE
    End If
    
    'add new record
    vRs.AddNew
    
    If WriteIssue(vRs, vIssue) = False Then
        'GoTo RAE
    End If
    
    vRs.Update
   
    
    AddIssue = True
    
RAE:
    Set vRs = Nothing
End Function
Public Function EditIssue(vIssue As aIssue) As Boolean
    
    Dim vRs As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    EditIssue = False
    
    sSQL = "SELECT * FROM tbltrans WHERE ID=" & vIssue.ID
    
    If ConnectRS(PrimeDB, vRs, sSQL) = False Then
        MsgBox Err.Description, vbExclamation
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRs) = False Then
        MsgBox Err.Description, vbExclamation
        GoTo RAE
    End If
    
    'edit
    If WriteIssue(vRs, vIssue) = False Then
        'GoTo RAE
    End If
    
    vRs.Update

    EditIssue = True
    
RAE:
    Set vRs = Nothing
End Function

Public Function DeleteIssue(ByVal iIssueID As Long) As Boolean
    
    Dim vRs As New ADODB.Recordset
    Dim sSQL As String
        
    
    On Error GoTo RAE
    'default
    DeleteIssue = False
    
    sSQL = "DELETE * FROM tbltrans WHERE ID=" & iIssueID

    Dim sErrD As String
    Dim iErrN As Long
    If ConnectRS(PrimeDB, vRs, sSQL, False, iErrN, sErrD) = False Then
        If iErrN = -2147467259 Then
            'it includes releted data
            MsgBox "Unable to delete entry. It includes other related record." & vbNewLine & vbNewLine & _
                    "Details: " & sErrD, vbExclamation
        Else
            'WriteErrorLog "modRAgent", "DeleteAgent", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
            'GoTo RAE
        End If
    End If
     
    DeleteIssue = True
    
RAE:
    Set vRs = Nothing
End Function
Public Function WriteIssue(ByRef vRs As ADODB.Recordset, ByRef vIssue As aIssue) As Boolean
    
    'default
    WriteIssue = False
    
    'On Error GoTo RAE

    With vIssue
       'vRs.Fields("ID") = .ID
       vRs.Fields("BookID") = .BookID
       vRs.Fields("MemberID") = .MemberID
       vRs.Fields("IDate") = .IssueDate
       vRs.Fields("RDate") = .Returndate
      
        
    End With

    WriteIssue = True
    Exit Function
    
'RAE:
 '  MsgBox Err.Address
End Function


Public Function ReadIssue(ByRef vRs As ADODB.Recordset, ByRef vIssue As aIssue) As Boolean
    
    'default
    ReadIssue = False
    
    On Error GoTo RAE
    
    With vIssue
       '.ID = ReadField(vRs.Fields("ID"))
       .BookID = ReadField(vRs.Fields("BookID"))
        .MemberID = ReadField(vRs.Fields("MemberID"))
        .BookID = ReadField(vRs.Fields("BookID"))
        .IssueDate = ReadField(vRs.Fields("IDate"))
        .Returndate = ReadField(vRs.Fields("RDate"))
        
    End With
    
    ReadIssue = True
    Exit Function
    
RAE:
    
End Function
Public Function GetIssueID(sIssueID As Long, vIssue As aIssue) As Boolean
    
    Dim vRs As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetIssueID = False
    
    sSQL = "SELECT * FROM tblTrans WHERE ID=" & sIssueID
    
    If ConnectRS(PrimeDB, vRs, sSQL) = False Then
       MsgBox Err.Description, vbExclamation
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRs) = False Then
        GoTo RAE
    End If
    
    vRs.MoveFirst
    
    If ReadIssue(vRs, vIssue) = False Then
        GoTo RAE
    End If
    
    GetIssueID = True
    
RAE:
    Set vRs = Nothing
End Function
 
