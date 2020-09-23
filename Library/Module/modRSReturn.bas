Attribute VB_Name = "modRSReturn"
Public Type aReturn
    ID As Long
    MemberID As String
    BookID As String
    Returndate As Date
    BReturn As Boolean
End Type

Public Function InsertReturn(vReturn As aReturn) As Boolean
    Dim vRs As New ADODB.Recordset
    Dim sSQL As String

'default
InsertReturn = False

    sSQL = "Select * from tbltrans where MemberID='" & vReturn.MemberID & "' and BookID='" & vReturn.BookID & "'" & "and BReturn=False "
    
    If ConnectRS(PrimeDB, vRs, sSQL) = False Then
        MsgBox Err.Description, vbExclamation
    End If
    
    If AnyRecordExisted(vRs) = False Then
        MsgBox Err.Description, vbExclamation
    End If
    
    If WriteReturn(vRs, vReturn) = False Then
        MsgBox Err.Description, vbExclamation
    End If
    
    vRs.Update
    
    InsertReturn = True

End Function
Public Function WriteReturn(ByRef vRs As ADODB.Recordset, ByRef vReturn As aReturn) As Boolean
    
    WriteReturn = False
    
    With vReturn
        'vRs.Fields("ID") = .ID
        vRs.Fields("MemberID") = .MemberID
        vRs.Fields("BookID") = .BookID
        vRs.Fields("BReturn") = .BReturn
        vRs.Fields("ReturnDate") = .Returndate
    End With
    
WriteReturn = True

End Function
Public Function ReadReturn(ByRef vRs As ADODB.Recordset, ByRef vReturn As aReturn) As Boolean
 
 ReadReturn = False
  
  With vReturn
    '.ID = ReadField(vRs.Fields("ID"))
    .BookID = ReadField(vRs.Fields("BookID"))
    .MemberID = ReadField(vRs.Fields("MemberID"))
    .BReturn = ReadField(vRs.Fields("BReturn"))
    .Returndate = ReadField(vRs.Fields("ReturnDate"))
End With

ReadReturn = True
    
 
End Function
Public Function GetReturnInform(MemberID As String, BookID As String, BReturn As Boolean, vReturn As aReturn) As Boolean
            Dim vRs As New ADODB.Recordset
            Dim sSQL As String
        
    GetReturnInform = False
    
    sSQL = "Select * from tblTrans Where MemberID='" & MemberID & "' and BookID= '" & BookID & "' and BReturn=" & False
    
    If ConnectRS(PrimeDB, vRs, sSQL) = False Then
        MsgBox Err.Description, vbExclamation
    End If
    
    If AnyRecordExisted(vRs) = False Then
        GoTo RAE
    End If
    
    vRs.MoveFirst
    
    If ReadReturn(vRs, vReturn) = False Then
        MsgBox Err.Description, vbExclamation
    End If
    
    GetReturnInform = True
     Exit Function
RAE:
Set vRs = Nothing
GetReturnInform = False
        
End Function

