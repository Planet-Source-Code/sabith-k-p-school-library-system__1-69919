Attribute VB_Name = "ModRsBooks"
Public Type aBooks
   ID As String
   Name As String
   Publisher As String
   Subject As String
   Author As String
   Price As String
   NoofBooks As Integer
   Barowed As String
   bDate As String
End Type
Public Function AddBookS(vBooks As aBooks) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String

    
    
    'default
    AddBookS = False
    
    sSQL = "SELECT * FROM tblBooks WHERE ID='" & vBooks.ID & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
       MsgBox Err.Description, vbExclamation
    End If
    
    If AnyRecordExisted(vRS) = True Then
        AddBookS = True
        GoTo RAE
    End If
    
    'add new record
    vRS.AddNew
    
    If WriteBooks(vRS, vBooks) = False Then
        'GoTo RAE
    End If
    
    vRS.Update
   
    
    AddBookS = True
    
RAE:
    Set vRS = Nothing
End Function
Public Function EditBookS(vBooks As aBooks) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    EditBookS = False
    
    sSQL = "SELECT * FROM tblBooks WHERE ID= '" & vBooks.ID & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        MsgBox Err.Description, vbExclamation
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        MsgBox Err.Description, vbExclamation
        GoTo RAE
    End If
    
    'edit
    If WriteBooks(vRS, vBooks) = False Then
        'GoTo RAE
    End If
    
    vRS.Update

    EditBookS = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function DeleteBookS(ByVal iBooksNo As String) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
        
    
    On Error GoTo RAE
    'default
    DeleteAgent = False
    
    sSQL = "DELETE * FROM tblBooks WHERE ID= '" & iBooksNo & "'"

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
     
    DeleteBookS = True
    
RAE:
    Set vRS = Nothing
End Function
Public Function WriteBooks(ByRef vRS As ADODB.Recordset, ByRef vBooks As aBooks) As Boolean
    
    'default
    WriteBooks = False
    
    'On Error GoTo RAE

    With vBooks
       vRS.Fields("ID") = .ID
       vRS.Fields("Name") = .Name
       vRS.Fields("Publisher") = .Publisher
       vRS.Fields("Subject") = .Subject
       vRS.Fields("Author") = .Author
       vRS.Fields("Price") = .Price
       'vRS.Fields("NoofBooks") = .NoofBooks
       vRS.Fields("bDate") = .bDate
       
        
    End With

    WriteBooks = True
    Exit Function
    
'RAE:
 '  MsgBox Err.Address
End Function


Public Function ReadBooks(ByRef vRS As ADODB.Recordset, ByRef vBooks As aBooks) As Boolean
    
    'default
    ReadBooks = False
    
    On Error GoTo RAE
    
    With vBooks
       .ID = ReadField(vRS.Fields("ID"))
       .Name = ReadField(vRS.Fields("Name"))
       .Publisher = ReadField(vRS.Fields("Publisher"))
       .Subject = ReadField(vRS.Fields("Subject"))
       .Author = ReadField(vRS.Fields("Author"))
       .Price = ReadField(vRS.Fields("Price"))
       '.NoofBooks = ReadField(vRS.Fields("NoofBooks"))
       .Barowed = ReadField(vRS.Fields("Barowed"))
       .bDate = ReadField(vRS.Fields("bDate"))
        
    End With
    
    ReadBooks = True
    Exit Function
    
RAE:
    
End Function
Public Function GetBooksID(sBooksID As String, vBooks As aBooks) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetBooksID = False
    
    sSQL = "SELECT * FROM tblBooks WHERE ID='" & sBooksID & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
       MsgBox Err.Description, vbExclamation
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadBooks(vRS, vBooks) = False Then
        GoTo RAE
    End If
    
    GetBooksID = True
    
RAE:
    Set vRS = Nothing
End Function



