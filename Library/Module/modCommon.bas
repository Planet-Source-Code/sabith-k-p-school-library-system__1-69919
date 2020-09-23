Attribute VB_Name = "modCommon"

Public Function LoadItemEntries(frm As Form, cmbItemname As ComboBox)
    On Error GoTo RAE
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim rec_count, X
    
    'set SQL Expression
    sSQL = "SELECT * From tblItem" & _
            " ORDER BY tblItem.Itemname"
             
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog Me.Name, "LoadEntries", "Unable to connect Recordset. SQL Expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
   With vRS
     .MoveLast
            rec_count = .RecordCount
    .MoveFirst
            For X = 1 To rec_count
                cmbItemname.AddItem !ItemName
    .MoveNext
            Next
End With

RAE:
    Set vRS = Nothing
      'listStudent.Refresh

End Function

