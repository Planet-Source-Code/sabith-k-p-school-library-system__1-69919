Attribute VB_Name = "modFunction"

    
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public BImg() As Byte
Public strImgN As String
Public findmode As Boolean
Public infomode As Boolean
Public vRS As New ADODB.Recordset
     


Public Function AnyRecordExisted(ByRef vRS As ADODB.Recordset) As Boolean
    If vRS.State = adStateClosed Then
        AnyRecordExisted = False
        Exit Function
    End If
    
    
    vRS.Requery
    
    If (vRS.BOF = True) And (vRS.EOF = True) Then
        AnyRecordExisted = False
    Else
        On Error GoTo errh
        vRS.MoveFirst
        AnyRecordExisted = True
    End If

    Exit Function
    '--------------------------
    
errh:
    AnyRecordExisted = False
End Function


Public Function ReadField(ByRef vField As Field) As Variant
    
    On Error GoTo errh

    If Not IsNull(vField.Value) Then
        ReadField = vField.Value
    Else
        Select Case vField.Type
            Case adBigInt
                ReadField = 0
            Case adBinary
                ReadField = 0
            Case adBoolean
                ReadField = False
            Case adByRef 'temp
                ReadField = 0
            Case adBSTR
                ReadField = ""
            Case adChar
                ReadField = ""
            Case adCurrency
                ReadField = 0
            Case adDate
                ReadField = CDate(0)
            Case adDBDate
                ReadField = CDate(0)
            Case adDBTime
                ReadField = FormatDateTime(CDate(0), vbLongTime)
            Case adDBTimeStamp
                ReadField = CDate(0)
            Case adDecimal
                ReadField = 0
            Case adDouble
                ReadField = 0
            Case adEmpty 'temp
                ReadField = ""
            Case adError
                ReadField = 0
            Case adNumeric
                ReadField = 0
            Case adDouble
                ReadField = 0
            Case Else
                ReadField = ""
            End Select
    End If
    
    Exit Function
    
errh:
    ReadField = ""
End Function
Public Function IsEmpty(s As String) As Boolean
    If Len(Trim(s)) < 1 Then
        IsEmpty = True
    Else
        IsEmpty = False
    End If
End Function
Public Function HLTxt(ByRef txt As Object)
On Error Resume Next
    txt.SelStart = 0
    txt.SelLength = Len(txt)
    txt.SetFocus
End Function
Public Function ShowReport(tbl As String, Datarpt As DataReport, Optional Datacon As String = "")
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    sSQL = "SELECT * FROM " & tbl & Datacon
        If ConnectRS(PrimeDB, vRS, sSQL) = False Then
            MsgBox "Report" & "Unable to Connect Recordset.SQL expression: '" & sSQL & "'"
        End If
        Set Datarpt.DataSource = vRS
End Function

Public Sub OpenMySite()
ShellExecute 0&, "Open", "http://hackw.blogspot.com", 0&, 0&, vbNormalFocus
End Sub
