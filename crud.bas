
'Established Connection
Dim objCon As ADODB.Connection
Sub Connect()
    Set objCon = New ADODB.Connection
    objCon.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\622052\Database1.accdb;Persist Security Info=False;"
    objCon.Open
    Debug.Print "Connection established..."
End Sub


'Closed Connection
Sub CloseConnection()
    On Error Resume Next
    objCon.Close
    Debug.Print "Connection closed..."
    Set objCon = Nothing
    On Error GoTo 0
End Sub

'Insert Record
Sub Insert_Record()
    strSQL = "INSERT INTO Employee(First_Name, Last_Name) Values('John', 'Kamei')"
    objCon.Execute strSQL
End Sub

'Update Record
Sub Update_Record()
    strSQL = "UPDATE Employee SET First_Name= 'Arvindchand' WHERE ID=4"
    objCon.Execute strSQL
End Sub

'Delete Record
Sub Delete_Data()
    strSQL = "DELETE FROM Employee WHERE ID=5"
    objCon.Execute strSQL
End Sub

'Read Records
Sub Read_Data()
    strSQL = "SELECT * FROM Employee"
    Set objRecordSet = New ADODB.Recordset
    objRecordSet.Open strSQL, objCon
    'Print Fields | Columns
    ThisWorkbook.Sheets(1).Range("A1").Select
    For Each objField In objRecordSet.Fields
        ActiveCell.Value = objField.Name
        ActiveCell.Offset(0, 1).Select
    Next
    Set oRange = ThisWorkbook.Sheets(1).Range("A2")
    oRange.CopyFromRecordset objRecordSet
    Range("A1").CurrentRegion.EntireColumn.AutoFit
End Sub

'Clear Sheet
Sub Clear_Sheet1()
    Dim oRange As Range
    Set oRange = ThisWorkbook.Sheets(1).UsedRange
    oRange.Clear
End Sub

'Main Subroutine To Perform CRUD
Sub Perform_CRUD()
    Call Clear_Sheet1
    Call Connect
    Call Insert_Record
    Call Update_Record
    Call Read_Data
    Call Delete_Data
    Call Read_Data
    Call CloseConnection
End Sub