Attribute VB_Name = "getRegistro"

Public Function GetUltimoR(Hoja As Worksheet) As Integer
    GetUltimoR = GetNuevoR(Hoja) - 1
End Function

Public Function GetNuevoR(Hoja As Worksheet) As Integer
    
    Dim Fila As Long
    Fila = 2
    
    Do While Hoja.Cells(Fila, 1) <> ""
        Fila = Fila + 1
    Loop
    
    GetNuevoR = Fila
    
End Function

Public Sub Agregar(cmbBox As ComboBox, sItem As String)

'Agrega los item únicos y en orden alfabético

For i = 0 To cmbBox.ListCount - 1
    Select Case StrComp(cmbBox.List(i), sItem, vbTextCompare)
        Case 0: Exit Sub 'ya existe en el combo y ya no lo agrega
        Case 1: cmbBox.AddItem sItem, i: Exit Sub 'Es menor, lo agrega antes del comparado
    End Select
Next

cmbBox.AddItem sItem 'Es mayor lo agrega al final

End Sub


Public Sub CopiarClientes()
    Dim Conn As ADODB.Connection
    Dim MiConexion
    Dim Rs As ADODB.Recordset
    Dim MiBase As String
    Dim Query As String
    
    MiBase = "cotizador.accdb"

    Set Conn = New ADODB.Connection
    MiConexion = Application.ThisWorkbook.Path & Application.PathSeparator & MiBase

    With Conn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open MiConexion
    End With
    
    'traer datos del cliente para verificar
    Query = "SELECT * FROM clientes"

    Set Rs = New ADODB.Recordset
    Rs.CursorLocation = adUseServer
    Rs.Open Source:=Query, _
    ActiveConnection:=Conn
    
       
    Sheets("clientes").Range("A2").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.ClearContents
    
    For i = 0 To Rs.Fields.Count - 1
    
        Cells(1, i + 1).Value = Rs.Fields(i).Name
    
    Next i
    
    Sheets("clientes").Range("A2").CopyFromRecordset Rs
    
    Rs.Close
   
    Conn.Close
    Set Rs = Nothing
    Set Conn = Nothing

End Sub


Public Sub CopiarContactoCliente()
    Dim Conn As ADODB.Connection
    Dim MiConexion
    Dim Rs As ADODB.Recordset
    Dim MiBase As String
    Dim Query As String
    
    MiBase = "cotizador.accdb"

    Set Conn = New ADODB.Connection
    MiConexion = Application.ThisWorkbook.Path & Application.PathSeparator & MiBase

    With Conn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open MiConexion
    End With
    
    'traer datos del contacto_cliente para verificar
    Query = "SELECT * FROM contacto_cliente"

    Set Rs = New ADODB.Recordset
    Rs.CursorLocation = adUseServer
    Rs.Open Source:=Query, _
    ActiveConnection:=Conn
    
       
    Sheets("contacto_cliente").Range("A2").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.ClearContents
    
    For i = 0 To Rs.Fields.Count - 1
    
        Cells(1, i + 1).Value = Rs.Fields(i).Name
    
    Next i
    
    Sheets("contacto_cliente").Range("A2").CopyFromRecordset Rs
    
    Rs.Close
   
    Conn.Close
    Set Rs = Nothing
    Set Conn = Nothing

End Sub

Public Sub CopiarProveedores()
    Dim Conn As ADODB.Connection
    Dim MiConexion
    Dim Rs As ADODB.Recordset
    Dim MiBase As String
    Dim Query As String
    
    MiBase = "cotizador.accdb"

    Set Conn = New ADODB.Connection
    MiConexion = Application.ThisWorkbook.Path & Application.PathSeparator & MiBase

    With Conn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open MiConexion
    End With
    
    'traer datos del proveedor para verificar
    Query = "SELECT * FROM proveedores"

    Set Rs = New ADODB.Recordset
    Rs.CursorLocation = adUseServer
    Rs.Open Source:=Query, _
    ActiveConnection:=Conn
    
       
    Sheets("proveedores").Range("A2").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.ClearContents
    
    For i = 0 To Rs.Fields.Count - 1
    
        Cells(1, i + 1).Value = Rs.Fields(i).Name
    
    Next i
    
    Sheets("proveedores").Range("A2").CopyFromRecordset Rs
    
    Rs.Close
   
    Conn.Close
    Set Rs = Nothing
    Set Conn = Nothing

End Sub

Public Sub CopiarContactoProveedor()
    Dim Conn As ADODB.Connection
    Dim MiConexion
    Dim Rs As ADODB.Recordset
    Dim MiBase As String
    Dim Query As String
    
    MiBase = "cotizador.accdb"

    Set Conn = New ADODB.Connection
    MiConexion = Application.ThisWorkbook.Path & Application.PathSeparator & MiBase

    With Conn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open MiConexion
    End With
    
    'traer datos del contacto_proveedor para verificar
    Query = "SELECT * FROM contacto_proveedor"

    Set Rs = New ADODB.Recordset
    Rs.CursorLocation = adUseServer
    Rs.Open Source:=Query, _
    ActiveConnection:=Conn
    
       
    Sheets("contacto_proveedor").Range("A2").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.ClearContents
    
    For i = 0 To Rs.Fields.Count - 1
    
        Cells(1, i + 1).Value = Rs.Fields(i).Name
    
    Next i
    
    Sheets("contacto_proveedor").Range("A2").CopyFromRecordset Rs
    
    Rs.Close
   
    Conn.Close
    Set Rs = Nothing
    Set Conn = Nothing

End Sub


Public Sub CopiarProductos()
    Dim Conn As ADODB.Connection
    Dim MiConexion
    Dim Rs As ADODB.Recordset
    Dim MiBase As String
    Dim Query As String
    
    MiBase = "cotizador.accdb"

    Set Conn = New ADODB.Connection
    MiConexion = Application.ThisWorkbook.Path & Application.PathSeparator & MiBase

    With Conn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open MiConexion
    End With
    
    'traer datos del producto para verificar
    Query = "SELECT * FROM productos"

    Set Rs = New ADODB.Recordset
    Rs.CursorLocation = adUseServer
    Rs.Open Source:=Query, _
    ActiveConnection:=Conn
    
       
    Sheets("productos").Range("A2").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.ClearContents
    
    For i = 0 To Rs.Fields.Count - 1
    
        Cells(1, i + 1).Value = Rs.Fields(i).Name
    
    Next i
    
    Sheets("productos").Range("A2").CopyFromRecordset Rs
    
    Rs.Close
   
    Conn.Close
    Set Rs = Nothing
    Set Conn = Nothing

End Sub

Public Sub CopiarEmpleados()
    Dim Conn As ADODB.Connection
    Dim MiConexion
    Dim Rs As ADODB.Recordset
    Dim MiBase As String
    Dim Query As String
    
    MiBase = "cotizador.accdb"

    Set Conn = New ADODB.Connection
    MiConexion = Application.ThisWorkbook.Path & Application.PathSeparator & MiBase

    With Conn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open MiConexion
    End With
    
    'traer datos del cliente para verificar
    Query = "SELECT * FROM empleados"

    Set Rs = New ADODB.Recordset
    Rs.CursorLocation = adUseServer
    Rs.Open Source:=Query, _
    ActiveConnection:=Conn
    
       
    Sheets("empleados").Range("A2").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.ClearContents
    
    For i = 0 To Rs.Fields.Count - 1
    
        Cells(1, i + 1).Value = Rs.Fields(i).Name
    
    Next i
    
    Sheets("empleados").Range("A2").CopyFromRecordset Rs
    
    Rs.Close
   
    Conn.Close
    Set Rs = Nothing
    Set Conn = Nothing

End Sub

Public Sub CopiarTransportadores()
    Dim Conn As ADODB.Connection
    Dim MiConexion
    Dim Rs As ADODB.Recordset
    Dim MiBase As String
    Dim Query As String
    
    MiBase = "cotizador.accdb"

    Set Conn = New ADODB.Connection
    MiConexion = Application.ThisWorkbook.Path & Application.PathSeparator & MiBase

    With Conn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open MiConexion
    End With
    
    'traer datos del trsnportador para verificar
    Query = "SELECT * FROM transportadores"

    Set Rs = New ADODB.Recordset
    Rs.CursorLocation = adUseServer
    Rs.Open Source:=Query, _
    ActiveConnection:=Conn
    
       
    Sheets("transportadores").Range("A2").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.ClearContents
    
    For i = 0 To Rs.Fields.Count - 1
    
        Cells(1, i + 1).Value = Rs.Fields(i).Name
    
    Next i
    
    Sheets("transportadores").Range("A2").CopyFromRecordset Rs
    
    Rs.Close
   
    Conn.Close
    Set Rs = Nothing
    Set Conn = Nothing

End Sub

Public Sub CopiarColores()
    Dim Conn As ADODB.Connection
    Dim MiConexion
    Dim Rs As ADODB.Recordset
    Dim MiBase As String
    Dim Query As String
    
    MiBase = "cotizador.accdb"

    Set Conn = New ADODB.Connection
    MiConexion = Application.ThisWorkbook.Path & Application.PathSeparator & MiBase

    With Conn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open MiConexion
    End With
    
    'traer datos del trsnportador para verificar
    Query = "SELECT * FROM colores"

    Set Rs = New ADODB.Recordset
    Rs.CursorLocation = adUseServer
    Rs.Open Source:=Query, _
    ActiveConnection:=Conn
    
       
    Sheets("colores").Range("A2").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.ClearContents
    
    For i = 0 To Rs.Fields.Count - 1
    
        Cells(1, i + 1).Value = Rs.Fields(i).Name
    
    Next i
    
    Sheets("colores").Range("A2").CopyFromRecordset Rs
    
    Rs.Close
   
    Conn.Close
    Set Rs = Nothing
    Set Conn = Nothing

End Sub

Public Sub CopiarMedidas()
    Dim Conn As ADODB.Connection
    Dim MiConexion
    Dim Rs As ADODB.Recordset
    Dim MiBase As String
    Dim Query As String
    
    MiBase = "cotizador.accdb"

    Set Conn = New ADODB.Connection
    MiConexion = Application.ThisWorkbook.Path & Application.PathSeparator & MiBase

    With Conn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open MiConexion
    End With
    
    'traer datos del trsnportador para verificar
    Query = "SELECT * FROM medidas"

    Set Rs = New ADODB.Recordset
    Rs.CursorLocation = adUseServer
    Rs.Open Source:=Query, _
    ActiveConnection:=Conn
    
       
    Sheets("medidas").Range("A2").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.ClearContents
    
    For i = 0 To Rs.Fields.Count - 1
    
        Cells(1, i + 1).Value = Rs.Fields(i).Name
    
    Next i
    
    Sheets("medidas").Range("A2").CopyFromRecordset Rs
    
    Rs.Close
   
    Conn.Close
    Set Rs = Nothing
    Set Conn = Nothing

End Sub


