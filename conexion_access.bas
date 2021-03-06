option explicit
'grabar
Sub AltaRegistrosAccess()

    Dim Conn As ADODB.Connection
    Dim MiConexion
    Dim Rs As ADODB.Recordset
    Dim MiBase As String

    MiBase = "MiBase.accdb"

    Set Conn = New ADODB.Connection
    MiConexion = Application.ThisWorkbook.Path & Application.PathSeparator & MiBase

    With Conn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open MiConexion
    End With

    Set Rs = New ADODB.Recordset
    Rs.CursorLocation = adUseServer
    Rs.Open Source:="MiTabla", _
        ActiveConnection:=Conn, _
        CursorType:=adOpenDynamic, _
        LockType:=adLockOptimistic, _
        Options:=adCmdTable

    'Cargar los datos a Tabla de Access
    With Rs
        .AddNew
        .Fields("Fecha") = Date
        .Fields("Nombre") = UserForm1.TextBox1.Value
        .Fields("Ventas") = UserForm1.TextBox2.Value
        .Fields("Comentarios") = UserForm1.TextBox3.Value
    End With

    Rs.Update

    'Cerrar la conexión
    Rs.Close
    Conn.Close
    Set Rs = Nothing
    Set Conn = Nothing

    MsgBox "Alta exitosa", vbInformation, "EXCELeINFO"
    Unload UserForm1
End Sub

'consultar
Private Sub CommandButton1_Click()
    Dim Conn As ADODB.Connection
    Dim MiConexion
    Dim Rs As ADODB.Recordset
    Dim MiBase As String
    Dim Query As String
    Dim i, j

    MiBase = "MiBase.accdb"

    Set Conn = New ADODB.Connection
    MiConexion = Application.ThisWorkbook.Path & Application.PathSeparator & MiBase

    With Conn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open MiConexion
    End With

    Query = "SELECT * FROM MiTabla WHERE nombre LIKE '%" & Me.TextBox1.Value & "%'"

    Set Rs = New ADODB.Recordset
    Rs.CursorLocation = adUseServer
    Rs.Open Source:=Query, _
    ActiveConnection:=Conn

    'Valir si la consulta devuelve resultados
    If Rs.EOF And Rs.BOF Then
        'Borrar la conexión al Recordset
        Rs.Close
        Conn.Close
        'Borrar la memoria
        Set Rs = Nothing
        Set Conn = Nothing
        
        MsgBox "No hay resultados para la consulta", vbInformation, "EXCELeINFO"
        Me.ListBox1.Clear
        Exit Sub
    End If

    'Asignar número de columnas
    With Me.ListBox1
        .ColumnCount = Rs.Fields.Count
    End With

    'Recorrer el Recordset
    Rs.MoveFirst
    i = 1

    With Me.ListBox1
        .Clear
        
        'Asignar los encabezados
        .AddItem
        
        For j = 0 To 4
            .List(0, j) = Rs.Fields(j).Name
        Next j
        
        Do
            .AddItem
            .List(i, 0) = Rs![ID]
            .List(i, 1) = Rs![Fecha]
            .List(i, 2) = Rs![Nombre]
            .List(i, 3) = Rs![Ventas]
            .List(i, 4) = Rs![Comentarios]
            i = i + 1
            Rs.MoveNext
            
        Loop Until Rs.EOF
    End With

    'Cerrar la conexión
    Rs.Close
    Conn.Close
    Set Rs = Nothing
    Set Conn = Nothing

End Sub

'cerrar formulario
Private Sub CommandButton2_Click()
    Unload Me
End Sub

'eliminar
Private Sub CommandButton3_Click()
    Dim Conn As ADODB.Connection
    Dim MiConexion
    Dim Rs As ADODB.Recordset
    Dim MiBase As String
    Dim Query As String
    Dim i, j
    Dim Cuenta As Integer
    Dim Numero As Integer
    Dim ValorElegido As Integer

    'Recorrer el listbox y detectar el item elegido
    '''''''''''''''''''''''''''''''
    Cuenta = Me.ListBox1.ListCount

    'Validamos que haya un elemento seleccionado
    For i = 0 To Cuenta - 1
        If Me.ListBox1.Selected(i) = True Then
            Numero = Numero + 1
        End If
    Next i

    If Numero = 0 Then MsgBox "Debes elegir un elemento", vbExclamation, "EXCELeINFO": Exit Sub

    For j = 0 To Cuenta - 1
        If Me.ListBox1.Selected(j) = True Then
            If Me.ListBox1.ListIndex = 0 Then MsgBox "Encabezado!", vbCritical, "EXCELeINFO": Exit Sub
        ValorElegido = Me.ListBox1.List(j)
        'MsgBox ValorElegido
        End If
    Next j

    '''''''''''''''''''''''''''''''

    MiBase = "MiBase.accdb"

    Set Conn = New ADODB.Connection
    MiConexion = Application.ThisWorkbook.Path & Application.PathSeparator & MiBase

    With Conn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open MiConexion
    End With

    Query = "DELETE * FROM MiTabla WHERE Id = " & ValorElegido

    Set Rs = New ADODB.Recordset
    Rs.CursorLocation = adUseServer
    Rs.Open Source:=Query, _
    ActiveConnection:=Conn

    'Cerrar la conexión
    'Rs.Close
    Conn.Close
    Set Rs = Nothing
    Set Conn = Nothing

    Call CommandButton1_Click

End Sub


'Exportando los datos con DAO
Sub exportar_datos()
    Dim bd As Database, rs As Recordset, r As Long, x As Long, td As TableDef

    'abriendo la base de datos
    Set bd = OpenDatabase("C:\mi_bd.mdb")

    'abriendo recordset
    Set rs = bd.OpenRecordset("MiTabla", dbOpenTable)

    'recogiendo todos los campos en una tabla
    r = 2 ‘empiezo en la fila 2 de la hoja 1

    Do While Len(Range("A" & r).Formula) > 0
    'repetir hasta la primera celda vacía de  la columna A

    With rs
        .AddNew
        .Fields("Alias") = Range("B" & r).Value
        .Fields("apellidos") = Range("C" & r).Value
        .Fields("nombre") = Range("D" & r).Value
        .Fields("dirección") = Range("E" & r).Value
        .Fields("población") = Range("F" & r).Value
        .Fields("tel") = Range("G" & r).Value

        .Update
    End With
    r = r + 1

    Loop

    x = rs.RecordCount

    'cerramos
    rs.Close
    Set rs = Nothing

    bd.Close
    Set bd = Nothing

    If Error = "" Then
    MsgBox "Exportación correcta se han creado " & x & " registros."

    End If

End Sub

'grabar registros desde la table excel a la tabla access (Viable para grabar registros)
'No crea conexión entre el libro y access
Sub GrabarAccess()

    Dim Conn As ADODB.Connection
    Dim MiConexion
    Dim Rs As ADODB.Recordset
    Dim MiBase As String
    Dim r As Long ', x As Long

    MiBase = "pruebas.accdb"

    Set Conn = New ADODB.Connection
    MiConexion = Application.ThisWorkbook.Path & Application.PathSeparator & MiBase

    With Conn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open MiConexion
    End With

    Set Rs = New ADODB.Recordset
    Rs.CursorLocation = adUseServer
    Rs.Open Source:="p1", _
        ActiveConnection:=Conn, _
        CursorType:=adOpenDynamic, _
        LockType:=adLockOptimistic, _
        Options:=adCmdTable

    'recogiendo todos los campos en una tabla
    r = 2 'empiezo en la fila 2 de la hoja 1

    Do While Len(Range("A" & r).Formula) > 0
    'repetir hasta la primera celda vacía de  la columna A    

    'Cargar los datos a Tabla de Access

        With Rs
            .AddNew
            .Fields("a") = Range("B" & r).Value
            .Fields("b") = Range("C" & r).Value
            .Fields("c") = Range("D" & r).Value
            .Fields("d") = Range("E" & r).Value
            .Fields("e") = Range("F" & r).Value
            .Fields("f") = Range("G" & r).Value
            .Update
        End With

    r = r + 1

    Loop

    'x = Rs.RecordCount

    
    'Cerrar la conexión
    Rs.Close
    Conn.Close
    Set Rs = Nothing
    Set Conn = Nothing

    MsgBox "Exportación correcta se han creado " & x & " registros."
    
End Sub

'Consultar base de datos access y cargar resultados del recordset en un listbox
Private Sub ConsultarAccess()
    Dim Conn As ADODB.Connection
    Dim MiConexion
    Dim Rs As ADODB.Recordset
    Dim MiBase As String
    Dim Query As String
    Dim i, j

    MiBase = "pruebas.accdb"

    Set Conn = New ADODB.Connection
    MiConexion = Application.ThisWorkbook.Path & Application.PathSeparator & MiBase

    With Conn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open MiConexion
    End With

    Query = "SELECT * FROM p2 WHERE nombre_contacto LIKE '%" & Me.TextBox1.Value & "%'"

    Set Rs = New ADODB.Recordset
    Rs.CursorLocation = adUseServer
    Rs.Open Source:=Query, _
    ActiveConnection:=Conn

    'Validar si la consulta devuelve resultados
    If Rs.EOF And Rs.BOF Then
        'Borrar la conexión al Recordset
        Rs.Close
        Conn.Close
        'Borrar la memoria
        Set Rs = Nothing
        Set Conn = Nothing
        
        MsgBox "No hay resultados para la consulta", vbInformation, "EXCELeINFO"
        Me.ListBox1.Clear
        Exit Sub
    End If

    'Asignar número de columnas
    With Me.ListBox1
        .ColumnCount = Rs.Fields.Count
    End With

    'Recorrer el Recordset
    Rs.MoveFirst
    i = 1

    With Me.ListBox1
        .Clear
        
        'Asignar los encabezados
        .AddItem
        
        For j = 0 To 5
            .List(0, j) = Rs.Fields(j).Name
        Next j
        
        Do
            .AddItem
            .List(i, 0) = Rs![id]
            .List(i, 1) = Rs![tipo_documento]
            .List(i, 2) = Rs![documento]
            .List(i, 3) = Rs![nombre_contacto]
            .List(i, 4) = Rs![nit]
            .List(i, 5) = Rs![razon_social]
            i = i + 1
            Rs.MoveNext
            
        Loop Until Rs.EOF
    End With

    'Cerrar la conexión
    Rs.Close
    Conn.Close
    Set Rs = Nothing
    Set Conn = Nothing

End Sub