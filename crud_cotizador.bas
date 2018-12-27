

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
        
        MsgBox "No hay resultados para la consulta", vbInformation
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

'Trae datos a una hoja excel utilizando un formulario para aplicar filtros a la consulta
sub FiltrarAccess()
    Dim Conn As ADODB.Connection
    Dim MiConexion
    Dim Rs As ADODB.Recordset
    Dim MiBase As String
    Dim Query As String
    Dim i, j
    Dim Fecha1
    Dim Fecha2
    Dim Ciudad As String

    MiBase = "pruebas.accdb"

    Set Conn = New ADODB.Connection
    MiConexion = Application.ThisWorkbook.Path & Application.PathSeparator & MiBase

    With Conn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open MiConexion
    End With

    Fecha1 = Me.TextBox1.Value
    Fecha2 = Me.TextBox2.Value
    Ciudad = Me.TextBox3.Value
    Query = "SELECT * FROM p3 WHERE [Fecha] >= #" & Fecha1 & "# AND [Fecha] <= #" & Fecha2 & "# AND [ciudad] = '" & Ciudad & "' ORDER BY [Fecha]"

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
        
        MsgBox "No hay resultados para la consulta", vbInformation
        Sheets("Reporte").Range("A1").CurrentRegion.Clear
        Exit Sub
    End If

    Sheets("Reporte").Range("A1").CurrentRegion.Clear
    For i = 0 To Rs.Fields.Count - 1

        Cells(1, i + 1).Value = Rs.Fields(i).Name

    Next i

    Sheets("Reporte").Range("A2").CopyFromRecordset Rs

    'Cerrar la conexión
    Rs.Close
    Conn.Close
    Set Rs = Nothing
    Set Conn = Nothing
end sub
    
