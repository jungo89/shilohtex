'aceptar sólo números
Private Sub txtCupo_Change()
    Dim Texto As Variant
    Dim Caracter As Variant
    Dim Largo As Integer
    
    On Error Resume Next
    
    Texto = Me.txtCupo.Value
    Largo = Len(Me.txtCupo.Value)
    For i = 1 To Largo
        Caracter = Mid(Texto, i, 1)
        If Caracter <> "" Then
            If Caracter < Chr(48) Or Caracter > Chr(57) Then
                Me.txtCupo.Value = Replace(Texto, Caracter, "")
            Else
            End If
        End If
    Next i
    On Error GoTo 0
    Caracter = 0
    Caracter1 = 0
End Sub

'aceptar sólo números
Private Sub txtCredito_Change()
    Dim Texto As Variant
    Dim Caracter As Variant
    Dim Largo As Integer
    
    On Error Resume Next
    
    Texto = Me.txtCredito.Value
    Largo = Len(Me.txtCredito.Value)
    For i = 1 To Largo
        Caracter = Mid(Texto, i, 1)
        If Caracter <> "" Then
            If Caracter < Chr(48) Or Caracter > Chr(57) Then
                Me.txtCredito.Value = Replace(Texto, Caracter, "")
            Else
            End If
        End If
    Next i
    On Error GoTo 0
    Caracter = 0
    Caracter1 = 0
End Sub

'aceptar sólo números
Private Sub txtSaldo_Change()
    Dim Texto As Variant
    Dim Caracter As Variant
    Dim Largo As Integer
    
    On Error Resume Next
    
    Texto = Me.txtSaldo.Value
    Largo = Len(Me.txtSaldo.Value)
    For i = 1 To Largo
        Caracter = Mid(Texto, i, 1)
        If Caracter <> "" Then
            If Caracter < Chr(48) Or Caracter > Chr(57) Then
                Me.txtSaldo.Value = Replace(Texto, Caracter, "")
            Else
            End If
        End If
    Next i
    On Error GoTo 0
    Caracter = 0
    Caracter1 = 0
End Sub

Private Sub txtNombreContacto_AfterUpdate()
'Determina el final del listado de Clientes
        Final = GetNuevoR(Hoja5)
        
        'Validación para impedir Clientes repetidos
        For Fila = 2 To Final
            If Me.txtNombreContacto.Text <> "" And UCase(Hoja5.Cells(Fila, 2)) = UCase(Me.txtNombreContacto.Text) Then
                MsgBox ("Cliente ya existe en la Base de Datos"), , Titulo
                LimpiarControles
                Me.txtNombreContacto.SetFocus
                Exit Sub
                Exit For
            End If
        Next
End Sub


Private Sub UserForm_Initialize()
    Dim Conn As ADODB.Connection
    Dim MiConexion
    Dim Rs As ADODB.Recordset
    Dim MiBase As String
    Dim Query As String
    
    MiBase = "pruebas.accdb"

    Set Conn = New ADODB.Connection
    MiConexion = Application.ThisWorkbook.Path & Application.PathSeparator & MiBase

    With Conn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open MiConexion
    End With
    
    'traer datos del cliente para verificar
    Query = "SELECT id, nombre_contacto FROM clientes"

    Set Rs = New ADODB.Recordset
    Rs.CursorLocation = adUseServer
    Rs.Open Source:=Query, _
    ActiveConnection:=Conn
    
    Sheets("clientes").Range("A2").CurrentRegion.Clear
    
    For i = 0 To Rs.Fields.Count - 1
    
        Cells(1, i + 1).Value = Rs.Fields(i).Name
    
    Next i
    
    Sheets("clientes").Range("A2").CopyFromRecordset Rs
   
End Sub


Private Sub cmdGuardar_Click()

    Dim Conn As ADODB.Connection
    Dim MiConexion
    Dim Rs As ADODB.Recordset
    Dim MiBase As String
    Dim Query As String
    Dim Titulo As String
    Dim xTextBox As Control
        
    On Error GoTo Salir
    
    Titulo = "Clientes"
    
    For Each xTextBox In Controls
        If xTextBox.Name Like "txt*" And xTextBox = Empty Then
            MsgBox "Debe completar todos los campos", , Titulo
            xTextBox.SetFocus
            Exit Sub
        End If
    Next
    
      
        
    If MsgBox("Son correctos los datos?" + Chr(13) + "Desea proceder?", vbOKCancel, Titulo) = vbOK Then
                
     
        MiBase = "pruebas.accdb"
    
        Set Conn = New ADODB.Connection
        MiConexion = Application.ThisWorkbook.Path & Application.PathSeparator & MiBase
    
        With Conn
            .Provider = "Microsoft.ACE.OLEDB.12.0"
            .Open MiConexion
        End With
    
    
        'crear recordset clientes
        Set Rs = New ADODB.Recordset
        Rs.CursorLocation = adUseServer
        Rs.Open Source:="clientes", _
            ActiveConnection:=Conn, _
            CursorType:=adOpenDynamic, _
            LockType:=adLockOptimistic, _
            Options:=adCmdTable
    
    
        'Cargar los datos a tabla clientes
        With Rs
            .AddNew
            .Fields("nombre_contacto") = txtNombreContacto
            .Fields("tipo_documento") = txtTipoDocumento
            .Fields("documento") = txtDocumento
            .Fields("razon_social") = txtRazonSocial
            .Fields("comercio") = txtComercio
            .Fields("nicho") = txtNicho
            .Fields("segmentacion") = txtSegmentacion
            .Fields("producto") = txtProducto
            .Fields("distribucion") = txtDistribucion
            .Fields("cupo") = CCur(txtCupo)
            .Fields("credito") = CCur(txtCredito)
            .Fields("saldo") = CCur(txtSaldo)
            .Fields("categoria") = txtCategoria
        End With
    
        Rs.Update
        Rs.Close
    
        'determinar el id del registro que se graba
        Query = "SELECT id FROM clientes WHERE nombre_contacto LIKE '%" & Me.txtNombreContacto.Value & "%'"
    
        Set Rs = New ADODB.Recordset
        Rs.CursorLocation = adUseServer
        Rs.Open Source:=Query, _
        ActiveConnection:=Conn
    
        'guardar el id en una hoja auxiliar
        Sheets("contadores").Range("A1").CurrentRegion.Clear
        Sheets("contadores").Range("A2").CopyFromRecordset Rs
    
        Rs.Close
    
        'crear recordset contacto cliente
        Set Rs = New ADODB.Recordset
        Rs.CursorLocation = adUseServer
        Rs.Open Source:="contacto_cliente", _
            ActiveConnection:=Conn, _
            CursorType:=adOpenDynamic, _
            LockType:=adLockOptimistic, _
            Options:=adCmdTable
    
    
        'Cargar los datos a tabla datos_cliente
        With Rs
            .AddNew
            .Fields("id_cliente") = Sheets("contadores").Range("A2").Value
            .Fields("telefono") = txtTelefono
            .Fields("direccion") = txtDireccion
            .Fields("correo") = txtCorreo
            .Fields("barrio") = txtBarrio
            .Fields("ciudad") = txtCiudad
        End With
    
        'Cerrar la conexión
        Rs.Update
        Rs.Close
    
    
        Conn.Close
        Set Rs = Nothing
        Set Conn = Nothing
    
        MsgBox "Alta exitosa", vbInformation
        
        'Limpia los controles
        LimpiarControles
        
    Else
            Exit Sub
    End If
    
       
Salir:
     If Err <> 0 Then
        MsgBox Err.Description, vbExclamation, Titulo
     End If
    
End Sub


Private Sub LimpiarControles()
    Dim xTextBox As Control
        
        For Each xTextBox In Controls
            If xTextBox.Name Like "txt*" Then
                xTextBox = Empty
                Me.txtNombreContacto.SetFocus
            End If
        Next

End Sub

