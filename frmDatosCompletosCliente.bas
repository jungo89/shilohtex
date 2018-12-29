
'convertir a mayusculas contenido del textbox
Private Sub txtNombreContacto_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub




'aceptar sólo números
Private Sub txtDocumento_Change()
    Dim Texto As Variant
    Dim Caracter As Variant
    Dim Largo As Integer
    
    On Error Resume Next
    
    Texto = Me.txtDocumento.Value
    Largo = Len(Me.txtDocumento.Value)
    For i = 1 To Largo
        Caracter = Mid(Texto, i, 1)
        If Caracter <> "" Then
            If Caracter < Chr(48) Or Caracter > Chr(57) Then
                Me.txtDocumento.Value = Replace(Texto, Caracter, "")
            Else
            End If
        End If
    Next i
    On Error GoTo 0
    Caracter = 0
    Caracter1 = 0
End Sub


'aceptar sólo números
Private Sub txtTelefono_Change()
    Dim Texto As Variant
    Dim Caracter As Variant
    Dim Largo As Integer
    
    On Error Resume Next
    
    Texto = Me.txtTelefono.Value
    Largo = Len(Me.txtTelefono.Value)
    For i = 1 To Largo
        Caracter = Mid(Texto, i, 1)
        If Caracter <> "" Then
            If Caracter < Chr(48) Or Caracter > Chr(57) Then
                Me.txtTelefono.Value = Replace(Texto, Caracter, "")
            Else
            End If
        End If
    Next i
    On Error GoTo 0
    Caracter = 0
    Caracter1 = 0
End Sub

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
        Final = GetNuevoR(Hoja1)
        
        'Validación para impedir Clientes repetidos
        For Fila = 2 To Final
            If Me.txtNombreContacto.Text <> "" And UCase(Hoja1.Cells(Fila, 4)) = UCase(Me.txtNombreContacto.Text) Then
                MsgBox ("Cliente ya existe en la Base de Datos"), , Titulo
                LimpiarControles
                Me.txtNombreContacto.SetFocus
                Exit Sub
                Exit For
            End If
        Next
End Sub


Private Sub UserForm_Initialize()

'Call CopiarClientes

'Call CopiarContactoCliente

'Poblar combo ciudades
    Dim Fila As Integer
    Dim Final As Integer
 
    With Hoja23 'ciudades
       
    Final = GetUltimoR(Hoja23)

        For Fila = 2 To Final
            If .Cells(Fila, 4) <> "" Then
                Me.cboCiudad.AddItem (.Cells(Fila, 4))
            End If
        Next

    End With
    
    
'poblar combo TipoDocumento
    Me.cboTipoDocumento.AddItem "PERSONA JURIDICA"
    Me.cboTipoDocumento.AddItem "PERSONA NATURAL"
    Me.cboTipoDocumento.AddItem "REGIMEN SIMPLIFICADO"

 'poblar combo TipoContribuyente
    Me.cboTipoContribuyente.AddItem "GRAN CONTRIBUYENTE"
    Me.cboTipoContribuyente.AddItem "CONTRIBUYENTE MEDIANO ALTO"
    Me.cboTipoContribuyente.AddItem "CONTRIBUYENTE MEDIANO"
    Me.cboTipoContribuyente.AddItem "CONTRIBUYENTE PEQUEÑO"
    Me.cboTipoContribuyente.AddItem "CONTRIBUYENTE MICRO"
    
 'poblar combo Categoría
    Me.cboCategoria.AddItem "A"
    Me.cboCategoria.AddItem "C"
    Me.cboCategoria.AddItem "V"
    
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
                
     
        MiBase = "cotizador.accdb"
    
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
            .Fields("tipo_documento") = cboTipoDocumento
            .Fields("documento") = txtDocumento
            .Fields("nombre_contacto") = txtNombreContacto
            .Fields("nit") = txtNit
            .Fields("razon_social") = txtRazonSocial
            .Fields("comercio") = txtComercio
            .Fields("nicho") = txtNicho
            .Fields("segmentacion") = txtSegmentacion
            .Fields("producto") = txtProducto
            .Fields("distribucion") = txtDistribucion
            .Fields("cupo") = CCur(txtCupo)
            .Fields("credito") = CCur(txtCredito)
            .Fields("saldo") = CCur(txtSaldo)
            .Fields("categoria") = cboCategoria
            .Fields("tipo_contribuyente") = cboTipoContribuyente
        End With
    
        Rs.Update
        Rs.Close
    
        'determinar el id del registro que se graba
        'Query = "SELECT id FROM clientes WHERE nombre_contacto LIKE '%" & Me.txtNombreContacto.Value & "%'"
        Query = "SELECT id FROM clientes WHERE nombre_contacto = '" & Me.txtNombreContacto.Value & "'"
    
        Set Rs = New ADODB.Recordset
        Rs.CursorLocation = adUseServer
        Rs.Open Source:=Query, _
        ActiveConnection:=Conn
    
        'guardar el id en una hoja auxiliar
        'Sheets("contadores").Range("A1").CurrentRegion.Clear
        'Sheets("contadores").Range("A2").CopyFromRecordset Rs
        
           
        Sheets("contadores").Range("A2").Select
        Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
        Selection.ClearContents
        
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
            .Fields("ciudad") = cboCiudad
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
        
        Me.cboCiudad = Empty
        Me.cboTipoContribuyente = Empty
        Me.cboTipoDocumento = Empty
        Me.cboCategoria = Empty

End Sub





