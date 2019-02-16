VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRegistrarCliente_15 
   Caption         =   "Registrar Cliente"
   ClientHeight    =   7275
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15180
   OleObjectBlob   =   "frmRegistrarCliente_15.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmRegistrarCliente_15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Convertir entrada de campos texto a mayúsculas


Private Sub txtNombreContacto_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtRazonSocial_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtDireccion_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtBarrio_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtComercio_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNicho_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtsegmentacion_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtProducto_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtDistribucion_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


'Validar entradas para permitir ingreso de sólo caracteres o núermos dependiendo del tipo de campo

'aceptar sólo números
Private Sub txtDocumento_change()
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

'aceptar sólo números incluida coma para decimales

Private Sub txtCupo_Change()
    Me.txtCupo.BackColor = &HFFFFFF

End Sub

Private Sub txtCupo_Exit(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next
    Me.txtCupo.Value = FormatCurrency(Me.txtCupo.Value, 2)
End Sub

Private Sub txtCupo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Application.DecimalSeparator = "." Then
        If KeyAscii <> 46 And KeyAscii < 48 Or KeyAscii > 57 Then
            KeyAscii = 0
        End If
    Else
        If KeyAscii <> 44 And KeyAscii < 48 Or KeyAscii > 57 Then
            KeyAscii = 0
        End If
    End If
End Sub

'aceptar sólo números incluida coma para decimales

Private Sub txtCredito_Change()
    Me.txtCupo.BackColor = &HFFFFFF

End Sub

Private Sub txtCredito_Exit(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next
    Me.txtCredito.Value = FormatCurrency(Me.txtCredito.Value, 2)
End Sub

Private Sub txtCredito_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Application.DecimalSeparator = "." Then
        If KeyAscii <> 46 And KeyAscii < 48 Or KeyAscii > 57 Then
            KeyAscii = 0
        End If
    Else
        If KeyAscii <> 44 And KeyAscii < 48 Or KeyAscii > 57 Then
            KeyAscii = 0
        End If
    End If
End Sub


'aceptar sólo números incluida coma para decimales

Private Sub txtSaldo_Change()
    Me.txtSaldo.BackColor = &HFFFFFF

End Sub

Private Sub txtSaldo_Exit(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next
    Me.txtSaldo.Value = FormatCurrency(Me.txtSaldo.Value, 2)
End Sub

Private Sub txtSaldo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Application.DecimalSeparator = "." Then
        If KeyAscii <> 46 And KeyAscii < 48 Or KeyAscii > 57 Then
            KeyAscii = 0
        End If
    Else
        If KeyAscii <> 44 And KeyAscii < 48 Or KeyAscii > 57 Then
            KeyAscii = 0
        End If
    End If
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
    Me.cboTipoDocumento.AddItem "CEDULA CIUDADANIA"
    Me.cboTipoDocumento.AddItem "CEDULA EXTRANJERIA"
    

 'poblar combo TipoContribuyente
    Me.cboTipoContribuyente.AddItem "PERSONA NATURAL REG. COMUN"
    Me.cboTipoContribuyente.AddItem "PERSONA NATURAL REG. SIMPLIFICADO"
    Me.cboTipoContribuyente.AddItem "PERSONA NATURAL AUTORETENEDOR"
    Me.cboTipoContribuyente.AddItem "PERSONA NATURAL O JURIDICA LEY 1429"
    Me.cboTipoContribuyente.AddItem "PERSONA NATURAL REG. COMUN AGENTE AUTORETENEDOR"
    Me.cboTipoContribuyente.AddItem "PERSONA JURIDICA"
    Me.cboTipoContribuyente.AddItem "PERSONA JURIDICA AUTORETENEDOR"
    Me.cboTipoContribuyente.AddItem "GRAN CONTRIBUYENTE AUTORETENEDOR"
    Me.cboTipoContribuyente.AddItem "GRAN CONTRIBUYENTE NO AUTORETENEDOR"
    Me.cboTipoContribuyente.AddItem "ENTIDADES SIN ANIMO DE LUCRO"
    Me.cboTipoContribuyente.AddItem "INSTITUCIONES DEL ESTADO PUBLICOS Y OTROS"
    Me.cboTipoContribuyente.AddItem "PROVEEDOR SOCIEDADES DE CCIO. INTERNACIONAL"
    Me.cboTipoContribuyente.AddItem "TERCERO DEL EXTERIOR"
    
 'poblar combo Categoría
    Me.cboCategoria.AddItem "A"
    Me.cboCategoria.AddItem "B"
    Me.cboCategoria.AddItem "c"
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
        Query = "SELECT id FROM clientes WHERE nombre_contacto LIKE '%" & Me.txtNombreContacto.Value & "%'"
        'Query = "SELECT id FROM clientes WHERE nombre_contacto = '" & Me.txtNombreContacto.Value & "'"
    
        Set Rs = New ADODB.Recordset
        Rs.CursorLocation = adUseServer
        Rs.Open Source:=Query, _
        ActiveConnection:=Conn
    
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





