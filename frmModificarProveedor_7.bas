'inicializar controles del formulario al cargar
'----------------------------------------------------------------------------------------------


Private Sub UserForm_Initialize()
    Dim Fila As Integer
    Dim Final As Integer
 
        
    With Hoja6 'contacto_proveedor
       
    Final = GetUltimoR(Hoja6)

        For Fila = 2 To Final
            If .Cells(Fila, 3) <> "" Then
                Me.cboNombreContacto.AddItem (.Cells(Fila, 3))
            End If
        Next

    End With
  
    'Poblar combo ciudades

    With Hoja23 'ciudades
       
    Final = GetUltimoR(Hoja23)

        For Fila = 2 To Final
            If .Cells(Fila, 4) <> "" Then
                Me.cboCiudad.AddItem (.Cells(Fila, 4))
            End If
        Next

    End With
    
    
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
    
    'poblar combo TipoDocumento
    'Me.cboTipoDocumento.AddItem "NIT"
    'Me.cboTipoDocumento.AddItem "CEDULA DE CIUDADANIA"

    'poblar combo FormaPago
    Me.cboFormapago.AddItem "CONTADO"
    Me.cboFormapago.AddItem "CREDITO"
   
End Sub

'Convertir entrada de campos texto a mayúsculas


Private Sub txtRazonSocial_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNombreContacto_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtDireccion_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtBarrio_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


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


Private Sub txtCelular_Change()
    Dim Texto As Variant
    Dim Caracter As Variant
    Dim Largo As Integer
    
    On Error Resume Next
    
    Texto = Me.txtCelular.Value
    Largo = Len(Me.txtCelular.Value)
    For i = 1 To Largo
        Caracter = Mid(Texto, i, 1)
        If Caracter <> "" Then
            If Caracter < Chr(48) Or Caracter > Chr(57) Then
                Me.txtCelular.Value = Replace(Texto, Caracter, "")
            Else
            End If
        End If
    Next i
    On Error GoTo 0
    Caracter = 0
    Caracter1 = 0
End Sub


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



Private Sub cboNombreContacto_Change()
    Dim Fila As Long
    Dim Final As Long
    Dim idProveedor As Integer
    
    'Me.txtRazonSocial = Empty
    'Me.txtDocumento = Empty
    'Me.cboTipoDocumento = Empty
    Me.cboFormapago = Empty
    Me.txtCelular = Empty
    Me.txtTelefono = Empty
    Me.txtCorreo = Empty
    Me.txtDireccion = Empty
    Me.txtBarrio = Empty
    Me.cboCiudad = Empty
    
  
    With Hoja6 ' contacto_proveedor
                    
        Final = GetUltimoR(Hoja6)
    
        For Fila = 2 To Final
            If .Cells(Fila, 3) = cboNombreContacto Then
                Me.txtCelular = .Cells(Fila, 4)
                Me.txtTelefono = .Cells(Fila, 5)
                Me.txtCorreo = .Cells(Fila, 7)
                Me.txtDireccion = .Cells(Fila, 6)
                Me.txtBarrio = .Cells(Fila, 8)
                Me.cboCiudad = .Cells(Fila, 9)
                
                idProveedor = .Cells(Fila, 2)
            End If
                        
        Next
    
    End With
    
    With Hoja4 ' proveedores
                    
        Final = GetUltimoR(Hoja4)
    
        For Fila = 2 To Final
            If .Cells(Fila, 1) = idProveedor Then
                'Me.txtRazonSocial = .Cells(Fila, 4)
                'Me.txtDocumento = .Cells(Fila, 3)
                Me.cboTipoContribuyente = .Cells(Fila, 6)
                Me.cboFormapago = .Cells(Fila, 5)
                
            End If
                        
        Next
    
    End With
    
    
End Sub

Private Sub cmdGuardar_Click()

    Dim Fila As Integer
    Dim Final As Integer
    
    Dim id As Integer
    Dim idProveedor As Integer
    
       
    
    Dim Conn As ADODB.Connection
    Dim MiConexion
    'Dim Rs As ADODB.Recordset
    Dim MiBase As String
    Dim Query As String
    Dim Titulo As String
       
    
    
    With Hoja6 'contacto_proveedor
      
        Final = GetUltimoR(Hoja6)
    
        For Fila = 2 To Final
            If Hoja6.Cells(Fila, 3) = (Me.cboNombreContacto.Text) Then
                                
                id = CInt(.Cells(Fila, 1).Value)
                idProveedor = CInt(.Cells(Fila, 2).Value)
                
'                razonSocial = Me.txtRazonSocial
               
             
                Exit For
            End If
        Next
    
    End With
       
    
    On Error GoTo Salir

    Titulo = "Proveedores"

    If MsgBox("Son correctos los datos?" + Chr(13) + "Desea proceder?", vbOKCancel, Titulo) = vbOK Then


        MiBase = "cotizador.accdb"

        Set Conn = New ADODB.Connection
        MiConexion = Application.ThisWorkbook.Path & Application.PathSeparator & MiBase

        With Conn
            .Provider = "Microsoft.ACE.OLEDB.12.0"
            .Open MiConexion
        End With
              
       
       'Actualizar datos de la tabla proveedores
       Query = "UPDATE proveedores SET forma_pago= '" & Me.cboFormapago.Value & "', tipo_contribuyente= '" & Me.cboTipoContribuyente.Value & "'" & _
            "WHERE id = " & idProveedor & ""
       
       Conn.Execute Query
       
       'Actualizar datos de la tabla contacto_proveedor
       Query = "UPDATE contacto_proveedor SET nombre_contacto = '" & Me.cboNombreContacto.Value & "'" & _
                ", celular = '" & Me.txtCelular.Value & "', telefono= '" & Me.txtTelefono.Value & "', direccion= '" & Me.txtDireccion.Value & "'" & _
                ", correo = '" & Me.txtCorreo.Value & "', barrio= '" & Me.txtBarrio.Value & "', ciudad= '" & Me.cboCiudad.Value & "'" & _
                "WHERE id_proveedor = " & idProveedor & ""
       
       Conn.Execute Query
       

       Conn.Close

        MsgBox "Modificación exitosa", vbInformation

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
    Dim xComboBox As Control
    
        
        For Each xTextBox In Controls
            If xTextBox.Name Like "txt*" Then
                xTextBox = Empty
                Me.cboNombreContacto.SetFocus
            End If
        Next

        For Each xComboBox In Controls
            If xComboBox.Name Like "cbo*" Then
                xComboBox = Empty
                Me.cboNombreContacto.SetFocus
            End If
        Next
        
End Sub
