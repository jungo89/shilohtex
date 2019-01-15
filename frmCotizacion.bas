'aceptar sólo números incluida coma para decimales

Private Sub txtCupo_Change()
    Me.txtCupo.BackColor = &HFFFFFF
    
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

Private Sub txtCredito_Change()
    Me.txtCredito.BackColor = &HFFFFFF
    
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

Private Sub txtSaldo_Change()
    Me.txtSaldo.BackColor = &HFFFFFF
    
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

Private Sub txtValorUnitario_Change()
    Me.txtValorUnitario.BackColor = &HFFFFFF
    
    On Error Resume Next
    Me.txtValorUnitario.Value = FormatCurrency(Me.txtValorUnitario.Value, 2)
End Sub

Private Sub txtValorUnitario_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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

Private Sub txtSubTotalCotizado_Change()
    Me.txtSubTotalCotizado.BackColor = &HFFFFFF
    
    On Error Resume Next
    Me.txtSubTotalCotizado.Value = FormatCurrency(Me.txtSubTotalCotizado.Value, 2)
End Sub

Private Sub txtSubTotalCotizado_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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

Private Sub txtValor30Dias_Change()
    Me.txtValor30Dias.BackColor = &HFFFFFF
    
    On Error Resume Next
    Me.txtValor30Dias.Value = FormatCurrency(Me.txtValor30Dias.Value, 2)
End Sub

Private Sub txtValor30Dias_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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

Private Sub txtValor60Dias_Change()
    Me.txtValor60Dias.BackColor = &HFFFFFF
    
    On Error Resume Next
    Me.txtValor60Dias.Value = FormatCurrency(Me.txtValor60Dias.Value, 2)
End Sub

Private Sub txtValor60Dias_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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

Private Sub txtReteFuente_Change()
    Me.txtReteFuente.BackColor = &HFFFFFF
    
    On Error Resume Next
    Me.txtReteFuente.Value = FormatCurrency(Me.txtReteFuente.Value, 2)
End Sub

Private Sub txtReteFuente_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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

Private Sub txtReteIca_Change()
    Me.txtReteIca.BackColor = &HFFFFFF
    
    On Error Resume Next
    Me.txtReteIca.Value = FormatCurrency(Me.txtReteIca.Value, 2)
End Sub

Private Sub txtReteIca_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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

Private Sub txtTotalCotizado_Change()
    Me.txtTotalCotizado.BackColor = &HFFFFFF
    
    On Error Resume Next
    Me.txtTotalCotizado.Value = FormatCurrency(Me.txtTotalCotizado.Value, 2)
End Sub

Private Sub txtTotalCotizado_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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

Private Sub UserForm_Initialize()
    Dim Fila As Integer
    Dim Final As Integer
 
    With Hoja1 'clientes
       
    Final = GetUltimoR(Hoja1)

        For Fila = 2 To Final
            If .Cells(Fila, 4) <> "" Then
                Me.cboNombreContacto.AddItem (.Cells(Fila, 4))
            End If
        Next

    End With
    
    With Hoja4 'proveedores
       
    Final = GetUltimoR(Hoja4)

        For Fila = 2 To Final
            If .Cells(Fila, 2) <> "" Then
                Me.cboProveedor.AddItem (.Cells(Fila, 2))
            End If
        Next

    End With
    
    With Hoja9 'empleados
       
    Final = GetUltimoR(Hoja9)

        For Fila = 2 To Final
            If .Cells(Fila, 2) <> "" And .Cells(Fila, 3) = "ASESORA COMERCIAL" Then
                Me.cboAsesora.AddItem (.Cells(Fila, 2))
            End If
        Next

    End With
    
    With Hoja9 'empleados
       
    Final = GetUltimoR(Hoja9)

        For Fila = 2 To Final
            If .Cells(Fila, 2) <> "" And .Cells(Fila, 3) = "AUXILIAR DE BODEGA" Then
                Me.cboBodega.AddItem (.Cells(Fila, 2))
            End If
        Next

    End With
    
    With Hoja19 'transportadores
       
    Final = GetUltimoR(Hoja19)

        For Fila = 2 To Final
            If .Cells(Fila, 2) <> "" Then
                Me.cboTransportador.AddItem (.Cells(Fila, 2))
            End If
        Next

    End With
    
 
    'combo forma de pago
    Me.cboFormaDePago.AddItem "CONTADO"
    Me.cboFormaDePago.AddItem "CONTRA ENTREGA"
    Me.cboFormaDePago.AddItem "CREDITO"
    
    'combo prioridad
    Me.cboPrioridad.AddItem "INMEDIATO"
    Me.cboPrioridad.AddItem "FELIPE"
    Me.cboPrioridad.AddItem "LOGIFUTURO"
   
   
End Sub


Private Sub cboNombreContacto_Change()
    Dim Fila As Long
    Dim Final As Long
    
    txtRazonSocial = Empty
    txtDocumento = Empty
    txtNit = Empty
    cboTelefono.Clear
    cboCorreo.Clear
    cboDireccion.Clear
    cboBarrio.Clear
    cboCiudad.Clear
    txtTipoContribuyente = Empty
    txtNicho = Empty
    txtCupo.Text = Empty
    txtCredito.Text = Empty
    txtSaldo.Text = Empty
    txtInteres.Text = Empty
    txtCategoria = Empty
    
    
    With Hoja1 ' clientes
                    
        Final = GetUltimoR(Hoja1)
    
        For Fila = 2 To Final
            If .Cells(Fila, 4) = cboNombreContacto Then
                txtRazonSocial.Text = (.Cells(Fila, 6))
                txtDocumento.Text = (.Cells(Fila, 3))
                txtNit.Text = (.Cells(Fila, 5))
                txtTipoContribuyente.Text = (.Cells(Fila, 16))
                txtNicho.Text = (.Cells(Fila, 8))
                txtCupo.Text = (.Cells(Fila, 12))
                txtCredito.Text = (.Cells(Fila, 13))
                txtSaldo.Text = (.Cells(Fila, 14))
                txtCategoria.Text = (.Cells(Fila, 15))
            End If
        Next
    
    End With
    
    With Hoja5 ' datos_cliente
                    
        Final = GetUltimoR(Hoja5)
    
        For Fila = 2 To Final
            If .Cells(Fila, 8) = cboNombreContacto Then
                cboTelefono.AddItem (.Cells(Fila, 3))
                cboDireccion.AddItem (.Cells(Fila, 4))
                cboCorreo.AddItem (.Cells(Fila, 5))
                cboBarrio.AddItem (.Cells(Fila, 6))
                cboCiudad.AddItem (.Cells(Fila, 7))
            End If
        Next
    
    End With

End Sub

Private Sub cboFormaDePago_Change()
    CboDias.Clear
    CboDias.Enabled = True
    txtInteres.Enabled = True
    
    lbl30Dias.Visible = True
    lblHasta30Dias.Visible = True
    txtFecha30Dias.Visible = True
    txtValor30Dias.Visible = True
    
    lbl60Dias.Visible = True
    lblHasta60Dias.Visible = True
    txtFecha60Dias.Visible = True
    txtValor60Dias.Visible = True
    
    If cboFormaDePago <> "CREDITO" Then
        CboDias.Enabled = False
        txtInteres.Enabled = False
        
        lbl30Dias.Visible = False
        lblHasta30Dias.Visible = False
        txtFecha30Dias.Visible = False
        txtValor30Dias.Visible = False
    
        lbl60Dias.Visible = False
        lblHasta60Dias.Visible = False
        txtFecha60Dias.Visible = False
        txtValor60Dias.Visible = False
        
    Else
        CboDias.Enabled = True
        txtInteres.Enabled = True
        For i = 30 To 60 Step 30
            CboDias.AddItem i
        Next i
    End If
    
End Sub


Private Sub cboProveedor_Change()
    Dim Fila As Long
    Dim Final As Long
    
    cboProducto.Clear
    cboColor.Clear
    txtCantidad = Empty
    txtValorUnitario = Empty
    txtDisponible = Empty
    txtStock = Empty
    txtPedir = Empty
    
  
    With Hoja2 ' productos
                    
        Final = GetUltimoR(Hoja2)
    
        For Fila = 2 To Final
            If .Cells(Fila, 17) = cboProveedor Then
                 Agregar cboProducto, .Cells(Fila, 3)
            End If
        Next
    
    End With
End Sub

Private Sub cboProducto_Change()
    Dim Fila As Long
    Dim Final As Long
    
    cboColor.Clear
    txtCantidad = Empty
    txtValorUnitario = Empty
    txtDisponible = Empty
    txtStock = Empty
    txtPedir = Empty
    
  
    With Hoja2 ' productos
                    
        Final = GetUltimoR(Hoja2)
    
        For Fila = 2 To Final
            If .Cells(Fila, 3) = cboProducto Then
                 Agregar cboColor, .Cells(Fila, 4)
                 'txtValorUnitario = .Cells(Fila, 10)
                 
            End If
        Next
    
    End With

End Sub

Private Sub cboColor_Change()
    Dim Fila As Long
    Dim Final As Long
    
    txtCantidad = Empty
    txtValorUnitario = Empty
    txtDisponible = Empty
    txtStock = Empty
    txtPedir = Empty
    
  
    With Hoja2 ' productos
                    
        Final = GetUltimoR(Hoja2)
    
        For Fila = 2 To Final
            If .Cells(Fila, 17) = cboProveedor And .Cells(Fila, 3) = cboProducto And .Cells(Fila, 4) = cboColor Then
                 txtValorUnitario = .Cells(Fila, 10)
                 txtCantidad = .Cells(Fila, 6) & " Por " & .Cells(Fila, 7)
                 txtDisponible = .Cells(Fila, 14)
                 txtStock = .Cells(Fila, 15)
                 txtPedir = .Cells(Fila, 16)
            End If
        Next
    
    End With
End Sub

Private Sub btnFechaElaboracion_Click()
banderaCalendario = 1
    Call LanzarCalendario(Me, "txtFechaElaboracion")
End Sub

Private Sub btnFechaEntrega_Click()
banderaCalendario = 2
    Call LanzarCalendario(Me, "txtFechaEntrega")
End Sub