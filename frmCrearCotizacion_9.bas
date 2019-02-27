
'cotizar (Cantidad solicitada por el cliente)
'cotizado (Lo que realmente se le vende al cliente)
'pendiente (diferencia entre lo solicitado y lo entregado
'unidades (Unidad de empaque del producto)
'valor unitario flete
'valor total flete
'producto
'color
'medida
'porcentaje iva
'valor unitario
'subtotal cotizado

'cambiar formatos porcentuales sin decimal
'quitar decimales a l

'Encabezados del listbox en mayúsculas

Dim i As Long
'Dim sTotal As Currency


Private Sub lblEliminarItem_Click()
    EliminarItem
End Sub

Private Sub lblProductos_Click()
    AgregarItems
End Sub


'Procedimientos para implementar reglas de negocios de la cotización
'----------------------------------------------------------------------------------------------

Public Sub AgregarItems()
'Agrega los items al listbox

'Dim sTotal As Currency

        If Me.cboProveedor.Text = "" Or Me.cboProducto.Text = "" Or Me.cboColor.Text = "" Then MsgBox ("Elija un producto"): Exit Sub
        If Trim(Me.txtUnidades.Text) = "" Or Me.txtUnidades = Empty Or Trim(Me.txtUnidadesSolicitadas.Text) = "" Or Me.txtUnidadesSolicitadas = Empty Then MsgBox ("Debe ingresar la unidades"): Exit Sub
       
       'MsgBox (Me.txtCantidadSolicitada.Text)
       'MsgBox (Me.txtCantidad.Text)
       'MsgBox (Me.txtCantidadSolicitada.Text - Me.txtCantidad.Text)
       'MsgBox (Me.txtUnidades.Text)
       'MsgBox (0)
       'MsgBox (0)
       'MsgBox (Me.cboProducto.Value)
       'MsgBox (Me.cboColor.Value)
       'MsgBox (Me.txtMedida.Text)
       
       'MsgBox (Me.txtValorUnitarioIva.Text)
       'MsgBox (Me.txtSubtotal.Text)
       
       
       
       
       
        With frmCotizacion
            .lstDetalleFact1.AddItem Me.txtCantidadSolicitada.Text 'cantidad solicitada unidades por producto
            .lstDetalleFact1.List(i, 1) = Me.txtCantidad.Text 'cantidad entregada unidades por producto
            .lstDetalleFact1.List(i, 2) = Me.txtCantidadSolicitada.Text - Me.txtCantidad.Text 'cantidad pendiente por entregar
            .lstDetalleFact1.List(i, 3) = Me.txtUnidades.Text 'unidades entregadas
            .lstDetalleFact1.List(i, 4) = 0 'valor unitario flete
            .lstDetalleFact1.List(i, 5) = 0 'valor total flete
            
            .lstDetalleFact1.List(i, 6) = Me.cboProducto.Value 'producto
            .lstDetalleFact1.List(i, 7) = Me.txtMedida.Text 'producto
            .lstDetalleFact1.List(i, 8) = Me.cboColor.Value 'color
            .lstDetalleFact1.List(i, 9) = Me.cboIva.Value 'porcentaje de iva
                        
            .lstDetalleFact2.AddItem Me.txtValorUnitarioIva.Text 'valor unitario con iva
            .lstDetalleFact2.List(i, 1) = Me.txtSubtotal.Text 'subtotal
            
            'MsgBox (.lstDetalleFact1.List(i, 6))

        i = i + 1
        End With
        
        sumarImporte
                
        'sTotal = sTotal + (Me.txtSubtotal)
        'Me.txtSubTotalCotizado.Text = sTotal
        
    
        With Me
           '.ComboBox1.ListIndex = -1
            .cboProveedor = Empty
            .cboProducto = Empty
            .cboColor = Empty
            .txtCantidad = ""
            .txtMedida = ""
            .txtValorUnitario = ""
            .txtValorEmpaque = ""
            .cboIva = Empty
            .txtValorUnitarioIva = ""
            .txtValorEmpaqueIva = ""
            .txtUnidades = ""
            '.txtMetros = ""
            .txtSubtotal = ""
            .txtDisponible = ""
            .txtStock = ""
            .txtPedir = ""
        End With

End Sub

Public Sub EliminarItem()

Dim index As Integer

' Elimina el item seleccionado y resta el importe de la columna de importes
' Elimina el item seleccionado y resta el importe de la columna de importes



    If Me.lstDetalleFact1.ListIndex = -1 Then
        MsgBox "Seleccionar un producto para eliminar", vbInformation
        Exit Sub
    End If
    
    Me.lstDetalleFact2.ListIndex = -1 ' Eliminar la "barra de selección"

    index = lstDetalleFact1.ListIndex
    
    Me.lstDetalleFact1.RemoveItem (index)
    Me.lstDetalleFact2.RemoveItem (index)
    
    Me.lstDetalleFact1.ListIndex = -1 ' Eliminar la "barra de selección"
    

Me.sumarImporte

'sTotal = sTotal + (Me.txtSubtotal)
'Me.txtSubTotalCotizado.Text = sTotal
End Sub

Public Sub sumarImporte()

Dim i As Integer
Dim sTotal As Currency


sTotal = 0
        For i = 0 To Me.lstDetalleFact2.ListCount - 1
        
            sTotal = sTotal + Me.lstDetalleFact2.List(i, 1) 'Aquí hago la sumatoria del importe, utilizando el punto decimal

        Next i
        'MsgBox (sTotal)
        
Me.txtSubTotalCotizado.Text = sTotal


'            If sTotal > 0 Then ' aqui se hacen los calculos para el subtotal, iva y total
'
'                    Me.txtIva.Text = (sTotal / 100) * IvaPorcentaje
'                    xIVA = Me.txtIva.Text
'                    Me.txtTotal.Text = sTotal + xIVA
'                    Me.txtLetras.Text = UCase(cMoneda(Me.txtTotal.Text))
'                Else
'                    Me.txtSubtotal.Text = Empty
'                    Me.txtIva.Text = Empty
'                    Me.txtTotal.Text = Empty
'                    Me.txtLetras.Text = Empty
'            End If
            
End Sub


Private Sub txtCantidad_Change()
    Dim val As String
    val = Me.cboIva
    
    'Me.txtCantidad.BackColor = &HFFFFFF
    
    'On Error Resume Next
    'Me.txtCantidad.Value = FormatNumber(Me.txtCantidad.Value, 2)
    
    Select Case val
                    
        Case Is = "0,0%"
            val = 0#
        Case Is = "1,0%"
            val = 0.01
        Case Is = "1,5%"
            val = 0.015
        Case Is = "2,0%"
            val = 0.02
        Case Is = "2,5%"
            val = 0.025
        Case Is = "3,0%"
            val = 0.03
        Case Is = "3,5%"
            val = 0.035
        Case Is = "4,0%"
            val = 0.04
        Case Is = "4,5%"
            val = 0.045
        Case Is = "5,0%"
            val = 0.05
        Case Is = "5,5%"
            val = 0.055
        Case Is = "6,0%"
            val = 0.06
        Case Else
            val = 0#
    End Select
    

    If Me.txtValorUnitario <> "" And Me.txtUnidades <> "" And Me.txtCantidad <> "" Then
        Me.txtSubtotal = Application.WorksheetFunction.RoundUp((Me.txtValorUnitarioIva * Me.txtUnidades * Me.txtCantidad), 0)
        'Me.txtValorUnitario = Application.WorksheetFunction.RoundUp((Me.txtValorUnitario) * (1 + val), 0)
        'Me.cboInteres.Value = Formatdouble(Me.cboInteres.Value, 2)
        'MsgBox (cboIva)
        'MsgBox (val)
        
    Else
        Me.txtSubtotal = Empty
    End If
End Sub

Private Sub txtCantidad_Exit(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next
    Me.txtCantidad.Value = FormatNumber(Me.txtCantidad.Value, 0)
End Sub

Private Sub txtCantidad_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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

Private Sub txtCantidadSolicitada_Change()

End Sub

'Configurar formato de los controles controles
'----------------------------------------------------------------------------------------------
'aceptar sólo números incluida coma para decimales


Private Sub txtCupo_Change()
    Me.txtCupo.BackColor = &HFFFFFF
    
    On Error Resume Next
    Me.txtCupo.Value = FormatCurrency(Me.txtCupo.Value, 0)
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
    Me.txtCredito.Value = FormatCurrency(Me.txtCredito.Value, 0)
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
    Me.txtSaldo.Value = FormatCurrency(Me.txtSaldo.Value, 0)
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
    Me.txtValorUnitario.Value = FormatCurrency(Me.txtValorUnitario.Value, 0)
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

Private Sub txtValorEmpaque_Change()
    Me.txtValorEmpaque.BackColor = &HFFFFFF

    On Error Resume Next
    Me.txtValorEmpaque.Value = FormatCurrency(Me.txtValorEmpaque.Value, 0)
End Sub

Private Sub txtValorEmpaque_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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

Private Sub txtValorUnitarioIva_Change()
    Me.txtValorUnitarioIva.BackColor = &HFFFFFF

    On Error Resume Next
    Me.txtValorUnitarioIva.Value = FormatCurrency(Me.txtValorUnitarioIva.Value, 0)
End Sub

Private Sub txtValorUnitarioIva_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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

Private Sub txtValorEmpaqueIva_Change()
    Me.txtValorEmpaqueIva.BackColor = &HFFFFFF

    On Error Resume Next
    Me.txtValorEmpaqueIva.Value = FormatCurrency(Me.txtValorEmpaqueIva.Value, 0)
End Sub

Private Sub txtValorEmpaqueIva_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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

Private Sub txtUnidades_Change()
    Me.txtUnidades.BackColor = &HFFFFFF
    
    Dim val As String
    val = Me.cboIva
    
    Select Case val
        Case Is = "0,0%"
            val = 0#
        Case Is = "1,0%"
            val = 0.01
        Case Is = "1,0%"
            val = 0.015
        Case Is = "2,0%"
            val = 0.02
        Case Is = "2,5%"
            val = 0.025
        Case Is = "3,0%"
            val = 0.03
        Case Is = "3,5%"
            val = 0.035
        Case Is = "4,0%"
            val = 0.04
        Case Is = "4,5%"
            val = 0.045
        Case Is = "5,0%"
            val = 0.05
        Case Is = "5,5%"
            val = 0.055
        Case Is = "6,0%"
            val = 0.06
        Case Else
            val = 0#
        
    End Select
    

    If Me.txtValorUnitario <> "" And Me.txtUnidades <> "" Then
        Me.txtSubtotal = Application.WorksheetFunction.RoundUp((Me.txtValorUnitario * Me.txtUnidades * Me.txtCantidad + ((Me.txtValorUnitario * Me.txtUnidades * Me.txtCantidad) * val)), 0)
    Else
        
        Me.txtSubtotal = Empty
    End If

End Sub


Private Sub txtSubtotal_Change()
    Me.txtSubtotal.BackColor = &HFFFFFF

    On Error Resume Next
    Me.txtSubtotal.Value = FormatCurrency(Me.txtSubtotal.Value, 0)
End Sub

Private Sub txtSubtotal_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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
    Me.txtSubTotalCotizado.Value = FormatCurrency(Me.txtSubTotalCotizado.Value, 0)
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
    Me.txtValor30Dias.Value = FormatCurrency(Me.txtValor30Dias.Value, 0)
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
    Me.txtValor60Dias.Value = FormatCurrency(Me.txtValor60Dias.Value, 0)
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
    Me.txtReteFuente.Value = FormatCurrency(Me.txtReteFuente.Value, 0)
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
    Me.txtReteIca.Value = FormatCurrency(Me.txtReteIca.Value, 0)
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
    Me.txtTotalCotizado.Value = FormatCurrency(Me.txtTotalCotizado.Value, 0)
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

Private Sub cboPorcentaje_Change()
    Me.cboPorcentaje.BackColor = &HFFFFFF
    
    On Error Resume Next
    Me.cboPorcentaje.Value = FormatPercent(Me.cboPorcentaje.Value, 1)
    Me.cboIva = Me.cboPorcentaje
   
End Sub


Private Sub cboPorcentaje_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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

Private Sub cboInteres_Change()
    Me.cboInteres.BackColor = &HFFFFFF
    
    On Error Resume Next
    Me.cboInteres.Value = FormatPercent(Me.cboInteres.Value, 1)
    
End Sub

Private Sub cboInteres_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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

Private Sub cboIva_Change()
    Me.cboIva.BackColor = &HFFFFFF
   
    
    On Error Resume Next
    Me.cboIva.Value = FormatPercent(Me.cboIva.Value, 1)
    
    Dim nomIva As String
    Dim valIva As Double
    
    nomIva = Me.cboIva
    
    Select Case nomIva
        Case Is = "0,0%"
            valIva = 0#
        Case Is = "1,0%"
            valIva = 0.01
        Case Is = "1,5%"
            valIva = 0.015
        Case Is = "2,0%"
            valIva = 0.02
        Case Is = "2,5%"
            valIva = 0.025
        Case Is = "3,0%"
            valIva = 0.03
        Case Is = "3,5%"
            valIva = 0.035
        Case Is = "4,0%"
            valIva = 0.04
        Case Is = "4,5%"
            valIva = 0.045
        Case Is = "5,0%"
            valIva = 0.05
        Case Is = "5,5%"
            valIva = 0.055
        Case Is = "6,0%"
            valIva = 0.06
        Case Else
            valIva = 0#
    End Select
    
    'MsgBox (valIva)
      
    txtValorUnitarioIva = Application.WorksheetFunction.RoundUp(txtValorUnitario * (1 + valIva), 0)
    txtValorEmpaqueIva = Application.WorksheetFunction.RoundUp(txtValorEmpaque * (1 + valIva), 0)
    
    If Me.txtValorUnitario <> "" And Me.txtUnidades <> "" And Me.txtCantidad <> "" Then
        Me.txtSubtotal = Application.WorksheetFunction.RoundUp((Me.txtValorUnitarioIva * Me.txtUnidades * Me.txtCantidad), 0)
        'Me.txtValorUnitario = Application.WorksheetFunction.RoundUp((Me.txtValorUnitario) * (1 + val), 0)
        'Me.cboInteres.Value = Formatdouble(Me.cboInteres.Value, 2)
        'MsgBox (cboIva)
        'MsgBox (val)
        
    Else
        Me.txtSubtotal = Empty
    End If
   
End Sub

Private Sub cboIva_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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



'inicializar controles del formulario al cargar
'----------------------------------------------------------------------------------------------

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
    
    'configurar número de factura y tamño de las columnas del listbox
    
'    Me.Label11.Caption = Hoja12.Range("C6") & "% IVA:"
'
'    Me.lbl_nFactura.Caption = "No. " & Hoja7.Range("F2").Value + 1 'Llamamos el número de la factura
'
'
    'With lstDetalleFact1
    '    .ColumnCount = 10
    '    .ColumnWidths = "45 pt;45 pt;45 pt;45 pt;80 pt;80 pt;250 pt;120 pt;80 pt;80 pt" ' Unidades de medida, 72 pt(puntos)=1 Pulgada
    'End With
    
 
    'combo forma de pago
    Me.cboFormaDePago.AddItem "CONTADO"
    Me.cboFormaDePago.AddItem "CONTRA ENTREGA"
    Me.cboFormaDePago.AddItem "CREDITO"
    
    'combo prioridad
    Me.cboPrioridad.AddItem "INMEDIATO"
    Me.cboPrioridad.AddItem "DENTRO DE BOGOTA"
    Me.cboPrioridad.AddItem "FUERA DE BOGOTA"
    
    'combo interes
    Me.cboInteres.AddItem "0,01"
    Me.cboInteres.AddItem "0,015"
    Me.cboInteres.AddItem "0,020"
    Me.cboInteres.AddItem "0,025"
    Me.cboInteres.AddItem "0,030"
    Me.cboInteres.AddItem "0,035"
    Me.cboInteres.AddItem "0,040"
    Me.cboInteres.AddItem "0,045"
    Me.cboInteres.AddItem "0,050"
    Me.cboInteres.AddItem "0,055"
    Me.cboInteres.AddItem "0,060"
    
    
    'combo Iva
    Me.cboIva.AddItem "0,00"
    Me.cboIva.AddItem "0,01"
    Me.cboIva.AddItem "0,015"
    Me.cboIva.AddItem "0,020"
    Me.cboIva.AddItem "0,025"
    Me.cboIva.AddItem "0,030"
    Me.cboIva.AddItem "0,035"
    Me.cboIva.AddItem "0,040"
    Me.cboIva.AddItem "0,045"
    Me.cboIva.AddItem "0,050"
    Me.cboIva.AddItem "0,055"
    Me.cboIva.AddItem "0,060"
    
    Me.cboIva = "0,00"
    
    Me.txtFechaElaboracion.Text = Now
    
    Me.txtFecha30Dias.Text = Date + 35
    Me.txtFecha60Dias.Text = Date + 65
    Me.txtFecha90Dias.Text = Date + 95
    
    Me.cboIva.Enabled = False
    
    
    
   
End Sub


'Eventos de controles
'----------------------------------------------------------------------------------------------

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
            'poblar combos
                cboTelefono.AddItem (.Cells(Fila, 3))
                cboDireccion.AddItem (.Cells(Fila, 4))
                cboCorreo.AddItem (.Cells(Fila, 5))
                cboBarrio.AddItem (.Cells(Fila, 6))
                cboCiudad.AddItem (.Cells(Fila, 7))
                
            'asignar valor por defecto en la primer fila
                cboTelefono = (.Cells(Fila, 3))
                cboDireccion = (.Cells(Fila, 4))
                cboCorreo = (.Cells(Fila, 5))
                cboBarrio = (.Cells(Fila, 6))
                cboCiudad = (.Cells(Fila, 7))
                
            End If
        Next
    
    End With
    
    
    
    
End Sub

Private Sub chkPCotizacion_Change()
    If Me.chkPCotizacion = False Then
        Me.cboIva.Enabled = False
    Else
        Me.cboIva.Enabled = True
    End If
End Sub



Private Sub cboFormaDePago_Change()
    CboDias.Clear
    CboDias.Enabled = True
    cboInteres.Enabled = True
    
    lbl30Dias.Visible = True
    lblHasta30Dias.Visible = True
    txtFecha30Dias.Visible = True
    txtValor30Dias.Visible = True
    
    lbl60Dias.Visible = True
    lblHasta60Dias.Visible = True
    txtFecha60Dias.Visible = True
    txtValor60Dias.Visible = True
    
    lbl90Dias.Visible = True
    lblHasta90Dias.Visible = True
    txtFecha90Dias.Visible = True
    txtValor90Dias.Visible = True
    
    If cboFormaDePago <> "CREDITO" Then
        CboDias.Enabled = False
        cboInteres.Enabled = False
        
        lbl30Dias.Visible = False
        lblHasta30Dias.Visible = False
        txtFecha30Dias.Visible = False
        txtValor30Dias.Visible = False
    
        lbl60Dias.Visible = False
        lblHasta60Dias.Visible = False
        txtFecha60Dias.Visible = False
        txtValor60Dias.Visible = False
        
        lbl90Dias.Visible = False
        lblHasta90Dias.Visible = False
        txtFecha90Dias.Visible = False
        txtValor90Dias.Visible = False
        
    Else
        CboDias.Enabled = True
        cboInteres.Enabled = True
        For i = 30 To 90 Step 30
            CboDias.AddItem i
        Next i
    End If
    
End Sub


Private Sub cboProveedor_Change()
    Dim Fila As Long
    Dim Final As Long
    
    Me.cboProducto.Clear
    Me.cboColor.Clear
    Me.txtCantidad = Empty
    Me.txtMedida = Empty
    Me.txtValorUnitario = Empty
    Me.txtDisponible = Empty
    Me.txtStock = Empty
    Me.txtPedir = Empty
    
  
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
    
    Me.cboColor.Clear
    Me.txtCantidad = Empty
    Me.txtMedida = Empty
    Me.txtValorUnitario = Empty
    Me.txtDisponible = Empty
    Me.txtStock = Empty
    Me.txtPedir = Empty
    
  
    With Hoja2 ' productos
                    
        Final = GetUltimoR(Hoja2)
    
        For Fila = 2 To Final
            If .Cells(Fila, 3) = Me.cboProducto Then
                 Agregar cboColor, .Cells(Fila, 4)
                 'txtValorUnitario = .Cells(Fila, 10)
                 
            End If
        Next
    
    End With

End Sub

Private Sub cboColor_Change()
    Dim Fila As Long
    Dim Final As Long
    
    Dim nomIva As String
    Dim valIva As Double
    
    Me.txtMedida = ""
    Me.txtValorUnitario = ""
    Me.txtValorEmpaque = ""
    'Me.cboIva = 0
    Me.txtValorUnitarioIva = ""
    Me.txtValorEmpaqueIva = ""
    Me.txtUnidadesSolicitadas = ""
    Me.txtCantidadSolicitada = ""
    Me.txtUnidades = ""
    Me.txtUnidadesSolicitadas = ""
    Me.txtCantidad = ""
    Me.txtDisponible = ""
    Me.txtStock = ""
    Me.txtPedir = ""
    Me.txtSubtotal = ""
    
    'MsgBox (chkPCotizacion)
    
  
    With Hoja2 ' productos
                    
        Final = GetUltimoR(Hoja2)
    
        For Fila = 2 To Final
            If .Cells(Fila, 17) = Me.cboProveedor And .Cells(Fila, 3) = Me.cboProducto And .Cells(Fila, 4) = Me.cboColor Then
                
                'verificamos si se van a liquidar productos con iva
                If Me.chkPCotizacion = False Then
                    Me.cboIva = 0
                Else
                    Me.cboIva = .Cells(Fila, 11)
                End If
    
                'se convierte formato de procentaje a decimales para cálculos
                nomIva = Me.cboIva
    
                Select Case nomIva
                    Case Is = 0
                        valIva = 0#
                    Case Is = "0,0%"
                        valIva = 0#
                    Case Is = "1,0%"
                        valIva = 0.01
                    Case Is = "1,5%"
                        valIva = 0.015
                    Case Is = "2,0%"
                        valIva = 0.02
                    Case Is = "2,5%"
                        valIva = 0.025
                    Case Is = "3,0%"
                        valIva = 0.03
                    Case Is = "3,5%"
                        valIva = 0.035
                    Case Is = "4,0%"
                        valIva = 0.04
                    Case Is = "4,5%"
                        valIva = 0.045
                    Case Is = "5,0%"
                        valIva = 0.05
                    Case Is = "5,5%"
                        valIva = 0.055
                    Case Is = "6,0%"
                        valIva = 0.06
                End Select
                
                'traigo campos de la tabla productos según el producto seleccionado
                
                Me.txtMedida = .Cells(Fila, 5)
                'cantidad por unidad de empaque (metros z rollo, unidades x caja)
                
                Me.txtCantidad = .Cells(Fila, 6)
                Me.txtCantidadSolicitada = .Cells(Fila, 6)
 
                Me.txtValorUnitario = .Cells(Fila, 10)
                
                Me.txtValorEmpaque = txtValorUnitario * txtCantidad
                Me.cboColor = .Cells(Fila, 4)
                
                'MsgBox (valIva)
                 
                txtValorUnitarioIva = Application.WorksheetFunction.RoundUp(Me.txtValorUnitario * (1 + valIva), 0)
                txtValorEmpaqueIva = Application.WorksheetFunction.RoundUp(Me.txtValorEmpaque * (1 + valIva), 0)
                 
                Me.txtDisponible = .Cells(Fila, 14)
                Me.txtStock = .Cells(Fila, 15)
                Me.txtPedir = .Cells(Fila, 16)
            End If
        Next
    
    End With
    

    
End Sub


'Private Sub btnFechaElaboracion_Click()
'banderaCalendario = 1
'    Call LanzarCalendario(Me, "txtFechaElaboracion")
'End Sub

Private Sub btnFechaEntrega_Click()
banderaCalendario = 2
    Call LanzarCalendario(Me, "txtFechaEntrega")
End Sub




