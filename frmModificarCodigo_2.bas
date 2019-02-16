'inicializar controles del formulario al cargar
'----------------------------------------------------------------------------------------------

Private Sub UserForm_Initialize()
    Dim Fila As Integer
    Dim Final As Integer
 
        
    With Hoja4 'proveedores
       
    Final = GetUltimoR(Hoja4)

        For Fila = 2 To Final
            If .Cells(Fila, 2) <> "" Then
                Me.cboProveedor.AddItem (.Cells(Fila, 2))
            End If
        Next

    End With
 
End Sub


'Eventos de controles
'----------------------------------------------------------------------------------------------


Private Sub cboProveedor_Change()
    Dim Fila As Long
    Dim Final As Long
    
    Me.cboProducto.Clear
    Me.cboColor.Clear
    Me.txtCategoria = Empty
    Me.txtPresentacion = Empty
    Me.txtCantidad = Empty
    Me.txtMedida = Empty
    Me.txtCosto = Empty
    Me.txtUtilidad = Empty
    Me.txtVenta = Empty
    Me.txtIva = Empty
    Me.txtVentaIva = Empty
    
  
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
    Me.txtCategoria = Empty
    Me.txtPresentacion = Empty
    Me.txtCantidad = Empty
    Me.txtMedida = Empty
    Me.txtCosto = Empty
    Me.txtUtilidad = Empty
    Me.txtVenta = Empty
    Me.txtIva = Empty
    Me.txtVentaIva = Empty
    
  
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
    
    Me.txtCategoria = Empty
    Me.txtPresentacion = Empty
    Me.txtCantidad = Empty
    Me.txtMedida = Empty
    Me.txtCosto = Empty
    Me.txtUtilidad = Empty
    Me.txtVenta = Empty
    Me.txtIva = Empty
    Me.txtVentaIva = Empty

    With Hoja2 ' productos
                    
        Final = GetUltimoR(Hoja2)
    
        For Fila = 2 To Final
            If .Cells(Fila, 17) = Me.cboProveedor And .Cells(Fila, 3) = Me.cboProducto And .Cells(Fila, 4) = Me.cboColor Then
                Me.txtCategoria = .Cells(Fila, 13)
                Me.txtPresentacion = .Cells(Fila, 7)
                Me.txtCantidad = .Cells(Fila, 6)
                Me.txtMedida = .Cells(Fila, 5)
                Me.txtCosto = .Cells(Fila, 8)
                Me.txtUtilidad = .Cells(Fila, 9)
                Me.txtVenta = .Cells(Fila, 10)
                Me.txtIva = .Cells(Fila, 11)
                Me.txtVentaIva = .Cells(Fila, 12)
            End If
        Next
    
    End With
    

    
End Sub


Private Sub txtProducto_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtMedida_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

'Validar entradas para permitir ingreso de sólo caracteres o números dependiendo del tipo de campo

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

'aceptar sólo números incluida coma para decimales

Private Sub txtCosto_Change()
    Me.txtCosto.BackColor = &HFFFFFF
    
    If Me.txtCosto = "" Then
        Me.txtVenta = Empty
        Me.txtVentaIva = Empty
        Me.txtUtilidad = Empty
        Me.txtIva = Empty
    End If
    
    If Me.txtCosto <> "" And Me.txtUtilidad <> "" Then
        Me.txtVenta = Application.WorksheetFunction.RoundUp(Me.txtCosto * (1 + (Me.txtUtilidad / 100)), 0)
    Else
        Me.txtVenta = Empty
    End If
    
    If Me.txtVenta <> "" And Me.txtIva <> "" Then
        Me.txtVentaIva = Application.WorksheetFunction.RoundUp(Me.txtVenta * (1 + (Me.txtIva / 100)), 0)
    Else
        Me.txtVentaIva = Empty
    End If
    
    On Error Resume Next
    Me.txtCosto.Value = FormatCurrency(Me.txtCosto.Value, 2)
    
End Sub

Private Sub txtCosto_Exit(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next
    Me.txtCosto.Value = FormatCurrency(Me.txtCosto.Value, 2)
End Sub

Private Sub txtCosto_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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

Private Sub txtUtilidad_Change()
    Me.txtCosto.BackColor = &HFFFFFF
    
    If Me.txtCosto <> "" And Me.txtUtilidad <> "" Then
        Me.txtVenta = Application.WorksheetFunction.RoundUp(Me.txtCosto * (1 + (Me.txtUtilidad / 100)), 0)
    Else
        Me.txtVenta = Empty
    End If
    
End Sub

Private Sub txtUtilidad_Exit(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next
    Me.txtUtilidad.Value = FormatNumber(Me.txtUtilidad.Value, 2)
End Sub

Private Sub txtUtilidad_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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

Private Sub txtVenta_Change()
    Me.txtVenta.BackColor = &HFFFFFF
    
    On Error Resume Next
    Me.txtVenta.Value = FormatCurrency(Me.txtVenta.Value, 2)
End Sub

Private Sub txtVenta_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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

Private Sub txtIva_Change()
    Me.txtIva.BackColor = &HFFFFFF
    
    If Me.txtVenta <> "" And Me.txtIva <> "" Then
        Me.txtVentaIva = Application.WorksheetFunction.RoundUp(Me.txtVenta * (1 + (Me.txtIva / 100)), 0)
    Else
        Me.txtVentaIva = Empty
    End If
End Sub

Private Sub txtIva_Exit(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next
    Me.txtIva.Value = FormatNumber(Me.txtIva.Value, 2)
End Sub

Private Sub txtIva_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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

Private Sub txtVentaIva_Change()
    Me.txtVentaIva.BackColor = &HFFFFFF
    
    On Error Resume Next
    Me.txtVentaIva.Value = FormatCurrency(Me.txtVentaIva.Value, 2)
End Sub

Private Sub txtVentaIva_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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

