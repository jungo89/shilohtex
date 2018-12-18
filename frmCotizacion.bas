

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
    
End Sub

Private Sub cboNombreContacto_Change()
    Dim Fila As Long
    Dim Final As Long
    
    txtRazonSocial = Empty
    cboTelefono.Clear
    cboDireccion.Clear
    cboBarrio.Clear
    cboCiudad.Clear
    
    With Hoja1 ' clientes
                    
        Final = GetUltimoR(Hoja1)
    
        For Fila = 2 To Final
            If .Cells(Fila, 4) = cboNombreContacto Then
                txtRazonSocial.Text = (.Cells(Fila, 6))
            End If
        Next
    
    End With
    
    With Hoja5 ' datos_cliente
                    
        Final = GetUltimoR(Hoja5)
    
        For Fila = 2 To Final
            If .Cells(Fila, 7) = cboNombreContacto Then
                cboTelefono.AddItem (.Cells(Fila, 3))
                cboDireccion.AddItem (.Cells(Fila, 4))
                cboBarrio.AddItem (.Cells(Fila, 5))
                cboCiudad.AddItem (.Cells(Fila, 6))
            End If
        Next
    
    End With

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


