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
  
   
End Sub

Private Sub cboNombreContacto_Change()
    Dim Fila As Long
    Dim Final As Long
    Dim idProveedor As Integer
    
    'Me.txtRazonSocial = Empty
    'Me.txtDocumento = Empty
    'Me.txtTipoDocumento = Empty
    Me.txtFormaPago = Empty
    Me.txtCelular = Empty
    Me.txtTelefono = Empty
    Me.txtCorreo = Empty
    Me.txtDireccion = Empty
    Me.txtBarrio = Empty
    Me.txtCiudad = Empty
    
  
    With Hoja6 ' contacto_proveedor
                    
        Final = GetUltimoR(Hoja6)
    
        For Fila = 2 To Final
            If .Cells(Fila, 3) = cboNombreContacto Then
                Me.txtCelular = .Cells(Fila, 4)
                Me.txtTelefono = .Cells(Fila, 5)
                Me.txtCorreo = .Cells(Fila, 7)
                Me.txtDireccion = .Cells(Fila, 6)
                Me.txtBarrio = .Cells(Fila, 8)
                Me.txtCiudad = .Cells(Fila, 9)
                
                idProveedor = .Cells(Fila, 2)
                
            End If
                        
        Next
    
    End With
    
    'MsgBox (idProveedor)
    
    With Hoja4 ' proveedores
                    
        Final = GetUltimoR(Hoja4)
    
        For Fila = 2 To Final
            If .Cells(Fila, 1) = idProveedor Then
                'Me.txtRazonSocial = .Cells(Fila, 4)
                'Me.txtDocumento = .Cells(Fila, 3)
                Me.cboTipoContribuyente = .Cells(Fila, 6)
                Me.txtFormaPago = .Cells(Fila, 5)
                
            End If
                        
        Next
    
    End With
    
    
End Sub
