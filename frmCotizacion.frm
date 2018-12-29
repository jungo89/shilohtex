VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCotizacion 
   Caption         =   "Cotización"
   ClientHeight    =   13065
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   28155
   OleObjectBlob   =   "frmCotizacion.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "frmCotizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


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
    cboFormaDePago.Clear
    cboPrioridad.Clear
    CboDias.Clear
    txtCupo.Text = Empty
    txtCredito.Text = Empty
    txtSaldo.Text = Empty
    txtCategoria = Empty
    
    
    With Hoja1 ' clientes
                    
        Final = GetUltimoR(Hoja1)
    
        For Fila = 2 To Final
            If .Cells(Fila, 4) = cboNombreContacto Then
                txtRazonSocial.Text = (.Cells(Fila, 6))
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
                cboBarrio.AddItem (.Cells(Fila, 5))
                cboCiudad.AddItem (.Cells(Fila, 6))
            End If
        Next
    
    End With
    
  
     'combo forma de pago
    Me.cboFormaDePago.AddItem "Contado"
    Me.cboFormaDePago.AddItem "Contra entrega"
    Me.cboFormaDePago.AddItem "Crédito"
    
    'combo prioridad
    Me.cboPrioridad.AddItem "Inmediato"
    Me.cboPrioridad.AddItem "Felipe"
    Me.cboPrioridad.AddItem "LogiFuturo"
   

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
    
    If cboFormaDePago <> "Crédito" Then
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



