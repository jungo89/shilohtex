VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmConsultarCodigo_3 
   Caption         =   "Consultar Código"
   ClientHeight    =   7680
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6795
   OleObjectBlob   =   "frmConsultarCodigo_3.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmConsultarCodigo_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'inicializar controles del formulario al cargar
'----------------------------------------------------------------------------------------------

Private Sub UserForm_Initialize()
    Dim Fila As Integer
    Dim Final As Integer
 
        
    With Hoja6 'contacto_proveedor
       
    Final = GetUltimoR(Hoja6)

        For Fila = 2 To Final
            If .Cells(Fila, 3) <> "" Then
                Me.cboProveedor.AddItem (.Cells(Fila, 3))
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

Private Sub txtCantidad_Change()
    Me.txtCantidad.BackColor = &HFFFFFF
   
    On Error Resume Next
    Me.txtCantidad.Value = FormatNumber(Me.txtCantidad.Value, 0)
    
End Sub


Private Sub txtCosto_Change()
    Me.txtCosto.BackColor = &HFFFFFF
   
    On Error Resume Next
    Me.txtCosto.Value = FormatCurrency(Me.txtCosto.Value, 2)
    
End Sub

Private Sub txtUtilidad_Change()
    Me.txtUtilidad.BackColor = &HFFFFFF
    
    On Error Resume Next
    Me.txtUtilidad.Value = FormatPercent(Me.txtUtilidad.Value, 1)
    
End Sub

Private Sub txtVenta_Change()
    Me.txtVenta.BackColor = &HFFFFFF
   
    On Error Resume Next
    Me.txtVenta.Value = FormatCurrency(Me.txtVenta.Value, 2)
    
End Sub

Private Sub txtIva_Change()
    Me.txtIva.BackColor = &HFFFFFF
    
    On Error Resume Next
    Me.txtIva.Value = FormatPercent(Me.txtIva.Value, 1)
    
End Sub

Private Sub txtVentaIva_Change()
    Me.txtVentaIva.BackColor = &HFFFFFF
   
    On Error Resume Next
    Me.txtVentaIva.Value = FormatCurrency(Me.txtVentaIva.Value, 2)
    
End Sub
