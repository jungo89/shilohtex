VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmModificarCodigo_2 
   Caption         =   "Modificar Código"
   ClientHeight    =   8205.001
   ClientLeft      =   720
   ClientTop       =   2460
   ClientWidth     =   6630
   OleObjectBlob   =   "frmModificarCodigo_2.frx":0000
End
Attribute VB_Name = "frmModificarCodigo_2"
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
            If .Cells(Fila, 3) = Me.cboProducto And .Cells(Fila, 17) = Me.cboProveedor Then
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
            If .Cells(Fila, 17) = Me.cboProveedor And .Cells(Fila, 3) = Me.cboProducto And .Cells(Fila, 4) = Me.cboColor _
            Then
                Me.txtCategoria = .Cells(Fila, 13)
                Me.txtPresentacion = .Cells(Fila, 7)
                Me.txtCantidad = .Cells(Fila, 6)
                Me.txtMedida = .Cells(Fila, 5)
                Me.txtCosto = .Cells(Fila, 8)
                Me.txtUtilidad = (.Cells(Fila, 9) * 100)
                Me.txtVenta = .Cells(Fila, 10)
                Me.txtIva = (.Cells(Fila, 11) * 100)
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

Private Sub txtCosto_Change()
    Me.txtCosto.BackColor = &HFFFFFF
    
    If Me.txtCosto = "" Then
        Me.txtVenta = Empty
        Me.txtVentaIva = Empty
        Me.txtUtilidad = Empty
        Me.txtIva = Empty
    End If
    
    On Error GoTo Formato
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
    
    'Me.txtCosto.Value = FormatCurrency(Me.txtCosto.Value, 2)
    
Formato:
     If Err <> 0 Then
        'MsgBox Err.Description, vbExclamation, "Error de digitación"
        MsgBox "Verifique el valor digitado", vbExclamation, "Error de digitación"
     End If
    

    
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

Private Sub txtUtilidad_Change()
    Me.txtCosto.BackColor = &HFFFFFF
    
    'On Error Resume Next
    On Error GoTo Formato
    If Me.txtCosto <> "" And Me.txtUtilidad <> "" Then
        Me.txtVenta = Application.WorksheetFunction.RoundUp(Me.txtCosto * (1 + (Me.txtUtilidad / 100)), 0)
    Else
        Me.txtVenta = Empty
    End If
    
Formato:
     If Err <> 0 Then
        'MsgBox Err.Description, vbExclamation, "Error de digitación"
        MsgBox "Verifique el valor digitado", vbExclamation, "Error de digitación"
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
    
    On Error GoTo Formato
    If Me.txtVenta <> "" And Me.txtIva <> "" Then
        Me.txtVentaIva = Application.WorksheetFunction.RoundUp(Me.txtVenta * (1 + (Me.txtIva / 100)), 0)
    Else
        Me.txtVentaIva = Empty
    End If
    
Formato:
     If Err <> 0 Then
        'MsgBox Err.Description, vbExclamation, "Error de digitación"
        MsgBox "Verifique el valor digitado", vbExclamation, "Error de digitación"
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

Private Sub cmdGuardar_Click()

    Dim Fila As Integer
    Dim Final As Integer
    Dim id As Integer
    Dim costo As Currency
    Dim utilidad As Double
    Dim venta As Currency
    Dim iva As Double
    Dim ventaIva As Currency
    
    
    Dim Conn As ADODB.Connection
    Dim MiConexion
    'Dim Rs As ADODB.Recordset
    Dim MiBase As String
    Dim Query As String
    Dim Titulo As String
    
    
    With Hoja2 'producto
      
        Final = GetUltimoR(Hoja2)
    
        For Fila = 2 To Final
            If Hoja2.Cells(Fila, 17) = (Me.cboProveedor.Text) And _
                Hoja2.Cells(Fila, 3) = (Me.cboProducto.Text) And _
                Hoja2.Cells(Fila, 4) = (Me.cboColor.Text) _
            Then
                
                
                id = CInt(.Cells(Fila, 1).Value)
                
                costo = CCur(Me.txtCosto)
                utilidad = (Me.txtUtilidad / 100)
                venta = CCur(Me.txtVenta)
                iva = (Me.txtIva / 100)
                ventaIva = CCur(Me.txtVentaIva)
               
             
                Exit For
            End If
        Next
    
    End With
       
    
    On Error GoTo Salir

    Titulo = "Productos"

    If MsgBox("Son correctos los datos?" + Chr(13) + "Desea proceder?", vbOKCancel, Titulo) = vbOK Then


        MiBase = "cotizador.accdb"

        Set Conn = New ADODB.Connection
        MiConexion = Application.ThisWorkbook.Path & Application.PathSeparator & MiBase

        With Conn
            .Provider = "Microsoft.ACE.OLEDB.12.0"
            .Open MiConexion
        End With
              
       Query = "UPDATE productos SET costo = '" & costo & "', utilidad = '" & utilidad & "', venta = '" & venta & "', iva = '" & iva & "', venta_iva = '" & ventaIva & "' WHERE id = " & id & ""
       
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
                Me.cboProveedor.SetFocus
            End If
        Next

        For Each xComboBox In Controls
            If xComboBox.Name Like "cbo*" Then
                xComboBox = Empty
                Me.cboProveedor.SetFocus
            End If
        Next
        
End Sub




