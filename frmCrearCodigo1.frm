VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCrearCodigo1 
   Caption         =   "Crear Código"
   ClientHeight    =   5865
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10110
   OleObjectBlob   =   "frmCrearCodigo1.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmCrearCodigo1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Convertir entrada de campos texto a mayúsculas

Private Sub txtProducto_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtMedida_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

'Validar entradas para permitir ingreso de sólo caracteres o números dependiendo del tipo de campo

'aceptar sólo números
Private Sub txtCantidad_Change()
    Dim Texto As Variant
    Dim Caracter As Variant
    Dim Largo As Integer
    
    On Error Resume Next
    
    Texto = Me.txtCantidad.Value
    Largo = Len(Me.txtCantidad.Value)
    For i = 1 To Largo
        Caracter = Mid(Texto, i, 1)
        If Caracter <> "" Then
            If Caracter < Chr(48) Or Caracter > Chr(57) Then
                Me.txtCantidad.Value = Replace(Texto, Caracter, "")
            Else
            End If
        End If
    Next i
    On Error GoTo 0
    Caracter = 0
    Caracter1 = 0
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
    
End Sub

Private Sub txtCosto_Exit(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next
    Me.txtCosto.Value = FormatNumber(Me.txtCosto.Value, 2)
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
    Me.txtVenta.Value = FormatNumber(Me.txtVenta.Value, 2)
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
    Me.txtVentaIva.Value = FormatNumber(Me.txtVentaIva.Value, 2)
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

'aceptar sólo números
Private Sub txtDisponible_change()
    Dim Texto As Variant
    Dim Caracter As Variant
    Dim Largo As Integer
    
    On Error Resume Next
    
    Texto = Me.txtDisponible.Value
    Largo = Len(Me.txtDisponible.Value)
    For i = 1 To Largo
        Caracter = Mid(Texto, i, 1)
        If Caracter <> "" Then
            If Caracter < Chr(48) Or Caracter > Chr(57) Then
                Me.txtDisponible.Value = Replace(Texto, Caracter, "")
            Else
            End If
        End If
    Next i
    On Error GoTo 0
    Caracter = 0
    Caracter1 = 0
End Sub

'aceptar sólo números
Private Sub txtStock_change()
    Dim Texto As Variant
    Dim Caracter As Variant
    Dim Largo As Integer
    
    On Error Resume Next
    
    Texto = Me.txtStock.Value
    Largo = Len(Me.txtStock.Value)
    For i = 1 To Largo
        Caracter = Mid(Texto, i, 1)
        If Caracter <> "" Then
            If Caracter < Chr(48) Or Caracter > Chr(57) Then
                Me.txtStock.Value = Replace(Texto, Caracter, "")
            Else
            End If
        End If
    Next i
    On Error GoTo 0
    Caracter = 0
    Caracter1 = 0
End Sub

'aceptar sólo números
Private Sub txtPedir_change()
    Dim Texto As Variant
    Dim Caracter As Variant
    Dim Largo As Integer
    
    On Error Resume Next
    
    Texto = Me.txtPedir.Value
    Largo = Len(Me.txtPedir.Value)
    For i = 1 To Largo
        Caracter = Mid(Texto, i, 1)
        If Caracter <> "" Then
            If Caracter < Chr(48) Or Caracter > Chr(57) Then
                Me.txtPedir.Value = Replace(Texto, Caracter, "")
            Else
            End If
        End If
    Next i
    On Error GoTo 0
    Caracter = 0
    Caracter1 = 0
End Sub

'GENERAR NVA VALIDACION PARA EVITAR QUE SE DUPLIQUEN PRODUCTOS
'VALIDAR LLAVE PROVEEDOR, PRODUCTO, COLOR, MEDIDA

'Private Sub txtNombreContacto_AfterUpdate()
''Determina el final del listado de Clientes
'        Final = GetNuevoR(Hoja1)
'
'        'Validación para impedir Clientes repetidos
'        For Fila = 2 To Final
'            If Me.txtNombreContacto.Text <> "" And UCase(Hoja1.Cells(Fila, 4)) = UCase(Me.txtNombreContacto.Text) Then
'                MsgBox ("Cliente ya existe en la Base de Datos"), , Titulo
'                LimpiarControles
'                Me.txtNombreContacto.SetFocus
'                Exit Sub
'                Exit For
'            End If
'        Next
'End Sub


Private Sub UserForm_Initialize()

'Poblar combo proveedor
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
    
'Poblar combo color
    With Hoja24 'colores
       
    Final = GetUltimoR(Hoja24)

        For Fila = 2 To Final
            If .Cells(Fila, 2) <> "" Then
                Me.cboColor.AddItem (.Cells(Fila, 2))
            End If
        Next

    End With
    
'Poblar combo medida
    With Hoja25 'medidas
       
    Final = GetUltimoR(Hoja25)

        For Fila = 2 To Final
            If .Cells(Fila, 2) <> "" Then
                Me.cboMedida.AddItem (.Cells(Fila, 2))
            End If
        Next

    End With
    
'poblar combo Presentación
    Me.cboPresentacion.AddItem "BULTO"
    Me.cboPresentacion.AddItem "CAJA"
    Me.cboPresentacion.AddItem "PACA"
    Me.cboPresentacion.AddItem "ROLLO"

'poblar combo Categoría
    Me.cboCategoria.AddItem "CREMALLERAS Y CIERRES"
    Me.cboCategoria.AddItem "ESPUMAS"
    Me.cboCategoria.AddItem "HERRAJES"
    Me.cboCategoria.AddItem "HILOS"
    Me.cboCategoria.AddItem "MALLAS"
    Me.cboCategoria.AddItem "OTROS"
    Me.cboCategoria.AddItem "REATAS Y RIBETES"
    Me.cboCategoria.AddItem "SERVICIOS"
    Me.cboCategoria.AddItem "TELAS"
End Sub


Private Sub cmdGuardar_Click()

    Dim Conn As ADODB.Connection
    Dim MiConexion
    Dim Rs As ADODB.Recordset
    Dim MiBase As String
    Dim Query As String
    Dim Titulo As String
    Dim xTextBox As control
        
    On Error GoTo Salir
    
    Titulo = "Productos"
    
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
    
    
        'crear recordset productos
        Set Rs = New ADODB.Recordset
        Rs.CursorLocation = adUseServer
        Rs.Open Source:="productos", _
            ActiveConnection:=Conn, _
            CursorType:=adOpenDynamic, _
            LockType:=adLockOptimistic, _
            Options:=adCmdTable
    
    
   'determinar el codigo del proveedor
        
    Dim Fila As Integer
    Dim Final As Integer
    Dim tmp As Integer

       With Hoja4 'proveedores
          
       Final = GetUltimoR(Hoja4)
    
           For Fila = 2 To Final
               If .Cells(Fila, 2) = Me.cboProveedor Then
                   tmp = .Cells(Fila, 1).Value
                   Exit For
               End If
           Next
    
       End With
       
       'MsgBox (tmp)
        
      'Cargar los datos a tabla clientes
        With Rs
            .AddNew
            .Fields("id_proveedor") = tmp
            .Fields("producto") = txtProducto
            .Fields("color") = cboColor
            .Fields("medida") = cboMedida
            .Fields("cantidad") = txtCantidad
            .Fields("presentacion") = cboPresentacion
            .Fields("costo") = CCur(txtCosto)
            .Fields("utilidad") = CDbl(txtUtilidad / 100)
            .Fields("venta") = CCur(txtVenta)
            .Fields("iva") = CDbl(txtIva / 100)
            .Fields("venta_iva") = CCur(txtVentaIva)
            .Fields("categoria") = cboCategoria
            .Fields("disponible") = CInt(txtDisponible)
            .Fields("stock") = CInt(txtStock)
            .Fields("pedir") = CInt(txtPedir)
        End With
    
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
    Dim xTextBox As control
        
        For Each xTextBox In Controls
            If xTextBox.Name Like "txt*" Then
                xTextBox = Empty
                Me.txtProducto.SetFocus
            End If
        Next
        
        Me.cboProveedor = Empty
        Me.cboColor = Empty
        Me.cboMedida = Empty
        Me.cboPresentacion = Empty
        Me.cboCategoria = Empty

End Sub






