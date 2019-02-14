VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRegistrarProveedor_6 
   Caption         =   "Registrar Proveedor"
   ClientHeight    =   6870
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7590
   OleObjectBlob   =   "frmRegistrarProveedor_6.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmRegistrarProveedor_6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Convertir entrada de campos texto a mayúsculas


Private Sub txtNombre_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtRazonSocial_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtDireccion_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtBarrio_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub



'Validar entradas para permitir ingreso de sólo caracteres o números dependiendo del tipo de campo

'aceptar sólo números
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


Private Sub txtNombre_AfterUpdate()
'Determina el final del listado de proveedores
        Final = GetNuevoR(Hoja4)
        
        'Validación para impedir Clientes repetidos
        For Fila = 2 To Final
            If Me.txtNombre.Text <> "" And UCase(Hoja4.Cells(Fila, 2)) = UCase(Me.txtNombre.Text) Then
                MsgBox ("Proveedor ya existe en la Base de Datos"), , Titulo
                LimpiarControles
                Me.txtNombre.SetFocus
                Exit Sub
                Exit For
            End If
        Next
End Sub


Private Sub UserForm_Initialize()

'Poblar combo ciudades
    Dim Fila As Integer
    Dim Final As Integer
 
    With Hoja23 'ciudades
       
    Final = GetUltimoR(Hoja23)

        For Fila = 2 To Final
            If .Cells(Fila, 4) <> "" Then
                Me.cboCiudad.AddItem (.Cells(Fila, 4))
            End If
        Next

    End With
    
    
'poblar combo TipoDocumento
    Me.cboTipoDocumento.AddItem "PERSONA JURIDICA"
    Me.cboTipoDocumento.AddItem "PERSONA NATURAL"
    Me.cboTipoDocumento.AddItem "REGIMEN SIMPLIFICADO"

'poblar combo FormaPago
    Me.cboFormaPago.AddItem "CONTADO"
    Me.cboFormaPago.AddItem "CREDITO"
    
End Sub


Private Sub cmdGuardar_Click()

    Dim Conn As ADODB.Connection
    Dim MiConexion
    Dim Rs As ADODB.Recordset
    Dim MiBase As String
    Dim Query As String
    Dim Titulo As String
    Dim xTextBox As Control
        
    On Error GoTo Salir
    
    Titulo = "Proveedores"
    
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
    
    
        'crear recordset proveedores
        Set Rs = New ADODB.Recordset
        Rs.CursorLocation = adUseServer
        Rs.Open Source:="proveedores", _
            ActiveConnection:=Conn, _
            CursorType:=adOpenDynamic, _
            LockType:=adLockOptimistic, _
            Options:=adCmdTable
    
    
        'Cargar los datos a tabla proveedores
        With Rs
            .AddNew
            .Fields("nombre") = txtNombre
            .Fields("tipo_documento") = cboTipoDocumento
            .Fields("documento") = txtDocumento
            .Fields("razon_social") = txtRazonSocial
            .Fields("nit") = txtNit
            .Fields("forma_pago") = cboFormaPago
        End With
    
        Rs.Update
        Rs.Close
    
        'determinar el id del registro que se graba
        Query = "SELECT id FROM proveedores WHERE nombre LIKE '%" & Me.txtNombre.Value & "%'"
        'Query = "SELECT id FROM proveedores WHERE nombre = '" & Me.txtNombre.Value & "'"
    
        Set Rs = New ADODB.Recordset
        Rs.CursorLocation = adUseServer
        Rs.Open Source:=Query, _
        ActiveConnection:=Conn
    
        Sheets("contadores").Range("B2").Select
        Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
        Selection.ClearContents
        
        Sheets("contadores").Range("B2").CopyFromRecordset Rs
    
        Rs.Close
    
        'crear recordset contacto proveedor
        Set Rs = New ADODB.Recordset
        Rs.CursorLocation = adUseServer
        Rs.Open Source:="contacto_proveedor", _
            ActiveConnection:=Conn, _
            CursorType:=adOpenDynamic, _
            LockType:=adLockOptimistic, _
            Options:=adCmdTable
    
    
        'Cargar los datos a tabla datos_proveedor
        With Rs
            .AddNew
            .Fields("id_proveedor") = Sheets("contadores").Range("B2").Value
            .Fields("telefono") = txtTelefono
            .Fields("direccion") = txtDireccion
            .Fields("correo") = txtCorreo
            .Fields("barrio") = txtBarrio
            .Fields("ciudad") = cboCiudad
        End With
    
        'Cerrar la conexión
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
    Dim xTextBox As Control
        
        For Each xTextBox In Controls
            If xTextBox.Name Like "txt*" Then
                xTextBox = Empty
                Me.txtNombre.SetFocus
            End If
        Next
        
        Me.cboCiudad = Empty
        Me.cboTipoDocumento = Empty
        Me.cboFormaPago = Empty

End Sub


