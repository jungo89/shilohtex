'Pruebas para grabación de datos completos del cliente
Private Sub cmdGuardar_Click()

    'Pendiente implementar la validación de los campos currency
    'Pendiente validar si existe conexión a la base de datos
    'verificar la implementación de las excepciones
    'Falta limpiar los controles después de la grabación

        Dim Conn As ADODB.Connection
        Dim MiConexion
        Dim Rs As ADODB.Recordset
        Dim MiBase As String
        Dim Query As String
        Dim Titulo As String
        Dim xTextBox As Control
            
        On Error GoTo Salir
        
        Titulo = "Clientes"
        
        'Extender validación de campos a combobox y checkbox
            For Each xTextBox In Controls
                If xTextBox.Name Like "txt*" And xTextBox = Empty Then
                    MsgBox "Debe completar todos los campos", , Titulo
                    xTextBox.SetFocus
                    Exit Sub
                End If
            Next
            
        
        MiBase = "pruebas.accdb"

        Set Conn = New ADODB.Connection
        MiConexion = Application.ThisWorkbook.Path & Application.PathSeparator & MiBase

        With Conn
            .Provider = "Microsoft.ACE.OLEDB.12.0"
            .Open MiConexion
        End With


        'crear recordset clientes
        Set Rs = New ADODB.Recordset
        Rs.CursorLocation = adUseServer
        Rs.Open Source:="clientes", _
            ActiveConnection:=Conn, _
            CursorType:=adOpenDynamic, _
            LockType:=adLockOptimistic, _
            Options:=adCmdTable


        'Cargar los datos a tabla clientes
        With Rs
            .AddNew
            .Fields("nombre_contacto") = txtNombreContacto
            .Fields("tipo_documento") = txtTipoDocumento
            .Fields("documento") = txtDocumento
            .Fields("razon_social") = txtRazonSocial
            .Fields("comercio") = txtComercio
            .Fields("nicho") = txtNicho
            .Fields("segmentacion") = txtSegmentacion
            .Fields("producto") = txtProducto
            .Fields("distribucion") = txtDistribucion
            .Fields("cupo") = CCur(txtCupo)
            .Fields("credito") = CCur(txtCredito)
            .Fields("saldo") = CCur(txtSaldo)
            .Fields("categoria") = txtCategoria
        End With

        Rs.Update
        Rs.Close

        'determinar el id del registro que se graba
        'Revisar, debe compararse con la coincidencia exacta. (Existen clientes que se crean con sólo un nombre)
        Query = "SELECT id FROM clientes WHERE nombre_contacto LIKE '%" & Me.txtNombreContacto.Value & "%'"
        'Query = "SELECT id FROM clientes WHERE nombre_contacto = '"& Me.txtNombreContacto.Value &"'"

        Set Rs = New ADODB.Recordset
        Rs.CursorLocation = adUseServer
        Rs.Open Source:=Query, _
        ActiveConnection:=Conn

        'guardar el id en una hoja auxiliar
        'Verificar si es mejor guardarlo en una variable. No olvidar el cast
        Sheets("contadores").Range("A1").CurrentRegion.Clear
        Sheets("contadores").Range("A2").CopyFromRecordset Rs

        Rs.Close

        'crear recordset contacto cliente
        Set Rs = New ADODB.Recordset
        Rs.CursorLocation = adUseServer
        Rs.Open Source:="contacto_cliente", _
            ActiveConnection:=Conn, _
            CursorType:=adOpenDynamic, _
            LockType:=adLockOptimistic, _
            Options:=adCmdTable


        'Cargar los datos a tabla datos_cliente
        With Rs
            .AddNew
            .Fields("id_cliente") = Sheets("contadores").Range("A2").Value
            .Fields("telefono") = txtTelefono
            .Fields("direccion") = txtDireccion
            .Fields("correo") = txtCorreo
            .Fields("barrio") = txtBarrio
            .Fields("ciudad") = txtCiudad
        End With

        'Cerrar la conexión
        Rs.Update
        Rs.Close


        Conn.Close
        Set Rs = Nothing
        Set Conn = Nothing

        MsgBox "Alta exitosa", vbInformation
        
        
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
                    Me.txtNombreContacto.SetFocus
                End If
            Next

End Sub


Private Sub UserForm_Initialize()
    ' Dim Fila As Integer
    ' Dim Final As Integer
 
    ' With Hoja1 'clientes
       
    ' Final = GetUltimoR(Hoja1)

    '     For Fila = 2 To Final
    '         If .Cells(Fila, 4) <> "" Then
    '             Me.cboNombreContacto.AddItem (.Cells(Fila, 4))
    '         End If
    '     Next

    ' End With
    
    ' With Hoja4 'proveedores
       
    ' Final = GetUltimoR(Hoja4)

    '     For Fila = 2 To Final
    '         If .Cells(Fila, 2) <> "" Then
    '             Me.cboProveedor.AddItem (.Cells(Fila, 2))
    '         End If
    '     Next

    ' End With


   
End Sub


'Determina el final del listado de productos
        Final = GetNuevoR(Hoja2)
        
        'Validación para impedir registros repetidos
        For Fila = 2 To Final
            If Hoja2.Cells(Fila, 1) Like Me.txt_CodProd.Text Then
                Me.txt_CodProd.BackColor = &H8080FF
                MsgBox ("Registro ya existe" + Chr(13) + "Ingrese un código diferente")
                Me.txt_CodProd.SetFocus
                Exit Sub
            End If
        Next


'Determina el final del listado de Clientes
        Final = GetNuevoR(Hoja9)
        
        'Validación para impedir Clientes repetidos
        For Fila = 2 To Final
            If Hoja9.Cells(Fila, 1) = UCase(Me.txt_Cliente.Text) Then
                MsgBox ("Cliente ya existe en la Base de Datos"), , Titulo
                LimpiarControles
                Me.txt_Cliente.SetFocus
                Exit Sub
                Exit For
            End If
        Next



'Validar que sólo se ingresen números (a-z y símbolos)
Private Sub txtNumero_Change()
    Dim Texto As Variant
    Dim Caracter As Variant
    Dim Largo As Integer
    On Error Resume Next
    Texto = Me.txtNumero.Value
    Largo = Len(Me.txtNumero.Value)
    For i = 1 To Largo
        Caracter = Mid(Texto, i, 1)
        If Caracter <> "" Then
            If Caracter < Chr(48) Or Caracter > Chr(57) Then
                Me.txtNumero.Value = Replace(Texto, Caracter, "")
            Else
            End If
        End If
    Next i
    On Error GoTo 0
    Caracter = 0
    Caracter1 = 0
End Sub
'
'Validar que sólo se ingrese texto (0-9)
Private Sub txtTexto_Change()
    Dim Texto As Variant
    Dim Caracter As Variant
    Dim Largo As String
    On Error Resume Next
    Texto = Me.txtTexto.Value
    Largo = Len(Me.txtTexto.Value)
    For i = 1 To Largo
        Caracter = CInt(Mid(Texto, i, 1))
        If Caracter <> "" Then
            If Not Application.WorksheetFunction.IsText(Caracter) Then
                Me.txtTexto.Value = Replace(Texto, Caracter, "")
            Else
            End If
        End If
    Next i
    On Error GoTo 0
End Sub        

On Error GoTo Salir

Salir:
     If Err <> 0 Then
        MsgBox Err.Description, vbExclamation, Titulo
     End If