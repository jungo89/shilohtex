'Convertir entrada de campos texto a mayúsculas


Private Sub txtNombre_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCargo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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


Private Sub txtNombre_AfterUpdate()
'Determina el final del listado de empleados
        Final = GetNuevoR(Hoja9)
        
        'Validación para impedir Empleados repetidos
        For Fila = 2 To Final
            If Me.txtNombre.Text <> "" And UCase(Hoja9.Cells(Fila, 2)) = UCase(Me.txtNombre.Text) Then
                MsgBox ("Empleado ya existe en la Base de Datos"), , Titulo
                LimpiarControles
                Me.txtNombre.SetFocus
                Exit Sub
                Exit For
            End If
        Next
End Sub


Private Sub UserForm_Initialize()

'poblar combo Cargo
    Me.cboCargo.AddItem "ASESORA COMERCIAL"
    Me.cboCargo.AddItem "AUXILIAR DE BODEGA"
    Me.cboCargo.AddItem "ANALISTA CONTABLE"
    Me.cboCargo.AddItem "SUPERVISOR"

    
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
    
    Titulo = "Empleados"
    
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
    
    
        'crear recordset empleados
        Set Rs = New ADODB.Recordset
        Rs.CursorLocation = adUseServer
        Rs.Open Source:="empleados", _
            ActiveConnection:=Conn, _
            CursorType:=adOpenDynamic, _
            LockType:=adLockOptimistic, _
            Options:=adCmdTable
    
    
        'Cargar los datos a tabla empleados
        With Rs
            .AddNew
            .Fields("nombre") = txtNombre
            .Fields("cargo") = cboCargo
            .Fields("telefono_empresa") = txtTelefono
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
    Dim xTextBox As Control
        
                
        For Each xTextBox In Controls
            If xTextBox.Name Like "txt*" Then
                xTextBox = Empty
                Me.txtNombre.SetFocus
            End If
        Next
        
        Me.cboCargo = Empty
        
End Sub




