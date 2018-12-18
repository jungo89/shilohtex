'Mensaje de confimación de cambios
iRep = Msgbox("Desea guardar los cambios?", _
vbYesNo + vbQuestion + vbDefaultButton2, _
"Confirmación")


'Función para determinar la última fila con datos de una hoja
Public Function GetUltimoR(Hoja As Worksheet) As Integer
    GetUltimoR = GetNuevoR(Hoja) - 1
End Function

Public Function GetNuevoR(Hoja As Worksheet) As Integer
    
    Dim Fila As Long
    Fila = 2
    
    Do While Hoja.Cells(Fila, 1) <> ""
        Fila = Fila + 1
    Loop
    
    GetNuevoR = Fila
    
End Function


'Cargar datos a combobox para eventos change  y enter
Private Sub cboNombreContacto_Change()
Dim Fila As Long
Dim Final As Long
Dim Registro As Integer

    
Final = GetUltimoR(Hoja1)

    For Fila = 2 To Final
        If cboNombreContacto.Text = Hoja1.Cells(Fila, 4) Then
            Me.cboNombreContacto.Text = Hoja1.Cells(Fila, 4)
            Exit For
        
        End If
    Next


End Sub

Private Sub cboNombreContacto_Enter()
Dim Fila As Long
Dim Final As Long
Dim Lista As String


For Fila = 1 To cboNombreContacto.ListCount
    cboNombreContacto.RemoveItem 0
Next Fila

Final = GetUltimoR(Hoja1)
    
    For Fila = 2 To Final
        Lista = Hoja1.Cells(Fila, 4)
        cboNombreContacto.AddItem (Lista)
    Next
End Sub




Private Sub ComboBox1_Change()

ComboBox2.Clear
ComboBox2.SetFocus

UF = Sheets("Equipo").Range("C" & Rows.Count).End(xlUp).Row

For i = 10 To UF
    If Sheets("Equipo").Cells(i, "C") = ComboBox1 Then
        AddItem ComboBox2, Sheets("Equipo").Cells(i, "D")
    End If
Next

End Sub


Private Sub ComboBox2_Change()

ComboBox3.Clear
ComboBox3.SetFocus

UF = Sheets("Equipo").Range("C" & Rows.Count).End(xlUp).Row

For i = 10 To UF
    If Sheets("Equipo").Cells(i, "C") = ComboBox1 And _
        Sheets("Equipo").Cells(i, "D") = ComboBox2 Then
            AddItem ComboBox3, Sheets("Equipo").Cells(i, "E")
    End If
Next

End Sub


Private Sub ComboBox3_Change()

ComboBox4.Clear
ComboBox4.SetFocus

UF = Sheets("Equipo").Range("C" & Rows.Count).End(xlUp).Row

For i = 10 To UF
    If Sheets("Equipo").Cells(i, "C") = ComboBox1 And _
        Sheets("Equipo").Cells(i, "E") = ComboBox3 Then
            AddItem ComboBox4, Sheets("Equipo").Cells(i, "F")
    End If
Next

End Sub


Rem PARA ACTUALIZAR LOS COMBOBOX

Private Sub CmdActualizar_Click()

ComboBox1.Clear

UF = Sheets("Equipo").Range("C" & Rows.Count).End(xlUp).Row

For i = 10 To UF
    AddItem ComboBox1, Sheets("Equipo").Cells(i, "C")
Next

End Sub

Rem FIN


Sub AddItem(cmbBox As ComboBox, sItem As String)

'Agrega los item únicos y en orden alfabético

For i = 0 To cmbBox.ListCount - 1
    Select Case StrComp(cmbBox.List(i), sItem, vbTextCompare)
        Case 0: Exit Sub 'ya existe en el combo y ya no lo agrega
        Case 1: cmbBox.AddItem sItem, i: Exit Sub 'Es menor, lo agrega antes del comparado
    End Select
Next

cmbBox.AddItem sItem 'Es mayor lo agrega al final

End Sub


Private Sub ComboBox4_Change()

TextBox7.Value = ComboBox4.Value
TextBox2.SetFocus

End Sub


Private Sub UserForm_Activate()

ThisWorkbook.Activate
ComboBox1.Clear

UF = Sheets("Equipo").Range("C" & Rows.Count).End(xlUp).Row

For i = 10 To UF
    AddItem ComboBox1, Sheets("Equipo").Cells(i, "C")
Next

End Sub







Private Sub CmdCancelar_Click()

End

End Sub


'Cargar datos a combobox para eventos change  y enter
Private Sub cboTelefono_Change()
Dim Fila As Long
Dim Final As Long

Dim buscar as String
buscarv = cboNombreContacto
    
Final = GetUltimoR(Hoja1)

    For Fila = 2 To Final
        If cboTelefono.Text = Hoja1.Cells(Fila, 4) Then
            Me.cboTelefono.Text = Hoja1.Cells(Fila, 4)
            Exit For
        
        End If
    Next


End Sub

Private Sub cboTelefono_Enter()
Dim Fila As Long
Dim Final As Long
Dim Lista As String


For Fila = 1 To cboTelefono.ListCount
    cboTelefono.RemoveItem 0
Next Fila

Final = GetUltimoR(Hoja1)
    
    For Fila = 2 To Final
        Lista = Hoja1.Cells(Fila, 4)
        cboTelefono.AddItem (Lista)
    Next
End Sub

Option Explicit





sub combos_relacionados()
    Dim id_tienda As Integer
    Dim id_depto As Integer
    Dim id_empleado As Integer

    Private Sub cbo_Tiendas_Change()
        Dim Fila As Integer
        Dim Final As Integer


        cbo_Departamentos.Clear

        With Hoja3 ' Tabla Departamentos
                
                id_tienda = cbo_Tiendas.ListIndex + 1
                
            Final = .Cells(1, 1).End(xlDown).Row

            For Fila = 2 To Final
                If Mid(.Cells(Fila, 1), 1, 1) = id_tienda Then
                    cbo_Departamentos.AddItem (.Cells(Fila, 2))
                End If
            Next

        End With

    End Sub

    Private Sub cbo_Departamentos_Change()
        Dim Fila As Integer
        Dim Final As Integer

        cbo_Empleados.Clear
        Call LimpiarTextBoxes

        With Hoja4 'Tabla Empleados

                id_depto = cbo_Departamentos.ListIndex + 1
                

            Final = .Cells(1, 1).End(xlDown).Row

            For Fila = 2 To Final
                If Mid(.Cells(Fila, 1), 1, 1) = id_tienda And _
                Mid(.Cells(Fila, 1), 2, 1) = id_depto Then
                
                    cbo_Empleados.AddItem (.Cells(Fila, 2))
                    
                End If
            Next

        End With
    End Sub

    Private Sub cbo_Empleados_Change()
        Dim Fila As Integer
        Dim Final As Integer

        Call LimpiarTextBoxes

        With Hoja4 'Tabla Empleados

                id_empleado = cbo_Empleados.ListIndex + 1
                

            Final = .Cells(1, 1).End(xlDown).Row

            For Fila = 2 To Final
                If Mid(.Cells(Fila, 1), 1, 1) = id_tienda And _
                Mid(.Cells(Fila, 1), 2, 1) = id_depto And _
                Mid(.Cells(Fila, 1), 3, 1) = id_empleado Then
                
                    txtNombre = .Cells(Fila, 2)
                    txtFechaing = .Cells(Fila, 3)
                    txtPuesto = .Cells(Fila, 4)
                    txtEmail = .Cells(Fila, 5)
                    
                End If
            Next

        End With
    End Sub

    Private Sub UserForm_Initialize()
        Dim Fila As Integer
        Dim Final As Integer

        With Hoja2 'Tabla Tiendas

        Final = .Cells(1, 1).End(xlDown).Row

            For Fila = 2 To Final
                If .Cells(Fila, 1) <> "" Then
                    cbo_Tiendas.AddItem (.Cells(Fila, 2))
                End If
            Next

        End With
    End Sub

    Sub LimpiarTextBoxes()
        txtNombre = Empty
        txtFechaing = Empty
        txtPuesto = Empty
        txtEmail = Empty
    End Sub

    Private Sub CommandButton1_Click()
        Unload Me
    End Sub
end sub
