sub 1
    Private Sub cboNombreContacto_Change()
        Dim Fila As Long
        Dim Final As Long

        cboTelefono.Clear
        'cboTelefono.SetFocus
            
        Final = GetUltimoR(Hoja1)
        
            For Fila = 2 To Final
                If Hoja5.Cells(Fila, 7) = cboNombreContacto.Text Then
                    Me.cboTelefono.Value = Hoja5.Cells(Fila, 3)
                    Me.cboDireccion.Value = Hoja5.Cells(Fila, 4)
                    Me.cboBarrio.Value = Hoja5.Cells(Fila, 5)
                    Me.cboCiudad.Value = Hoja5.Cells(Fila, 6)
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
end sub



'Segunda versión com combobox relacionados


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


