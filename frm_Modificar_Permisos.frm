VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Modificar_Permisos 
   Caption         =   "Modificar Permisos"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5160
   OleObjectBlob   =   "frm_Modificar_Permisos.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Modificar_Permisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()
Dim Fila As Long
Dim Final As Long

If ComboBox1.Text = "" Then
    Me.txt_Status = ""
End If

Final = GetUltimoR(Hoja6)

    For Fila = 2 To Final
        If ComboBox1.Text = Hoja6.Cells(Fila, 1) Then
            Me.txt_Status = Hoja6.Cells(Fila, 3)
            
            ' Valor FALSO o VERDADERO para las hojas de cálculo
                Me.CheckBox1.Value = Hoja6.Cells(Fila, 4)
                Me.CheckBox2.Value = Hoja6.Cells(Fila, 5)
                Me.CheckBox3.Value = Hoja6.Cells(Fila, 6)
                Me.CheckBox4.Value = Hoja6.Cells(Fila, 7)
                Me.CheckBox5.Value = Hoja6.Cells(Fila, 8)
                Me.CheckBox6.Value = Hoja6.Cells(Fila, 9)
                Me.CheckBox7.Value = Hoja6.Cells(Fila, 10)
                Me.CheckBox8.Value = Hoja6.Cells(Fila, 11)
                Me.CheckBox9.Value = Hoja6.Cells(Fila, 12)
                Me.CheckBox10.Value = Hoja6.Cells(Fila, 13)
                Me.CheckBox11.Value = Hoja6.Cells(Fila, 14)
                Me.CheckBox12.Value = Hoja6.Cells(Fila, 15)
                
            ' Valor FALSO o VERDADERO para los botones de la Ribbon
                Me.CheckBox13.Value = Hoja6.Cells(Fila, 16)
                Me.CheckBox14.Value = Hoja6.Cells(Fila, 17)
                Me.CheckBox15.Value = Hoja6.Cells(Fila, 18)
                Me.CheckBox16.Value = Hoja6.Cells(Fila, 19)
                Me.CheckBox17.Value = Hoja6.Cells(Fila, 20)
                Me.CheckBox18.Value = Hoja6.Cells(Fila, 21)
                Me.CheckBox19.Value = Hoja6.Cells(Fila, 22)
                Me.CheckBox20.Value = Hoja6.Cells(Fila, 23)
                Me.CheckBox21.Value = Hoja6.Cells(Fila, 24)
                Me.CheckBox22.Value = Hoja6.Cells(Fila, 25)
                Me.CheckBox23.Value = Hoja6.Cells(Fila, 26)
                Me.CheckBox24.Value = Hoja6.Cells(Fila, 27)
                Me.CheckBox25.Value = Hoja6.Cells(Fila, 28)
                Me.CheckBox26.Value = Hoja6.Cells(Fila, 29)
                Me.CheckBox27.Value = Hoja6.Cells(Fila, 30)
                Me.CheckBox28.Value = Hoja6.Cells(Fila, 31)
                Me.CheckBox29.Value = Hoja6.Cells(Fila, 32)
                Me.CheckBox30.Value = Hoja6.Cells(Fila, 33)
                
                
                
                
            Exit For
        End If
    Next
    

End Sub
Private Sub ComboBox1_Enter()
Dim Fila As Long
Dim Final As Long
Dim Lista As String



For Fila = 1 To ComboBox1.ListCount
    ComboBox1.RemoveItem 0
Next Fila

Final = GetUltimoR(Hoja6)
    
    For Fila = 2 To Final
        Lista = Hoja6.Cells(Fila, 1)
        ComboBox1.AddItem (Lista)
    Next
End Sub

Private Sub cmd_Guardar_Click()
Dim Fila As Long
Dim Final As Long

On Error GoTo Salir

If Me.ComboBox1.Text = Empty Then
    MsgBox "Debe seleccionar un usuario", vbInformation
    Exit Sub
End If


    Final = GetUltimoR(Hoja6)
    
    
    For Fila = 2 To Final
        If Me.ComboBox1.Text = Hoja6.Cells(Fila, 1) Then
        
        ' Valor FALSO o VERDADERO para las hojas de cálculo
            Hoja6.Cells(Fila, 4) = Me.CheckBox1.Value
            Hoja6.Cells(Fila, 5) = Me.CheckBox2.Value
            Hoja6.Cells(Fila, 6) = Me.CheckBox3.Value
            Hoja6.Cells(Fila, 7) = Me.CheckBox4.Value
            Hoja6.Cells(Fila, 8) = Me.CheckBox5.Value
            Hoja6.Cells(Fila, 9) = Me.CheckBox6.Value
            Hoja6.Cells(Fila, 10) = Me.CheckBox7.Value
            Hoja6.Cells(Fila, 11) = Me.CheckBox8.Value
            Hoja6.Cells(Fila, 12) = Me.CheckBox9.Value
            Hoja6.Cells(Fila, 13) = Me.CheckBox10.Value
            Hoja6.Cells(Fila, 14) = Me.CheckBox11.Value
            Hoja6.Cells(Fila, 15) = Me.CheckBox12.Value
            
         ' Valor FALSO o VERDADERO para los botones de la Ribbon
            Hoja6.Cells(Fila, 16) = Me.CheckBox13.Value
            Hoja6.Cells(Fila, 17) = Me.CheckBox14.Value
            Hoja6.Cells(Fila, 18) = Me.CheckBox15.Value
            Hoja6.Cells(Fila, 19) = Me.CheckBox16.Value
            Hoja6.Cells(Fila, 20) = Me.CheckBox17.Value
            Hoja6.Cells(Fila, 21) = Me.CheckBox18.Value
            Hoja6.Cells(Fila, 22) = Me.CheckBox19.Value
            Hoja6.Cells(Fila, 23) = Me.CheckBox20.Value
            Hoja6.Cells(Fila, 24) = Me.CheckBox21.Value
            Hoja6.Cells(Fila, 25) = Me.CheckBox22.Value
            Hoja6.Cells(Fila, 26) = Me.CheckBox23.Value
            Hoja6.Cells(Fila, 27) = Me.CheckBox24.Value
            Hoja6.Cells(Fila, 28) = Me.CheckBox25.Value
            Hoja6.Cells(Fila, 29) = Me.CheckBox26.Value
            Hoja6.Cells(Fila, 30) = Me.CheckBox27.Value
            Hoja6.Cells(Fila, 31) = Me.CheckBox28.Value
            Hoja6.Cells(Fila, 32) = Me.CheckBox29.Value
            Hoja6.Cells(Fila, 33) = Me.CheckBox30.Value
            
            Exit For

        End If
    Next
ThisWorkbook.Save
MsgBox "Cambios guardados satisfactoriamente", vbInformation, "Gestor de Inventarios"
    
    Unload Me
    Hoja1.txt_UsuarioActual.Text = Empty
    frm_Login.Show

Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Inventarios"
 End If
 
End Sub

Private Sub cmd_Salir_Click()
    Unload Me
End Sub

