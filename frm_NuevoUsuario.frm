VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_NuevoUsuario 
   Caption         =   "Registro de Usuarios"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5175
   OleObjectBlob   =   "frm_NuevoUsuario.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_NuevoUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_Registrar_Click()
    Dim Fila As Long
    Dim Final As Long
    Dim Registro As Integer
    
On Error GoTo Salir
        
Final = GetNuevoR(Hoja6)
        
        For Registro = 2 To Final
            If Hoja6.Cells(Registro, 1) = Me.txt_nUser.Text Then
                Me.txt_nUser.BackColor = &H8080FF
                MsgBox ("El usuario ya existe" + Chr(13) + "Ingrese un usuario diferente")
                Me.txt_nUser.SetFocus
                Exit Sub
                Exit For
            End If
        Next
        
      If Me.txt_pass1.Text = Me.txt_pass2.Text Then
                Me.txt_nUser.BackColor = &HFFFFFF
                Hoja6.Cells(Final, 1) = Me.txt_nUser.Text
                Hoja6.Cells(Final, 2) = Me.txt_pass1.Text
                        If Me.OptionButton1.Value = True Then
                            Hoja6.Cells(Final, 3) = "Usuario"
                                Else
                            Hoja6.Cells(Final, 3) = "Administrador"
                        End If
                        
            ' Valor FALSO o VERDADERO para las hojas de cálculo
                Hoja6.Cells(Final, 4) = Me.CheckBox1.Value
                Hoja6.Cells(Final, 5) = Me.CheckBox2.Value
                Hoja6.Cells(Final, 6) = Me.CheckBox3.Value
                Hoja6.Cells(Final, 7) = Me.CheckBox4.Value
                Hoja6.Cells(Final, 8) = Me.CheckBox5.Value
                Hoja6.Cells(Final, 9) = Me.CheckBox6.Value
                Hoja6.Cells(Final, 10) = Me.CheckBox7.Value
                Hoja6.Cells(Final, 11) = Me.CheckBox8.Value
                Hoja6.Cells(Final, 12) = Me.CheckBox9.Value
                Hoja6.Cells(Final, 13) = Me.CheckBox10.Value
                Hoja6.Cells(Final, 14) = Me.CheckBox11.Value
                Hoja6.Cells(Final, 15) = Me.CheckBox12.Value
               
            ' Valor FALSO o VERDADERO para los botones de la Ribbon
                Hoja6.Cells(Final, 16) = Me.CheckBox13.Value
                Hoja6.Cells(Final, 17) = Me.CheckBox14.Value
                Hoja6.Cells(Final, 18) = Me.CheckBox15.Value
                Hoja6.Cells(Final, 19) = Me.CheckBox16.Value
                Hoja6.Cells(Final, 20) = Me.CheckBox17.Value
                Hoja6.Cells(Final, 21) = Me.CheckBox18.Value
                Hoja6.Cells(Final, 22) = Me.CheckBox19.Value
                Hoja6.Cells(Final, 23) = Me.CheckBox20.Value
                Hoja6.Cells(Final, 24) = Me.CheckBox21.Value
                Hoja6.Cells(Final, 25) = Me.CheckBox22.Value
                Hoja6.Cells(Final, 26) = Me.CheckBox23.Value
                Hoja6.Cells(Final, 27) = Me.CheckBox24.Value
                Hoja6.Cells(Final, 28) = Me.CheckBox25.Value
                Hoja6.Cells(Final, 29) = Me.CheckBox26.Value
                Hoja6.Cells(Final, 30) = Me.CheckBox27.Value
                Hoja6.Cells(Final, 31) = Me.CheckBox28.Value
                Hoja6.Cells(Final, 32) = Me.CheckBox29.Value
                Hoja6.Cells(Final, 33) = Me.CheckBox30.Value
                '-----------------------------------------------
                Me.txt_nUser.Text = ""
                Me.txt_pass1.Text = ""
                Me.txt_pass2.Text = ""
                Me.CheckBox1.Value = False
                Me.CheckBox2.Value = False
                Me.CheckBox3.Value = False
                Me.CheckBox4.Value = False
                Me.CheckBox5.Value = False
                Me.CheckBox6.Value = False
                Me.CheckBox7.Value = False
                Me.CheckBox8.Value = False
                Me.CheckBox9.Value = False
                Me.CheckBox10.Value = False
                Me.CheckBox11.Value = False
                Me.CheckBox12.Value = False
                Me.CheckBox13.Value = False
                Me.CheckBox14.Value = False
                Me.CheckBox15.Value = False
                Me.CheckBox16.Value = False
                Me.CheckBox17.Value = False
                Me.CheckBox18.Value = False
                Me.CheckBox19.Value = False
                Me.CheckBox20.Value = False
                Me.CheckBox21.Value = False
                Me.CheckBox22.Value = False
                Me.CheckBox23.Value = False
                Me.CheckBox24.Value = False
                Me.CheckBox25.Value = False
                Me.CheckBox26.Value = False
                Me.CheckBox27.Value = False
                Me.CheckBox28.Value = False
                Me.CheckBox29.Value = False
                Me.CheckBox30.Value = False
                
                Me.txt_nUser.SetFocus
                ThisWorkbook.Save
                MsgBox "Usuario registrado satisfactoriamente", , "Gestor de Inventarios"
            Else
                MsgBox "Las contraseñas deben coincidir"
    End If
    

Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Inventarios"
 End If
 
End Sub

Private Sub cmd_Finalizar_Click()
    Unload Me
End Sub

