VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_EliminarUsuario 
   Caption         =   "Eliminar Usuarios"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3975
   OleObjectBlob   =   "frm_EliminarUsuario.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_EliminarUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ComboBox1_Change()
Dim Fila As Long
Dim Final As Long

If ComboBox1.Value = "" Then
    Me.txt_Status = ""
End If

Final = GetUltimoR(Hoja6)

    For Fila = 2 To Final
        If ComboBox1 = Hoja6.Cells(Fila, 1) Then
            Me.txt_Status = Hoja6.Cells(Fila, 3)
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

Private Sub CommandButton1_Click()
    Dim Fila As Long
    Dim Final As Long

On Error GoTo Salir

Final = GetUltimoR(Hoja6)

If Me.ComboBox1 = Empty Then
    MsgBox "Debe seleccionar un usuario!", vbInformation
    Exit Sub
End If

If MsgBox("¿Seguro que quiere eliminar este Usuario?", vbQuestion + vbYesNo) = vbYes Then
    
                For Fila = 2 To Final
                    If Me.ComboBox1 = Hoja6.Cells(Fila, 1) Then
                        Hoja6.Cells(Fila, 1).EntireRow.Delete
                        Exit For
                    End If
                Next
               
                MsgBox "Usuario eliminado", vbInformation + vbOKOnly
                ComboBox1_Enter
    Else
            Exit Sub
    
End If
    


    
    Me.ComboBox1 = ""
    Me.txt_Status = ""

Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Inventarios"
 End If

End Sub
Private Sub CommandButton2_Click()
Unload Me
End Sub
