VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmContactoCliente 
   Caption         =   "Datos Básicos del Cliente"
   ClientHeight    =   3630
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7710
   OleObjectBlob   =   "frmContactoCliente.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmContactoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


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

