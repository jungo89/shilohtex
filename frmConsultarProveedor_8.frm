VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmConsultarProveedor_8 
   Caption         =   "Consultar Proveedor"
   ClientHeight    =   5970
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7965
   OleObjectBlob   =   "frmConsultarProveedor_8.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmConsultarProveedor_8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'inicializar controles del formulario al cargar
'----------------------------------------------------------------------------------------------

Private Sub UserForm_Initialize()
    Dim Fila As Integer
    Dim Final As Integer
 
        
    With Hoja6 'contacto_proveedor
       
    Final = GetUltimoR(Hoja6)

        For Fila = 2 To Final
            If .Cells(Fila, 3) <> "" Then
                Me.cboNombreContacto.AddItem (.Cells(Fila, 3))
            End If
        Next

    End With
  
   
End Sub

Private Sub cboNombreContacto_Change()
    Dim Fila As Long
    Dim Final As Long
    Dim idProveedor As Integer
    
    'Me.txtRazonSocial = Empty
    'Me.txtDocumento = Empty
    'Me.txtTipoDocumento = Empty
    Me.txtFormaPago = Empty
    Me.txtCelular = Empty
    Me.txtTelefono = Empty
    Me.txtCorreo = Empty
    Me.txtDireccion = Empty
    Me.txtBarrio = Empty
    Me.txtCiudad = Empty
    
  
    With Hoja6 ' contacto_proveedor
                    
        Final = GetUltimoR(Hoja6)
    
        For Fila = 2 To Final
            If .Cells(Fila, 3) = cboNombreContacto Then
                Me.txtCelular = .Cells(Fila, 4)
                Me.txtTelefono = .Cells(Fila, 5)
                Me.txtCorreo = .Cells(Fila, 7)
                Me.txtDireccion = .Cells(Fila, 6)
                Me.txtBarrio = .Cells(Fila, 8)
                Me.txtCiudad = .Cells(Fila, 9)
                
                idProveedor = .Cells(Fila, 2)
                
            End If
                        
        Next
    
    End With
    
    'MsgBox (idProveedor)
    
    With Hoja4 ' proveedores
                    
        Final = GetUltimoR(Hoja4)
    
        For Fila = 2 To Final
            If .Cells(Fila, 1) = idProveedor Then
                'Me.txtRazonSocial = .Cells(Fila, 4)
                'Me.txtDocumento = .Cells(Fila, 3)
                Me.cboTipoContribuyente = .Cells(Fila, 6)
                Me.txtFormaPago = .Cells(Fila, 5)
                
            End If
                        
        Next
    
    End With
    
    
End Sub
