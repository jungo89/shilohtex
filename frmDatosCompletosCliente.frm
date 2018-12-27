VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDatosCompletosCliente 
   Caption         =   "Clientes"
   ClientHeight    =   7395
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16065
   OleObjectBlob   =   "frmDatosCompletosCliente.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmDatosCompletosCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdGuardar_Click()

    Dim Conn As ADODB.Connection
    Dim MiConexion
    Dim Rs As ADODB.Recordset
    Dim MiBase As String
    
    MiBase = "cotizador.accdb"

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
        .Fields("cupo") = txtCupo
        .Fields("credito") = txtCredito
        .Fields("saldo") = txtSaldo
        .Fields("categoria") = txtSaldo


    End With

    Rs.Update
    
    'limpiar recordset
    Rs.Delete

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
        .Fields("id_cliente") = 1108
        .Fields("telefono") = CDbl(txtTelefono)
        .Fields("direccion") = txtDireccion
        .Fields("barrio") = txtBarrio
        .Fields("ciudad") = txtCiudad



    End With


    'Cerrar la conexión
    Rs.Update
    Rs.Close

    Conn.Close
    Set Rs = Nothing
    Set Conn = Nothing

    MsgBox "Alta exitosa", vbInformation, "EXCELeINFO"
    'Me.Unload
End Sub


