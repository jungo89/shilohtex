option explicit

private pId as integer
private pIdproveedor as integer
private pTelefono as string
private pDireccion as string
private pBarrio as string
private pCiudad as string

Public Property let Id (Value as integer)
    pId = Value
End Property

Public Property get Id() as integer
    Id = pId
End Property

Public Property let IdProveedor (Value as integer)
    pIdProveedor = Value
End Property

Public Property get IdProveedor() as integer
    IdProveedor = pIdProveedor
End Property

Public Property let Telefono (Value as string)
    pTelefono = Value
End Property

Public Property get Telefono() as string
    Telefono = pTelefono
End Property

Public Property let Direccion (Value as string)
    pDireccion = Value
End Property

Public Property get Direccion() as string
    Direccion = pDireccion
End Property

Public Property let Barrio (Value as string)
    pBarrio = Value
End Property

Public Property get Barrio() as string
    Barrio = pBarrio
End Property

Public Property let Ciudad (Value as string)
    pCiudad = Value
End Property

Public Property get Ciudad() as string
    Ciudad = pCiudad
End Property

Public Sub RegistrarContactoProveedor()
'fila vacia por debajo del último registro del listado
    Hoja6.select 'contacto_proveedor
    Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Select
    ActiveCell.value = Me.Id
    ActiveCell.Offset(0, 1).value = Me.IdProveedor
    ActiveCell.Offset(0, 2).value = Me.Telefono
    ActiveCell.Offset(0, 3).value = Me.Direccion
    ActiveCell.Offset(0, 4).value = Me.Barrio
    ActiveCell.Offset(0, 5).value = Me.Ciudad
End Sub