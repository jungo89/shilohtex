option explicit

private pId as integer
private pIdCliente as integer
private pTelefono as string
private pDireccion as string
private pCorreo as string
private pBarrio as string
private pCiudad as string

Public Property let Id (Value as integer)
    pId = Value
End Property

Public Property get Id() as integer
    Id = pId
End Property

Public Property let IdCliente (Value as integer)
    pIdCliente = Value
End Property

Public Property get IdCliente() as integer
    IdCliente = pIdCliente
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

Public Property let Correo (Value as string)
    pCorreo = Value
End Property

Public Property get Correo() as string
    Correo = pCorreo
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

Public Sub RegistrarContactoCliente()
'fila vacia por debajo del último registro del listado
    Hoja5.Select 'contacto_cliente
    Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Select
    ActiveCell.value = Me.Id
    ActiveCell.Offset(0, 1).value = Me.IdCliente
    ActiveCell.Offset(0, 2).value = Me.Telefono
    ActiveCell.Offset(0, 3).value = Me.Direccion
    ActiveCell.Offset(0, 4).value = Me.Correo
    ActiveCell.Offset(0, 5).value = Me.Barrio
    ActiveCell.Offset(0, 6).value = Me.Ciudad
End Sub