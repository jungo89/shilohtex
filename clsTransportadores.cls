private pId as integer
private pEmpresa as string
private pNombreContacto as string
private pCargo as string
private pDireccion as string
private pTelefono as string
private pCorreo as string
private pCiudad as string

Public Property let Id (Value as integer)
    pId = Value
End Property

Public Property get Id() as integer
    Id = pId
End Property

Public Property let Empresa (Value as string)
    pEmpresa  = Value
End Property

Public Property get Empresa() as string
    Empresa  = pEmpresa 
End Property

Public Property let NombreContacto (Value as string)
    pNombreContacto = Value
End Property

Public Property get NombreContacto() as string
    NombreContacto = pNombreContacto
End Property

Public Property let Cargo (Value as string)
    pCargo = Value
End Property

Public Property get Cargo() as string
    Cargo = pCargo
End Property

Public Property let Direccion (Value as string)
    pDireccion = Value
End Property

Public Property get Direccion() as string
    Direccion = pDireccion
End Property

Public Property let Telefono (Value as string)
    pTelefono = Value
End Property

Public Property get Telefono() as string
    Telefono = pTelefono
End Property

Public Property let Correo (Value as string)
    pCorreo = Value
End Property

Public Property get Correo() as string
    Correo = pCorreo
End Property

Public Property let Ciudad (Value as string)
    pCiudad = Value
End Property

Public Property get Ciudad() as string
    Ciudad = pCiudad
End Property

Public Sub RegistrarTransportadores()
'fila vacia por debajo del último registro del listado
    Hoja19.Select   'transportadores
    Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Select
    ActiveCell.value = Me.Id
    ActiveCell.Offset(0, 1).value = Me.Id
    ActiveCell.Offset(0, 2).value = Me.Empresa
    ActiveCell.Offset(0, 3).value = Me.NombreContacto
    ActiveCell.Offset(0, 4).value = Me.Cargo
    ActiveCell.Offset(0, 5).value = Me.Direccion
    ActiveCell.Offset(0, 6).value = Me.Telefono
    ActiveCell.Offset(0, 7).value = Me.Correo
    ActiveCell.Offset(0, 8).value = Me.Ciudad
End Sub