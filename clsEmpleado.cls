
option explicit

private pId as integer
private pNombre as String
private pCargo as String
private pTelefonoEmpresa as string

Public Property let Id (Value as integer)
    pId = Value
End Property

Public Property get Id() as integer
    Id = pId
End Property

Public Property let Nombre (Value as string)
    pNombre = Value
End Property

Public Property get Nombre() as string
    Nombre = pNombre
End Property

Public Property let Cargo (Value as string)
    pCargo = Value
End Property

Public Property get Cargo() as string
    Cargo = pCargo
End Property

Public Property let TelefonoEmpresa (Value as string)
    pTelefonoEmpresa = Value
End Property

Public Property get TelefonoEmpresa() as string
    TelefonoEmpresa = pTelefonoEmpresa
End Property

Public Sub RegistrarEmpleado()
'fila vacia por debajo del último registro del listado
    Hoja9.Select 'empleados
    Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Select
    ActiveCell.value = Me.Id
    ActiveCell.Offset(0, 1).value = Me.Nombre
    ActiveCell.Offset(0, 2).value = Me.Cargo
    ActiveCell.Offset(0, 3).value = Me.TelefonoEmpresa
End Sub