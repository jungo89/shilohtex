option explicit

private pId as integer
private pUbicacion as string

Public Property let Id (Value as integer)
    pId = Value
End Property

Public Property get Id() as integer
    Id = pId
End Property

Public Property let Ubicacion (Value as string)
    pUbicacion = Value
End Property

Public Property get Ubicacion() as string
    Ubicacion = pUbicacion
End Property

Public Sub RegistrarUbicacionProducto()
'fila vacia por debajo del último registro del listado
    Hoja20.Select 'ubicacion_producto
    Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Select
    ActiveCell.value = Me.Id
    ActiveCell.Offset(0, 1).value = Me.Ubicacion
 
End Sub
