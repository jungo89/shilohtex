option explicit

private pId as integer
private pNombreEstadoImpuestos as string

Public Property let Id (Value as integer)
    pId = Value
End Property

Public Property get Id() as integer
    Id = pId
End Property

Public Property let NombreEstadoImpuestos (Value as string)
    pNombreEstadoImpuestos = Value
End Property

Public Property get NombreEstadoImpuestos() as string
    NombreEstadoImpuestos = pNombreEstadoImpuestos
End Property

Public Sub RegistrarEstadoDeImpuestosDePedidos()
'fila vacia por debajo del último registro del listado
    Hoja11.select 'estado_de_impuestos_de_pedidos
    Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Select
    ActiveCell.value = Me.Id
    ActiveCell.Offset(0, 1).value = Me.NombreEstadoImpuestos
End Sub