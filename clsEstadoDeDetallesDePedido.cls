option explicit

private pIdDeSituacion as integer
private pNombreDelEstado as string

Public Property let IdDeSituacion (Value as integer)
    pIdDeSituacion = Value
End Property

Public Property get IdDeSituacion() as integer
    IdDeSituacion = pIdDeSituacion
End Property

Public Property let NombreDelEstado (Value as string)
    pNombreDelEstado = Value
End Property

Public Property get NombreDelEstado() as string
    NombreDelEstado = pNombreDelEstado
End Property

Public Sub RegistrarEstadoDeDetallesDePedido()
'fila vacia por debajo del último registro del listado
    Hoja10.Select 'estado_de_detalles_de_pedidos
    Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Select
    ActiveCell.value = Me.IdDeSituacion
    ActiveCell.Offset(0, 1).value = Me.NombreDelEstado

End Sub